import { createClient } from '@supabase/supabase-js';
import * as dotenv from 'dotenv';
import puppeteer from 'puppeteer';
import ExcelJS from 'exceljs';
import fs from 'fs/promises';
import path from 'path';
dotenv.config();

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_KEY = process.env.SUPABASE_SERVICE_KEY;
const BUCKET = 'berichte';

if (!SUPABASE_URL || !SUPABASE_SERVICE_KEY) {
  throw new Error('SUPABASE_URL und SUPABASE_SERVICE_KEY müssen gesetzt sein!');
}

const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_KEY);

async function getKundenMitVersandTag(heute) {
  const { data, error } = await supabase
    .from('kunden')
    .select('*')
    .eq('istversand', heute.getDate())
    .eq('status', 'aktiv');

  if (error) throw error;
  return data;
}

async function getMitarbeitende(kunden_id) {
  const { data, error } = await supabase
    .from('mitarbeitende')
    .select('*')
    .eq('kunden_id', kunden_id)
    .is('deleted_at', null);

  if (error) throw error;
  return data;
}

async function getZeiten(kunden_id, mitarbeiter_id, von, bis) {
  const { data, error } = await supabase
    .from('tageszeiten')
    .select('*')
    .eq('kunden_id', kunden_id)
    .eq('mitarbeiter_id', mitarbeiter_id)
    .gte('datum', von)
    .lte('datum', bis)
    .order('datum', { ascending: true });

  if (error) throw error;
  return data;
}

function intervalToStr(interval) {
  if (!interval) return '';
  // PostgreSQL Interval wird von Supabase als z.B. "4:30:00" (4h 30m)
  return interval;
}

function sumIntervals(intervals) {
  // "4:30:00" + "2:45:00" = "7:15:00" etc.
  let totalSeconds = 0;
  for (const i of intervals) {
    if (!i) continue;
    const parts = i.split(':');
    if (parts.length !== 3) continue;
    const [h, m, s] = parts.map(Number);
    totalSeconds += h * 3600 + m * 60 + s;
  }
  const h = Math.floor(totalSeconds / 3600);
  const m = Math.floor((totalSeconds % 3600) / 60);
  const s = totalSeconds % 60;
  return `${h}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
}

async function renderPdf(template, vars, outPath) {
  let html = template;
  Object.entries(vars).forEach(([key, val]) => {
    html = html.replaceAll(`{{${key}}}`, val);
  });

  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: 'networkidle0' });
  await page.pdf({ path: outPath, format: 'A4', printBackground: true });
  await browser.close();
}

async function renderExcel(zeiten, outPath) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Monatsbericht');
  sheet.columns = [
    { header: 'Datum', key: 'datum', width: 12 },
    { header: 'Start', key: 'erster_start', width: 16 },
    { header: 'Ende', key: 'letzter_ende', width: 16 },
    { header: 'Pause', key: 'gesamt_pause', width: 10 },
    { header: 'Netto', key: 'gesamt_netto', width: 10 },
    { header: 'Über-/Minusstunden', key: 'ueber_unter_stunden', width: 14 },
    { header: 'Status', key: 'tagesstatus', width: 12 },
  ];

  for (const z of zeiten) {
    sheet.addRow({
      datum: z.datum,
      erster_start: z.erster_start ? z.erster_start.substring(11, 16) : '',
      letzter_ende: z.letzter_ende ? z.letzter_ende.substring(11, 16) : '',
      gesamt_pause: intervalToStr(z.gesamt_pause),
      gesamt_netto: intervalToStr(z.gesamt_netto),
      ueber_unter_stunden: intervalToStr(z.ueber_unter_stunden),
      tagesstatus: z.tagesstatus || '',
    });
  }

  // Summenzeile
  const nettoSum = sumIntervals(zeiten.map(z => z.gesamt_netto));
  const ueberSum = sumIntervals(zeiten.map(z => z.ueber_unter_stunden));
  sheet.addRow({});
  sheet.addRow({
    datum: 'Summe',
    gesamt_netto: nettoSum,
    ueber_unter_stunden: ueberSum,
  });

  await workbook.xlsx.writeFile(outPath);
}

function sanitizeFilename(str) {
  return str.replace(/[\/\\?%*:|"<>]/g, '').replace(/\s+/g, '_');
}

async function uploadToBucket(localPath, bucket, remotePath) {
  const fileBuffer = await fs.readFile(localPath);
  const { error } = await supabase.storage.from(bucket).upload(remotePath, fileBuffer, {
    upsert: true,
    contentType: localPath.endsWith('.pdf')
      ? 'application/pdf'
      : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  if (error) throw error;
}

async function main() {
  const heute = new Date();
  const heuteDatum = heute.toISOString().split('T')[0];
  const kunden = await getKundenMitVersandTag(heute);

  if (!kunden.length) {
    console.log('Keine Firmen für Bericht heute.');
    return;
  }

  const template = await fs.readFile(path.resolve('templates/report-template.html'), 'utf8');

  for (const kunde of kunden) {
    const { id: kunden_id, name: firma_name, lastversand, erstellungsdatum } = kunde;
    const von =
      lastversand
        ? new Date(heute.getFullYear(), heute.getMonth() - 1, lastversand + 1)
        : new Date(erstellungsdatum);
    const bis = new Date(heute);
    bis.setDate(bis.getDate() - 1);
    const zeitraum_start = von.toISOString().split('T')[0];
    const zeitraum_ende = bis.toISOString().split('T')[0];

    const mitarbeitende = await getMitarbeitende(kunden_id);
    for (const ma of mitarbeitende) {
      const { id: ma_id, name: ma_name } = ma;
      const zeiten = await getZeiten(kunden_id, ma_id, zeitraum_start, zeitraum_ende);
      if (!zeiten.length) continue;

      // Tabellen-HTML bauen
      const tableRows = zeiten.map(z => `
        <tr>
          <td>${z.datum}</td>
          <td>${z.erster_start ? z.erster_start.substring(11, 16) : ''}</td>
          <td>${z.letzter_ende ? z.letzter_ende.substring(11, 16) : ''}</td>
          <td>${intervalToStr(z.gesamt_pause)}</td>
          <td>${intervalToStr(z.gesamt_netto)}</td>
          <td>${intervalToStr(z.ueber_unter_stunden)}</td>
          <td>${z.tagesstatus || ''}</td>
        </tr>
      `).join('\n');
      const nettoSum = sumIntervals(zeiten.map(z => z.gesamt_netto));
      const ueberSum = sumIntervals(zeiten.map(z => z.ueber_unter_stunden));

      // PDF erstellen
      const pdfVars = {
        firma_name,
        mitarbeiter_name: ma_name,
        zeitraum_start,
        zeitraum_ende,
        table_rows: tableRows,
        summe_netto: nettoSum,
        summe_uebermin: ueberSum,
      };
      const baseFile = `${sanitizeFilename(firma_name)}_${heute.getMonth() + 1}_${heute.getFullYear()}_${sanitizeFilename(ma_name)}`;
      const pdfPath = `./tmp/${baseFile}.pdf`;
      const xlsxPath = `./tmp/${baseFile}.xlsx`;
      await fs.mkdir('./tmp', { recursive: true });
      await renderPdf(template, pdfVars, pdfPath);
      await renderExcel(zeiten, xlsxPath);

      // Upload in Storage
      const remotePdfPath = `${kunden_id}/${heute.getFullYear()}_${heute.getMonth() + 1}/${baseFile}.pdf`;
      const remoteXlsxPath = `${kunden_id}/${heute.getFullYear()}_${heute.getMonth() + 1}/${baseFile}.xlsx`;
      await uploadToBucket(pdfPath, BUCKET, remotePdfPath);
      await uploadToBucket(xlsxPath, BUCKET, remoteXlsxPath);

      console.log(`Bericht für ${firma_name} / ${ma_name} exportiert und hochgeladen.`);
    }
  }
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
