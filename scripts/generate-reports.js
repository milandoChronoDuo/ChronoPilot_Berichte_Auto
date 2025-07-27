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

console.log('Skript geladen. ENV gesetzt?', !!SUPABASE_URL, !!SUPABASE_SERVICE_KEY);

if (!SUPABASE_URL || !SUPABASE_SERVICE_KEY) {
  console.error('FEHLER: SUPABASE_URL oder SUPABASE_SERVICE_KEY fehlen!');
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_KEY);

function formatDateDE(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  return d.toLocaleDateString('de-DE');
}

async function getKundenMitVersandTag(heute) {
  console.log('Hole Kunden für Versandtag:', heute.getDate());
  const { data, error } = await supabase
    .from('kunden')
    .select('*')
    .eq('istversand', heute.getDate())
    .eq('status', 'aktiv');
  if (error) throw error;
  return data;
}

async function getMitarbeitende(kunden_id) {
  console.log('Hole Mitarbeitende für Kunde:', kunden_id);
  const { data, error } = await supabase
    .from('mitarbeitende')
    .select('*')
    .eq('kunden_id', kunden_id)
    .is('deleted_at', null);
  if (error) throw error;
  return data;
}

async function getZeiten(kunden_id, mitarbeiter_id, von, bis) {
  console.log(`Hole Tageszeiten für ${mitarbeiter_id} von ${von} bis ${bis}`);
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
  return interval;
}

function sumIntervals(intervals) {
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
  const browser = await puppeteer.launch({
    headless: 'new',
    args: ['--no-sandbox', '--disable-setuid-sandbox'],
  });
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
    { header: 'Status', key: 'tagesstatus', width: 16 },
    { header: 'Start', key: 'erster_start', width: 10 },
    { header: 'Ende', key: 'letzter_ende', width: 10 },
    { header: 'Pause', key: 'gesamt_pause', width: 10 },
    { header: 'Netto', key: 'gesamt_netto', width: 10 },
    { header: 'Über-/Minusstunden', key: 'ueber_unter_stunden', width: 14 }
  ];
  for (const z of zeiten) {
    sheet.addRow({
      datum: formatDateDE(z.datum),
      tagesstatus: z.tagesstatus || '',
      erster_start: z.erster_start ? z.erster_start.substring(11, 16) : '',
      letzter_ende: z.letzter_ende ? z.letzter_ende.substring(11, 16) : '',
      gesamt_pause: intervalToStr(z.gesamt_pause),
      gesamt_netto: intervalToStr(z.gesamt_netto),
      ueber_unter_stunden: intervalToStr(z.ueber_unter_stunden),
    });
  }
  const pauseSum = sumIntervals(zeiten.map(z => z.gesamt_pause));
  const nettoSum = sumIntervals(zeiten.map(z => z.gesamt_netto));
  const ueberSum = sumIntervals(zeiten.map(z => z.ueber_unter_stunden));
  sheet.addRow({});
  sheet.addRow({
    datum: 'Summe',
    gesamt_pause: pauseSum,
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
  console.log('Starte Berichtsexport. Heute:', new Date().toISOString());
  await fs.mkdir('./tmp', { recursive: true });

  const heute = new Date();
  const heuteDatum = heute.toISOString().split('T')[0];
  let kunden;
  try {
    kunden = await getKundenMitVersandTag(heute);
  } catch (err) {
    console.error('Fehler beim Laden der Kunden:', err);
    return;
  }
  console.log('Gefundene Kunden:', kunden.length, kunden.map(k => k.name).join(', '));

  if (!kunden.length) {
    console.log('Keine Firmen für Bericht heute.');
    return;
  }

  let template;
  try {
    template = await fs.readFile(path.resolve('templates/report-template.html'), 'utf8');
  } catch (err) {
    console.error('FEHLER: report-template.html nicht gefunden!', err);
    return;
  }

  for (const kunde of kunden) {
    const { id: kunden_id, name: firma_name, lastversand, erstellungsdatum } = kunde;
    // Zeitraum-Berechnung
    let von;
    if (lastversand) {
      von = new Date(heute.getFullYear(), heute.getMonth() - 1, lastversand + 1);
    } else {
      von = new Date(heute.getFullYear(), heute.getMonth() - 2, 1);
    }
    const bis = new Date(heute);
    bis.setDate(bis.getDate() - 1);
    const zeitraum_start = von.toISOString().split('T')[0];
    const zeitraum_ende = bis.toISOString().split('T')[0];

    let mitarbeitende;
    try {
      mitarbeitende = await getMitarbeitende(kunden_id);
    } catch (err) {
      console.error('Fehler beim Laden der Mitarbeitenden:', err);
      continue;
    }
    console.log(`Bearbeite Kunde: ${firma_name} (${kunden_id})`);
    console.log(`Zeitraum: ${zeitraum_start} bis ${zeitraum_ende}`);
    console.log('Gefundene Mitarbeitende:', mitarbeitende.map(m => m.name).join(', '));

    let berichteErzeugt = false;

    for (const ma of mitarbeitende) {
      const { id: ma_id, name: ma_name } = ma;
      let zeiten;
      try {
        zeiten = await getZeiten(kunden_id, ma_id, zeitraum_start, zeitraum_ende);
      } catch (err) {
        console.error(`Fehler beim Laden der Tageszeiten für ${ma_name}:`, err);
        continue;
      }
      console.log(`Lese tageszeiten für ${ma_name} (${ma_id}) von ${zeitraum_start} bis ${zeitraum_ende}`);
      console.log(`Gefundene tageszeiten:`, zeiten.length);

      if (!zeiten.length) continue;

      // Tabellenzeile bauen
      const tableRows = zeiten.map(z => `
        <tr>
          <td>${formatDateDE(z.datum)}</td>
          <td>${z.tagesstatus || ''}</td>
          <td>${z.erster_start ? z.erster_start.substring(11, 16) : ''}</td>
          <td>${z.letzter_ende ? z.letzter_ende.substring(11, 16) : ''}</td>
          <td>${intervalToStr(z.gesamt_pause)}</td>
          <td>${intervalToStr(z.gesamt_netto)}</td>
          <td>${intervalToStr(z.ueber_unter_stunden)}</td>
        </tr>
      `).join('\n');
      const pauseSum = sumIntervals(zeiten.map(z => z.gesamt_pause));
      const nettoSum = sumIntervals(zeiten.map(z => z.gesamt_netto));
      const ueberSum = sumIntervals(zeiten.map(z => z.ueber_unter_stunden));

      const pdfVars = {
        firma_name,
        mitarbeiter_name: ma_name,
        zeitraum_start: formatDateDE(zeitraum_start),
        zeitraum_ende: formatDateDE(zeitraum_ende),
        table_rows: tableRows,
        summe_pause: pauseSum,
        summe_netto: nettoSum,
        summe_uebermin: ueberSum,
        logo_path: path.resolve('templates/chronoduo.png')
      };
      const baseFile = `${sanitizeFilename(firma_name)}_${heute.getMonth() + 1}_${heute.getFullYear()}_${sanitizeFilename(ma_name)}`;
      const pdfPath = `./tmp/${baseFile}.pdf`;
      const xlsxPath = `./tmp/${baseFile}.xlsx`;

      try {
        await renderPdf(template, pdfVars, pdfPath);
        await renderExcel(zeiten, xlsxPath);
      } catch (err) {
        console.error(`Fehler beim Erstellen von PDF/Excel für ${ma_name}:`, err);
        continue;
      }

      const remotePdfPath = `${kunden_id}/${heute.getFullYear()}_${heute.getMonth() + 1}/${baseFile}.pdf`;
      const remoteXlsxPath = `${kunden_id}/${heute.getFullYear()}_${heute.getMonth() + 1}/${baseFile}.xlsx`;

      try {
        await uploadToBucket(pdfPath, BUCKET, remotePdfPath);
        await uploadToBucket(xlsxPath, BUCKET, remoteXlsxPath);
        console.log(`Bericht für ${firma_name} / ${ma_name} exportiert und hochgeladen.`);
        berichteErzeugt = true;
      } catch (err) {
        console.error(`Fehler beim Upload für ${ma_name}:`, err);
      }
    }

    // lastversand setzen, wenn mind. ein Bericht erzeugt wurde
    if (berichteErzeugt) {
      try {
        await supabase
          .from('kunden')
          .update({ lastversand: heute.getDate() })
          .eq('id', kunden_id);
        console.log(`lastversand für ${firma_name} aktualisiert!`);
      } catch (err) {
        console.error(`Fehler beim Aktualisieren von lastversand für ${firma_name}:`, err);
      }
    }
  }
}

main()
  .then(() => { console.log('Berichtsexport fertig!'); })
  .catch(err => {
    console.error('FEHLER im Skript:', err);
    process.exit(1);
  });
