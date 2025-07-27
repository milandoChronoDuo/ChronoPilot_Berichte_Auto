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

function formatSignedInterval(str) {
  if (!str) return '';
  // Typografisches Minus für PDFs
  if (str.startsWith('-')) return '−' + str.substring(1);
  return str;
}

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
  return interval;
}

function sumIntervals(intervals) {
  let totalSeconds = 0;
  for (const i of intervals) {
    if (!i) continue;
    const negative = i.startsWith('-');
    const parts = (negative ? i.slice(1) : i).split(':');
    if (parts.length !== 3) continue;
    const [h, m, s] = parts.map(Number);
    let seconds = h * 3600 + m * 60 + s;
    if (negative) seconds = -seconds;
    totalSeconds += seconds;
  }
  const abs = Math.abs(totalSeconds);
  const h = Math.floor(abs / 3600);
  const m = Math.floor((abs % 3600) / 60);
  const s = abs % 60;
  const sign = totalSeconds < 0 ? '-' : '';
  return `${sign}${h}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
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

async function getFeiertage(land, bundesland, von, bis) {
  const { data, error } = await supabase
    .from('feiertage')
    .select('datum')
    .eq('land', land)
    .eq('bundesland', bundesland)
    .gte('datum', von)
    .lte('datum', bis);
  if (error) throw error;
  return (data || []).map(f => f.datum);
}

function getNextVersanddatum(sollversand, heute, feiertage) {
  const jahr = heute.getMonth() === 11 ? heute.getFullYear() + 1 : heute.getFullYear();
  const monat = (heute.getMonth() + 1) % 12;
  let d = new Date(jahr, monat, sollversand);
  if (d <= heute) d = new Date(jahr, monat + 1, sollversand);
  // Schiebe zurück auf Freitag, falls Sa/So oder Feiertag
  while (d.getDay() === 0 || d.getDay() === 6 || feiertage.includes(d.toISOString().split('T')[0])) {
    d.setDate(d.getDate() - 1);
  }
  return d;
}

async function main() {
  console.log('Starte Berichtsexport. Heute:', new Date().toISOString());
  await fs.mkdir('./tmp', { recursive: true });

  const heute = new Date();
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

  // Logo laden als DataURL:
  const logoBuffer = await fs.readFile(path.resolve('templates/logo.png')); // Passe ggf. Dateiname an!
  const logoDataUrl = 'data:image/png;base64,' + logoBuffer.toString('base64');

  for (const kunde of kunden) {
    const { id: kunden_id, name: firma_name, lastversand, erstellungsdatum, sollversand, land, bundesland } = kunde;
    let von;
    if (lastversand) {
      von = new Date(heute.getFullYear(), heute.getMonth() - 1, lastversand);
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
      if (!zeiten.length) continue;

      // Summen berechnen
      const pauseSum = sumIntervals(zeiten.map(z => z.gesamt_pause));
      const nettoSum = sumIntervals(zeiten.map(z => z.gesamt_netto));
      const ueberSumRaw = sumIntervals(zeiten.map(z => z.ueber_unter_stunden));
      const ueberSum = formatSignedInterval(ueberSumRaw);

      // Tabellenzeilen aufbauen
      const tableRows = zeiten.map(z => `
        <tr>
          <td>${formatDateDE(z.datum)}</td>
          <td>${z.tagesstatus || ''}</td>
          <td>${z.erster_start ? z.erster_start.substring(11, 16) : ''}</td>
          <td>${z.letzter_ende ? z.letzter_ende.substring(11, 16) : ''}</td>
          <td>${intervalToStr(z.gesamt_pause)}</td>
          <td>${intervalToStr(z.gesamt_netto)}</td>
          <td>${formatSignedInterval(intervalToStr(z.ueber_unter_stunden))}</td>
        </tr>
      `).join('\n');

      const pdfVars = {
        firma_name,
        mitarbeiter_name: ma_name,
        zeitraum_start: formatDateDE(zeitraum_start),
        zeitraum_ende: formatDateDE(zeitraum_ende),
        table_rows: tableRows,
        summe_pause: pauseSum,
        summe_netto: nettoSum,
        summe_uebermin: ueberSum,
        logo_dataurl: logoDataUrl
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

    // lastversand & istversand setzen, wenn mind. ein Bericht erzeugt wurde
    if (berichteErzeugt && sollversand) {
      try {
        const nextMonth = (heute.getMonth() + 1) % 12;
        const nextYear = heute.getMonth() === 11 ? heute.getFullYear() + 1 : heute.getFullYear();
        const feiertage = await getFeiertage(
          land,
          bundesland,
          `${nextYear}-${String(nextMonth+1).padStart(2, '0')}-01`,
          `${nextYear}-${String(nextMonth+1).padStart(2, '0')}-31`
        );
        const nextVersand = getNextVersanddatum(sollversand, heute, feiertage);
        await supabase
          .from('kunden')
          .update({ lastversand: heute.getDate(), istversand: nextVersand.getDate() })
          .eq('id', kunden_id);
        console.log(`lastversand & istversand für ${firma_name} aktualisiert! Neuer istversand: ${nextVersand.toISOString().slice(0,10)}`);
      } catch (err) {
        console.error(`Fehler beim Aktualisieren von lastversand/istversand für ${firma_name}:`, err);
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
