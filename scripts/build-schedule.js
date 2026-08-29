// scripts/build-schedule.js
//
// Leser "Treningstider 2026-27.xlsx" og oppdaterer VIEWS / WEEK_DATES
// inne i "Treningstider 2026-27.html", uten å røre resten av filen
// (CSS, øvrig JS-logikk, auto-resize-script osv).
//
// Kjøres av GitHub Action-workflowen etter manuell godkjenning.

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const XLSX_PATH = path.join(__dirname, '..', 'Treningstider 2026-27.xlsx');
const HTML_PATH = path.join(__dirname, '..', 'Treningstider 2026-27.html');

const DAYS = ['Mandag', 'Tirsdag', 'Onsdag', 'Torsdag', 'Fredag', 'Lørdag', 'Søndag'];

const COLOR_CAT = {
  '99CCFF': 'hockey',
  '3399FF': 'hockey',
  'FFCCFF': 'andrelag',
  'FFFF00': 'kunstlop',
  'FFE1CC': 'kamp',
  '777777': 'gray',
  '999999': 'publikumsskoyting',
  'ADADAD': 'publikumsskoyting',
  'B7B7B7': 'publikumsskoyting',
};

// Normaliserer hallnavn som skrives ulikt i ulike faner
// (store bokstaver / bindestrek / forkortelser) til navnene
// som allerede brukes i HTML-filen i dag.
const HALL_MAP = {
  'SIDDISHALLEN': 'Siddishallen',
  'SIDDISHALLEN ': 'Siddishallen',
  'STAVANGER ISHALL - ISBANE 1': 'Stavanger Ishall – Isbane 1',
  'STAVANGER ISHALL – ISBANE 1': 'Stavanger Ishall – Isbane 1',
  'ISBANE 1': 'Stavanger Ishall – Isbane 1',
  'STAVANGER ISHALL - ISBANE 2': 'Stavanger Ishall – Isbane 2',
  'STAVANGER ISHALL – ISBANE 2': 'Stavanger Ishall – Isbane 2',
  'ISBANE 2': 'Stavanger Ishall – Isbane 2',
  'DNB - ARENA': 'DNB Arena',
  'DNB-ARENA': 'DNB Arena',
  'DNB ARENA': 'DNB Arena',
  'SØRMARKA ARENA': 'Sørmarka Arena',
  'NÆRBØ': 'Nærbø',
};

function normalizeHall(raw) {
  if (!raw) return raw;
  const key = raw.trim().toUpperCase();
  return HALL_MAP[key] || raw.trim();
}

const MONTHS_NO = [
  'januar', 'februar', 'mars', 'april', 'mai', 'juni',
  'juli', 'august', 'september', 'oktober', 'november', 'desember',
];

function cellText(cell) {
  if (!cell || cell.v === undefined || cell.v === null) return '';
  return String(cell.v).trim();
}

function buildTimeMap(sheet) {
  // Kolonne B er normalt 07:00, 15 min steg, til og med kolonne BJ = 22:00.
  // Enkelte ark (oppdaget 2026-08-28, for å gi plass til en tidlig 06:45-økt) har fått
  // satt inn en ekstra kolonne foran hele tidsgriddet, og har dermed én kolonne mer enn
  // normalt (63 i stedet for 62, inkl. kolonne A). Vi oppdager dette ut fra arkets
  // faktiske bredde i stedet for å anta et fast sluttpunkt, slik at kolonne B blir
  // hhv. 07:00 (normalt) eller 06:45 (utvidet grid) — uten dette ville alle økter i et
  // utvidet ark blitt registrert 15 minutter for sent.
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const startCol = XLSX.utils.decode_col('B');
  const endCol = range.e.c;
  const totalCols = endCol - range.s.c + 1;
  const startMinutes = totalCols >= 63 ? 6 * 60 + 45 : 7 * 60;
  const map = {};
  for (let col = startCol; col <= endCol; col++) {
    const mins = startMinutes + (col - startCol) * 15;
    const hh = Math.floor(mins / 60);
    const mm = mins % 60;
    map[col] = pad2(hh) + ':' + pad2(mm);
  }
  return map;
}

function pad2(n) {
  return (n < 10 ? '0' : '') + n;
}

function addMinutes(hhmm, add) {
  const [h, m] = hhmm.split(':').map(Number);
  let mins = h * 60 + m + add;
  mins = ((mins % (24 * 60)) + 24 * 60) % (24 * 60);
  return pad2(Math.floor(mins / 60)) + ':' + pad2(mins % 60);
}

// Parser en "hall-rutenett"-fane (Normalsesong, Uke 33-40).
function parseHallSheet(sheet) {
  const timeMap = buildTimeMap(sheet);
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const mergesByRow = {};
  (sheet['!merges'] || []).forEach((m) => {
    const r = m.s.r;
    if (!mergesByRow[r]) mergesByRow[r] = [];
    mergesByRow[r].push(m);
  });

  let currentHall = null;
  const sessions = [];
  let weekTitle = null;

  for (let r = range.s.r; r <= range.e.r; r++) {
    const aVal = cellText(sheet[XLSX.utils.encode_cell({ r, c: 0 })]);

    if (r === range.s.r && aVal.indexOf('Uke') === 0) {
      weekTitle = aVal;
      continue;
    }

    if (DAYS.indexOf(aVal) !== -1) {
      const rowMerges = mergesByRow[r] || [];
      rowMerges.forEach((m) => {
        const cell = sheet[XLSX.utils.encode_cell(m.s)];
        let team = cellText(cell);

        let cat = 'hockey';
        const rgb = cell.s && cell.s.fgColor ? cell.s.fgColor.rgb : null;
        if (rgb && COLOR_CAT[rgb]) cat = COLOR_CAT[rgb];

        if (!team) {
          // Tomme, lysegrå-fargede celler representerer publikumsskøyting
          // (ingen lagnavn i Excel-arket, kun fargekoding).
          if (cat === 'publikumsskoyting') {
            team = 'Publikumsskøyting';
          } else {
            return;
          }
        }

        // "F/S"-celler i Excel-filen skal vises som "Andre lag" på nettsiden,
        // uavhengig av farge (selve Excel-filen beholder "F/S" uendret).
        if (/^f\/s$/i.test(team.trim())) {
          team = 'Andre lag';
          cat = 'andrelag';
        }

        const kampMatch = team.match(/^(.*)\s*-\s*kamp$/i);
        if (kampMatch) {
          team = kampMatch[1].trim();
          cat = 'kamp';
        }

        const startT = timeMap[m.s.c];
        const endT = timeMap[m.e.c] ? addMinutes(timeMap[m.e.c], 15) : null;
        if (!startT || !endT) return;

        sessions.push({
          hall: normalizeHall(currentHall),
          day: aVal,
          start: startT,
          end: endT,
          team,
          cat,
          approx: false,
        });
      });
      continue;
    }

    if (aVal !== '') {
      currentHall = aVal;
    }
  }

  return { weekTitle, sessions };
}

// Parser Styrkerom-fanen: rader = 15-min tidsintervaller, kolonner = ukedager,
// samme team-verdi gjentatt i alle radene økten varer (ikke sammenslåtte celler).
function parseStyrkeromSheet(sheet) {
  const range = XLSX.utils.decode_range(sheet['!ref']);
  // Finn header-raden med ukedagene (MANDAG, TIRSDAG, ...)
  let headerRow = -1;
  const dayCols = {}; // colIndex -> dagnavn
  for (let r = range.s.r; r <= Math.min(range.s.r + 5, range.e.r); r++) {
    const rowVals = [];
    for (let c = range.s.c + 1; c <= range.e.c; c++) {
      rowVals.push(cellText(sheet[XLSX.utils.encode_cell({ r, c })]).toUpperCase());
    }
    if (rowVals.some((v) => v === 'MANDAG')) {
      headerRow = r;
      for (let c = range.s.c + 1; c <= range.e.c; c++) {
        const v = cellText(sheet[XLSX.utils.encode_cell({ r, c })]);
        const dayIdx = DAYS.findIndex((d) => d.toUpperCase() === v.toUpperCase());
        if (dayIdx !== -1) dayCols[c] = DAYS[dayIdx];
      }
      break;
    }
  }
  if (headerRow === -1) return { sessions: [] };

  const sessions = [];
  Object.keys(dayCols).forEach((colStr) => {
    const c = Number(colStr);
    const day = dayCols[c];
    let run = null; // {team, cat, startLabel, endLabel}

    function flush() {
      if (!run) return;
      const startT = run.startLabel.split('-')[0].trim();
      const endT = run.endLabel.split('-')[1].trim();
      sessions.push({
        hall: 'Styrkerom',
        day,
        start: startT,
        end: endT,
        team: run.team,
        cat: run.cat,
        approx: false,
      });
      run = null;
    }

    for (let r = headerRow + 1; r <= range.e.r; r++) {
      const timeLabel = cellText(sheet[XLSX.utils.encode_cell({ r, c: 0 })]);
      if (!timeLabel || timeLabel.indexOf(':') === -1) continue; // slutt på tabellen

      const cell = sheet[XLSX.utils.encode_cell({ r, c })];
      const team = cellText(cell);

      if (!team) {
        flush();
        continue;
      }

      let cat = 'hockey';
      const rgb = cell.s && cell.s.fgColor ? cell.s.fgColor.rgb : null;
      if (rgb && COLOR_CAT[rgb]) cat = COLOR_CAT[rgb];

      if (run && run.team === team) {
        run.endLabel = timeLabel;
      } else {
        flush();
        run = { team, cat, startLabel: timeLabel, endLabel: timeLabel };
      }
    }
    flush();
  });

  return { sessions };
}

// Parser "Uke NN, D. [month] - D. month YYYY" til range-tekst + per-dag datoer.
function parseWeekTitle(weekTitle) {
  if (!weekTitle) return null;
  const re = /Uke\s*(\d+),\s*(\d{1,2})\.\s*(?:([a-zæøå]+)\s*)?-\s*(\d{1,2})\.\s*([a-zæøå]+)\s*(\d{4})/i;
  const m = weekTitle.match(re);
  if (!m) return null;

  const [, , startDayStr, startMonthNameRaw, endDayStr, endMonthName, yearStr] = m;
  const startDay = Number(startDayStr);
  const endDay = Number(endDayStr);
  const year = Number(yearStr);
  const endMonthIdx = MONTHS_NO.findIndex((mn) => mn === endMonthName.toLowerCase());
  const startMonthIdx = startMonthNameRaw
    ? MONTHS_NO.findIndex((mn) => mn === startMonthNameRaw.toLowerCase())
    : endMonthIdx;

  if (endMonthIdx === -1 || startMonthIdx === -1) return null;

  const startDate = new Date(year, startMonthIdx, startDay);
  const days = {};
  for (let i = 0; i < 7; i++) {
    const d = new Date(startDate);
    d.setDate(d.getDate() + i);
    days[DAYS[i]] = pad2(d.getDate()) + '.' + pad2(d.getMonth() + 1);
  }
  const endDate = new Date(startDate);
  endDate.setDate(endDate.getDate() + 6);

  const range = startMonthIdx === endMonthIdx
    ? `${startDay}.–${endDay}. ${endMonthName} ${year}`
    : `${startDay}. ${startMonthNameRaw} – ${endDay}. ${endMonthName} ${year}`;

  return { range, days, endDate };
}

function sessionsToJsArray(sessions) {
  // Samme feltrekkefølge som dagens VIEWS-data, for minst mulig diff.
  const items = sessions.map((s) => {
    if (s.hall) {
      return `{"hall": ${JSON.stringify(s.hall)}, "day": ${JSON.stringify(s.day)}, "start": ${JSON.stringify(s.start)}, "end": ${JSON.stringify(s.end)}, "team": ${JSON.stringify(s.team)}, "cat": ${JSON.stringify(s.cat)}, "approx": ${s.approx}}`;
    }
    return `{"day": ${JSON.stringify(s.day)}, "start": ${JSON.stringify(s.start)}, "end": ${JSON.stringify(s.end)}, "team": ${JSON.stringify(s.team)}, "hall": ${JSON.stringify(s.hall)}, "cat": ${JSON.stringify(s.cat)}, "approx": ${s.approx}}`;
  });
  return '[' + items.join(', ') + ']';
}

// Oppdager ark dynamisk i stedet for en hardkodet liste (fikset 2026-08-28 —
// en hardkodet liste opp til "Uke 40" førte til at Uke 41-45 ble stille
// hoppet over selv om de fantes i xlsx-en). Returnerer
// { normalSheetName, ukeSheets: [{key, sheetName, nr}, ...] sortert stigende,
//   styrkeromSheetName }.
function discoverSheets(wb) {
  const sheetNames = wb.SheetNames;
  const normalSheetName = sheetNames.find((s) => s.startsWith('Normalsesong')) || null;
  const ukeSheets = [];
  sheetNames.forEach((s) => {
    const m = s.trim().match(/^Uke\s+(\d+)$/);
    if (m) {
      const nr = Number(m[1]);
      ukeSheets.push({ key: `uke${nr}`, sheetName: s, nr });
    }
  });
  ukeSheets.sort((a, b) => a.nr - b.nr);
  const styrkeromSheetName = sheetNames.find((s) => s.startsWith('Styrkerom')) || null;
  return { normalSheetName, ukeSheets, styrkeromSheetName };
}

function main() {
  const wb = XLSX.readFile(XLSX_PATH, { cellStyles: true });
  let html = fs.readFileSync(HTML_PATH, 'utf8');
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const viewsEntries = [];
  const weekDatesEntries = [];
  const tabEntries = []; // {key, label} — normalsesong først, så uker, styrkerom sist

  const { normalSheetName, ukeSheets, styrkeromSheetName } = discoverSheets(wb);

  if (normalSheetName) {
    const { sessions } = parseHallSheet(wb.Sheets[normalSheetName]);
    viewsEntries.push(`  normalsesong: ${sessionsToJsArray(sessions)},`);
    tabEntries.push({ key: 'normalsesong', label: 'Istrening - Normalsesong' });
  } else {
    console.warn('Fant ikke noe Normalsesong-ark.');
  }

  const included = [];
  const pruned = [];
  ukeSheets.forEach(({ key, sheetName, nr }) => {
    const { weekTitle, sessions } = parseHallSheet(wb.Sheets[sheetName]);
    const parsed = weekTitle ? parseWeekTitle(weekTitle) : null;
    if (parsed && parsed.endDate < today) {
      pruned.push(`${key} (${parsed.range})`);
      return; // uken er passert — fjern automatisk fra nettsiden (bekreftet med bruker 28.08.2026)
    }
    viewsEntries.push(`  ${key}: ${sessionsToJsArray(sessions)},`);
    tabEntries.push({ key, label: `Istrening - Uke ${nr}` });
    included.push(key);
    if (parsed) {
      weekDatesEntries.push(
        `  ${key}: { range: ${JSON.stringify(parsed.range)}, days: { ${DAYS.map(
          (d) => `${d}:${JSON.stringify(parsed.days[d])}`
        ).join(', ')} } },`
      );
    }
  });

  if (styrkeromSheetName) {
    const { sessions } = parseStyrkeromSheet(wb.Sheets[styrkeromSheetName]);
    viewsEntries.push(`  styrkerom: ${sessionsToJsArray(sessions)},`);
    tabEntries.push({ key: 'styrkerom', label: 'Styrkerom - Normalsesong' });
  }

  console.log(`Inkluderte uker: ${included.join(', ') || '(ingen)'}`);
  if (pruned.length) {
    console.log(`Fjernet (passert dato, ${today.toISOString().slice(0, 10)}): ${pruned.join(', ')}`);
  }

  const newViewsBlock = `const VIEWS = {\n${viewsEntries.join('\n')}\n};`;
  const newWeekDatesBlock = `const WEEK_DATES = {\n${weekDatesEntries.join('\n')}\n};`;
  const newTabBar = `<div class="tab-bar" id="tabBar">\n${tabEntries
    .map(
      ({ key, label }, i) =>
        `      <button type="button" class="tab-btn${i === 0 ? ' active' : ''}" data-view="${key}" onclick="switchTab('${key}')">${label}</button>`
    )
    .join('\n')}\n    </div>`;

  const viewsRe = /const VIEWS = \{[\s\S]*?\n\s*\};/;
  const weekDatesRe = /const WEEK_DATES = \{[\s\S]*?\n\s*\};/;
  const tabBarRe = /<div class="tab-bar" id="tabBar">[\s\S]*?<\/div>/;

  if (!viewsRe.test(html)) {
    throw new Error('Fant ikke "const VIEWS = {...}" i HTML-filen - avbryter uten å skrive noe.');
  }
  if (!weekDatesRe.test(html)) {
    throw new Error('Fant ikke "const WEEK_DATES = {...}" i HTML-filen - avbryter uten å skrive noe.');
  }
  if (!tabBarRe.test(html)) {
    throw new Error('Fant ikke tab-bar-blokken i HTML-filen - avbryter uten å skrive noe.');
  }

  html = html.replace(viewsRe, newViewsBlock);
  html = html.replace(weekDatesRe, newWeekDatesBlock);
  html = html.replace(tabBarRe, newTabBar);

  fs.writeFileSync(HTML_PATH, html, 'utf8');
  console.log('HTML-filen er oppdatert fra Excel-filen.');
}

main();
