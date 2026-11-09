// Juleøltest – backend i Google Sheets.
// Lim inn i Utvidelser → Apps Script i et nytt Google-regneark, og publiser som web-app.

const SVAR = 'Svar';   // én rad per deltager per øl
const OL = 'Øl';       // kolonne A: nr, kolonne B: navn, kolonne C: glass
const ANTALL_OL = 24;
const VEKTER = [1, 1, 2, 1];   // Farge, Aroma, Smak (dobbelt), Ettersmak

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(header); sh.setFrozenRows(1); }
  return sh;
}

function beerSheet_() {
  const sh = sheet_(OL, ['Nr', 'Navn', 'Glass']);
  const header = sh.getRange(1, 1, 1, 3).getValues()[0];
  if (header[0] !== 'Nr') sh.getRange(1, 1, 1, 3).setValues([['Nr', 'Navn', 'Glass']]);
  if (header.length < 3 || header[2] !== 'Glass') {
    sh.getRange(1, 3).setValue('Glass');
  }
  if (sh.getLastRow() <= 1) {
    const rows = [];
    for (let i = 1; i <= ANTALL_OL; i++) {
      rows.push([i, '', '']);
    }
    sh.getRange(2, 1, ANTALL_OL, 3).setValues(rows);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function svarArk_() {
  return sheet_(SVAR, ['Tidspunkt', 'DeltagerId', 'Navn', 'Øl', 'Farge', 'Aroma', 'Smak', 'Ettersmak', 'Score']);
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const action = String(d.action || 'score');

    if (action === 'registerGlass') {
      const beer = Number(d.beer);
      const glass = Number(d.glass);
      if (!(beer >= 1 && beer <= ANTALL_OL) || !(glass >= 1 && glass <= ANTALL_OL)) {
        return json_({ ok: false, error: 'Ugyldige verdier for øl eller glass' });
      }

      const sh = beerSheet_();
      const vals = sh.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < vals.length; i++) {
        if (Number(vals[i][0]) === beer) {
          sh.getRange(i + 1, 3).setValue(glass);
          found = true;
          break;
        }
      }
      if (!found) {
        sh.appendRow([beer, '', glass]);
      }
      return json_({ ok: true, beer: beer, glass: glass });
    }

    const beer = Number(d.beer);
    const scores = (d.scores || []).map(Number);
    if (!d.id || !(beer >= 1 && beer <= ANTALL_OL) || scores.length !== 4 || scores.some(s => !(s >= 1 && s <= 10))) {
      return json_({ ok: false, error: 'Ugyldige data' });
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sh = svarArk_();
      const vsum = VEKTER.reduce((a, b) => a + b, 0);
      const score = Math.round(scores.reduce((a, v, i) => a + v * VEKTER[i], 0) / vsum * 100) / 100;
      const row = [new Date(), String(d.id), String(d.name || '').slice(0, 40), beer].concat(scores, [score]);
      const vals = sh.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        if (String(vals[i][1]) === String(d.id) && Number(vals[i][3]) === beer) {
          sh.getRange(i + 1, 1, 1, row.length).setValues([row]);   // oppdater eksisterende svar
          return json_({ ok: true });
        }
      }
      sh.appendRow(row);
      return json_({ ok: true });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  const action = (e.parameter || {}).action;
  const rows = svarArk_().getDataRange().getValues().slice(1)
    .filter(r => r[1] !== '')
    .map(r => ({ id: String(r[1]), name: r[2], beer: Number(r[3]), s: r.slice(4, 8).map(Number) }));

  if (action === 'mine') {
    return json_({ rows: rows.filter(r => r.id === String(e.parameter.id)) });
  }

  if (action === 'beerlist') {
    const beers = beerSheet_().getDataRange().getValues().slice(1)
      .filter(r => r[0] !== '')
      .map(r => ({
        nr: Number(r[0]),
        name: String(r[1] || '').trim(),
        glass: r[2] === '' || r[2] == null ? null : Number(r[2])
      }))
      .filter(b => b.name !== '');
    return json_({ beers: beers });
  }

  const names = {};
  const glasses = {};
  const beerByGlass = {};
  beerSheet_().getDataRange().getValues().slice(1).forEach(r => {
    const nr = Number(r[0]);
    const glass = Number(r[2]);
    if (!Number.isNaN(nr) && r[1] !== '') names[nr] = String(r[1]);
    if (!Number.isNaN(nr) && !Number.isNaN(glass)) {
      glasses[nr] = glass;
      beerByGlass[glass] = nr;
    }
  });

  const enrichedRows = rows.map(row => {
    const answerGlass = Number(row.beer);
    const actualBeer = Number.isFinite(answerGlass) && beerByGlass[answerGlass] !== undefined
      ? Number(beerByGlass[answerGlass])
      : Number(row.beer);

    return {
      ...row,
      beer: actualBeer,
      glass: answerGlass,
      name: names[actualBeer] || row.name || `Øl ${actualBeer}`
    };
  });

  return json_({ rows: enrichedRows, names: names, glasses: glasses });
}