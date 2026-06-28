// ─── STATE ───────────────────────────────────────────────────
console.log('%cscript.js wersja: pivot-builder-v3 (auto-przeniesienie tekstu z Oś Y, naprawa przy każdym renderze)', 'color:#e8a23d;font-weight:bold;');

let allData = [], filteredData = [], cols = [], colTypes = {}, workbook = null, sheetNames = [];
let currentFileName = '';
let currentSheetName = '';
let sortCol = null, sortDir = 1;
let currentPage = 1;
const PAGE_SIZE = 50;
const chartInstances = {};

// ─── SHIFTS ─────────────────────────────────────────────────
// Default 2-shift split (06:00-18:00 / 18:00-06:00). Adjust SHIFTS
// once the real warehouse shift schedule is confirmed at launch —
// everything downstream (KPIs, insights, charts) reads from
// getScopedData(), so changing these two lines is enough.
const SHIFTS = [
  { name: 'Zmiana 1', start: 6, end: 18 },
  { name: 'Zmiana 2', start: 18, end: 6 },
];
let currentShift = 'all'; // 'all' | 0 | 1 (index into SHIFTS)

function setShiftFilter(v) {
  currentShift = v === 'all' ? 'all' : parseInt(v, 10);
  renderAll();
}

function hourInShift(hour, shiftIdx) {
  const s = SHIFTS[shiftIdx];
  if (s.start < s.end) return hour >= s.start && hour < s.end;
  return hour >= s.start || hour < s.end; // overnight shift wraps past midnight
}

// Returns allData filtered to the selected shift, using whatever
// time-like column is detected. If no time column exists, the shift
// filter has nothing to act on and silently falls back to allData.
// ─── ROW ENTITY DETECTION ───────────────────────────────────────
// What does "one row" actually mean in this file? Without this, every
// "Count" chart says generic "liczba wierszy", which tells the reader
// nothing. We scan column names once per file (priority order — first
// match wins) and use that word everywhere a row-count is described,
// so it reads as "liczba zamówień" / "liczba błędów" / "liczba paczek"
// instead of "liczba wierszy".
const ROW_ENTITY_RULES = [
  { re: /\bblad|error|wyjat|hospital|awari|usterk/i, label: 'zgłoszeń błędów' },
  { re: /numer.?zamow|order.?nr|order.?number|nr.?zamow|order.?id/i, label: 'zamówień' },
  { re: /tote|lpn|paczk|przesy[lł]/i, label: 'paczek' },
  { re: /transakcj|transaction/i, label: 'transakcji' },
  { re: /reklamacj|complaint/i, label: 'reklamacji' },
  { re: /dostaw|deliver|shipment/i, label: 'dostaw' },
  { re: /operator|pracownik|worker|employee|packer/i, label: 'wpisów pracowników' },
];
function getRowEntityLabel() {
  for (const rule of ROW_ENTITY_RULES) {
    if (cols.some(c => rule.re.test(c))) return rule.label;
  }
  return 'wpisów';
}

let periodFilterDays = 'all';

function setPeriodFilter(v) {
  periodFilterDays = v === 'all' ? 'all' : parseInt(v, 10);
  renderAll();
}

function getScopedData() {
  let data = allData;
  if (periodFilterDays !== 'all') {
    const dateCol = cols.find(c => colTypes[c] === 'date');
    if (dateCol) {
      const parsed = data.map(r => Date.parse(String(r[dateCol]).slice(0, 10))).filter(t => !isNaN(t));
      const maxDate = parsed.length ? Math.max(...parsed) : null;
      if (maxDate !== null) {
        const cutoff = maxDate - periodFilterDays * 86400000;
        data = data.filter(r => { const t = Date.parse(String(r[dateCol]).slice(0, 10)); return !isNaN(t) && t > cutoff; });
      }
    }
  }
  if (currentShift === 'all') return data;
  const timeCol = findCol(/time|hour|godzina/i) || cols.filter(c => colTypes[c] === 'date')[0] || null;
  if (!timeCol) return data;
  return data.filter(r => {
    const h = extractHour(r[timeCol]);
    return h !== null && hourInShift(h, currentShift);
  });
}

const COLORS = ['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4','#f97316','#ec4899','#6366f1','#84cc16','#14b8a6','#a855f7'];

// ─── UPLOAD ───────────────────────────────────────────────────
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');

function triggerUpload() { fileInput.click(); }
fileInput.addEventListener('change', e => { if(e.target.files[0]) readFile(e.target.files[0]); });
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('over'));
dropZone.addEventListener('drop', e => { e.preventDefault(); dropZone.classList.remove('over'); if(e.dataTransfer.files[0]) readFile(e.dataTransfer.files[0]); });

function showLoading(show) {
  let el = document.getElementById('loading-overlay');
  if (show && !el) {
    el = document.createElement('div');
    el.id = 'loading-overlay';
    el.style.cssText = 'position:fixed;inset:0;background:rgba(20,24,29,.75);display:flex;align-items:center;justify-content:center;z-index:9999;font-size:14px;font-weight:600;color:#e8ecef;';
    el.innerHTML = '<div style="background:#1c2228;border:1px solid #2e3640;padding:18px 28px;border-radius:10px;box-shadow:0 4px 20px rgba(0,0,0,.4);">⏳ Processing file…</div>';
    document.body.appendChild(el);
  } else if (!show && el) {
    el.remove();
  }
}

function readFile(file) {
  showLoading(true);
  const reader = new FileReader();
  reader.onload = e => {
    // Yield one frame so the loading overlay actually paints before the
    // (potentially multi-second, on large files) synchronous parse blocks the thread.
    setTimeout(() => {
      try {
        workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true });
        currentFileName = file.name;
        sheetNames = workbook.SheetNames;
        document.getElementById('ib-name').textContent = file.name;
        document.getElementById('header-filename').textContent = file.name;
        if(sheetNames.length > 1) {
          const sel = document.getElementById('sheet-select');
          sel.innerHTML = sheetNames.map((n,i) => `<option value="${i}">${n}</option>`).join('');
          document.getElementById('ib-sheet-wrap').style.display = 'flex';
        } else {
          document.getElementById('ib-sheet-wrap').style.display = 'none';
        }
        loadSheet(0);
      } catch(err) {
        alert('Error reading file: ' + err.message);
      } finally {
        showLoading(false);
      }
    }, 30);
  };
  reader.onerror = () => { showLoading(false); alert('Could not read the file.'); };
  reader.readAsArrayBuffer(file);
}

// Some manual/legacy exports put a free-text title row above the real
// header (e.g. "Outbound Pack Online - Daily Report"). Scan the first
// few rows and pick the one with the most filled cells as the header —
// a title row typically has only 1, while a real header row has many.
function findHeaderRowIndex(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  let bestIdx = 0, bestCount = -1;
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const nonEmpty = rows[i].filter(c => c !== '' && c != null).length;
    if (nonEmpty > bestCount) { bestCount = nonEmpty; bestIdx = i; }
  }
  return bestIdx;
}

function loadSheet(idx) {
  currentSheetName = sheetNames[idx] || '';
  const ws = workbook.Sheets[currentSheetName];
  const headerRow = findHeaderRowIndex(ws);
  const rawUntrimmed = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false, range: headerRow })
    .filter(row => Object.values(row).some(v => v !== '' && v != null));
  if (!rawUntrimmed.length) { alert('Sheet is empty'); return; }

  // Trim header names and string values. Manual exports very often carry
  // stray leading/trailing spaces (copy-paste, merged-cell artifacts) which
  // would otherwise silently split one real value into two categories
  // (e.g. "H&M" and "H&M " counted as different brands).
  const raw = rawUntrimmed.map(row => {
    const out = {};
    Object.keys(row).forEach(k => {
      const key = String(k).trim();
      let val = row[k];
      if (typeof val === 'string') val = val.trim();
      out[key] = val;
    });
    return out;
  });

  allData = raw;
  cols = Object.keys(raw[0]);

  // SAP sometimes exports dates as raw Excel serial numbers (e.g. 46054)
  // with no date format applied to the cell, so they come through as plain
  // numbers. If a column's name hints at a date and its values look like
  // plausible serials, convert them to ISO date strings in place so the
  // normal type detection / date-trend / shift logic picks them up.
  cols.forEach(c => {
    if (!/date|data\b/i.test(c)) return;
    const vals = raw.map(r => r[c]).filter(v => v !== '');
    if (!vals.length) return;
    const looksLikeSerial = vals.filter(v => /^\d{4,6}$/.test(String(v).trim())).length / vals.length > 0.8;
    if (!looksLikeSerial) return;
    raw.forEach(r => {
      const v = String(r[c]).trim();
      if (/^\d{4,6}$/.test(v)) {
        const serial = parseFloat(v);
        if (serial > 25569 && serial < 60000) {
          r[c] = new Date((serial - 25569) * 86400 * 1000).toISOString().slice(0, 10);
        }
      }
    });
  });

  colTypes = {};
  cols.forEach(c => colTypes[c] = detectType(c, raw));
  filteredData = [...allData];
  document.getElementById('ib-rows').textContent = allData.length.toLocaleString('en-US');
  document.getElementById('ib-cols').textContent = cols.length;
  currentShift = 'all';
  document.getElementById('shift-select').value = 'all';
  const hasTimeCol = !!(findCol(/time|hour|godzina/i) || cols.filter(c => colTypes[c] === 'date')[0]);
  document.getElementById('shift-wrap').style.display = hasTimeCol ? 'flex' : 'none';
  periodFilterDays = 'all';
  document.getElementById('period-select').value = 'all';
  const hasDateCol = cols.some(c => colTypes[c] === 'date');
  document.getElementById('period-wrap').style.display = hasDateCol ? 'flex' : 'none';
  pivotState = { filters: [], columns: [], rows: [], values: [] };
  fileInput.value = '';
  document.getElementById('upload-screen').style.display = 'none';
  document.getElementById('app').style.display = 'block';
  builderFilters = [];
  calculatedFields = [];
  if (typeof saveCurrentWorkspaceSnapshot === 'function') saveCurrentWorkspaceSnapshot();
  renderAll();
}

// ─── TYPE DETECTION ───────────────────────────────────────────
function detectType(col, data) {
  const vals = data.map(r => r[col]).filter(v => v !== '' && v != null);
  if (!vals.length) return 'unknown';
  const numC = vals.filter(v => toNum(v) !== null).length;
  if (numC / vals.length > 0.75) return 'numeric';
  const datePatterns = [/^\d{4}-\d{2}-\d{2}/, /^\d{2}[./-]\d{2}[./-]\d{4}/, /^\d{2}[./-]\d{2}[./-]\d{2}$/];
  const dateC = vals.filter(v => datePatterns.some(p => p.test(String(v).trim()))).length;
  if (dateC / vals.length > 0.5) return 'date';
  const uniq = new Set(vals.map(v => String(v))).size;
  if (uniq <= Math.min(30, vals.length * 0.55)) return 'categorical';
  return 'text';
}

function toNum(v) {
  if (typeof v === 'number') return v;
  let s = String(v).trim();
  if (!s) return null;
  // Clock-style duration "HH:MM" or "HH:MM:SS" (Excel's Time cell format)
  // is treated as a duration and converted to total minutes, so columns
  // like "Czas reakcji" / "Czas naprawy" can be summed and averaged like
  // any other numeric metric, instead of being stuck as text forever.
  const durMatch = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (durMatch) {
    const h = parseInt(durMatch[1], 10), m = parseInt(durMatch[2], 10), sec = durMatch[3] ? parseInt(durMatch[3], 10) : 0;
    return h * 60 + m + sec / 60;
  }
  // European decimal comma: "44,5" or "1.234,56" -> 44.5 / 1234.56.
  // Only treated as decimal when there's exactly one comma followed by
  // 1-2 digits at the end (a real thousands-group comma like "1,234"
  // is followed by exactly 3 digits, so it's excluded here).
  if (/^-?[\d.]*\d,\d{1,2}$/.test(s) && !/,\d{3}\b/.test(s)) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else {
    s = s.replace(/[\s,]/g, '');
  }
  // parseFloat silently stops at the first non-numeric character, so
  // "14:00" would otherwise become 14. Require the cleaned string to be
  // a full number, or reject it (catches error codes like #N/A and other
  // partial matches — real HH:MM was already handled above).
  if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function fmt(n, decimals = 2) {
  if (n == null || isNaN(n)) return '—';
  if (Math.abs(n) >= 1e6) return (n/1e6).toFixed(1).replace(/\.0$/, '') + 'M';
  if (Math.abs(n) >= 1e3) return (n/1e3).toFixed(1).replace(/\.0$/, '') + 'K';
  return Number.isInteger(n) ? n.toLocaleString('en-US') : n.toFixed(decimals);
}

// ─── CHART INSIGHTS ─────────────────────────────────────────
// One short, plain-language sentence under every chart: who leads,
// how concentrated the result is, where the weak point is, or — for
// time-ordered charts — which direction the trend is moving. This is
// the same number that's already in the chart, just named in words,
// so you don't have to eyeball every bar to get the takeaway.
function namesAtValue(labels, data, value) {
  const names = labels.filter((_, i) => data[i] === value);
  if (names.length === 1) return `«${names[0]}»`;
  if (names.length === 2) return `«${names[0]}» i «${names[1]}»`;
  return `«${names[0]}» i jeszcze ${names.length - 1}`;
}

function insightText(labels, data, opts = {}) {
  if (!labels || !labels.length || !data || !data.length) return '';
  const total = data.reduce((a, b) => a + b, 0);

  if (opts.trend) {
    const mid = Math.floor(data.length / 2) || 1;
    const firstAvg = data.slice(0, mid).reduce((a, b) => a + b, 0) / mid;
    const secondAvg = data.slice(mid).reduce((a, b) => a + b, 0) / ((data.length - mid) || 1);
    const peakVal = Math.max(...data);
    const lowVal = Math.min(...data);
    const peakIdx = data.indexOf(peakVal);
    const lowIdx = data.indexOf(lowVal);
    let dir = 'stabilny, bez wyraźnej dynamiki';
    if (secondAvg > firstAvg * 1.1) dir = 'rośnie';
    else if (secondAvg < firstAvg * 0.9) dir = 'spada';
    return `📌 Trend ${dir}. Szczyt — ${labels[peakIdx]} (${fmt(peakVal)}), najniżej — ${labels[lowIdx]} (${fmt(lowVal)}).`;
  }

  const maxVal = Math.max(...data);
  const minVal = Math.min(...data);
  const tiedTop = data.filter(v => v === maxVal).length;
  const tiedLow = data.filter(v => v === minVal).length;
  const share = total ? (maxVal / total * 100) : 0;

  let txt = tiedTop > 1
    ? `📌 ${tiedTop} wartości mają ten sam najwyższy wynik — ${namesAtValue(labels, data, maxVal)} (po ${fmt(maxVal)}${total ? `, razem ${ (share*tiedTop).toFixed(0)}% sumy` : ''}).`
    : `📌 Top — ${namesAtValue(labels, data, maxVal)} (${fmt(maxVal)}${total ? `, ${share.toFixed(0)}% sumy` : ''}).`;

  if (data.length > 2 && minVal !== maxVal) {
    txt += tiedLow > 1
      ? ` Najniższy wynik (${fmt(minVal)}) ma jednocześnie ${tiedLow}: ${namesAtValue(labels, data, minVal)}.`
      : ` Najniższy wynik — ${namesAtValue(labels, data, minVal)} (${fmt(minVal)}).`;
  }
  if (tiedTop === 1 && share > 50) txt += ' ⚠️ Silna koncentracja na jednej wartości — warto sprawdzić, czy to normalna sytuacja.';
  else if (data.length >= 4 && maxVal > 0 && minVal / maxVal < 0.15) txt += ' Duża różnica między Top a pozostałymi.';
  return txt;
}

function insightDiv(labels, data, opts = {}) {
  const t = insightText(labels, data, opts);
  return t ? `<div class="chart-insight">${t}</div>` : '';
}

// ─── RENDER ALL ───────────────────────────────────────────────
function renderAll() {
  const banner = document.getElementById('shift-empty-banner');
  const scopedEmpty = getScopedData().length === 0;
  banner.style.display = scopedEmpty ? 'block' : 'none';
  clearCustomCharts();
  // NOTE: the old UI (headline, topic tiles, relationship explorer, manual
  // axis/agg builder, auto-charts, outbound/insights sections) is no longer
  // shown — replaced by the Excel-style pivot builder below. All of those
  // functions are still fully defined further up in this file and work
  // correctly; they're just not called automatically anymore since their
  // DOM containers were removed. Re-add a call here (and the matching HTML)
  // to bring any of them back.
  renderPivotFieldList();
  renderPivotZones();
  renderExcelPivot();
  renderTable();
  renderFilterSelects();
  if (typeof renderOpsCenter === 'function') renderOpsCenter();
}

// ─── CUSTOM CHARTS CLEANUP (on new file/sheet load) ────────────
function toggleBuilderAdvanced() {
  const section = document.getElementById('builder-advanced');
  const btn = document.getElementById('builder-advanced-toggle-btn');
  const show = section.style.display === 'none';
  section.style.display = show ? 'block' : 'none';
  btn.textContent = show
    ? '▴ Schowaj opcje zaawansowane'
    : '▾ Opcje zaawansowane (tytuł, filtry, podział kolorami, limit kategorii)';
}

// ─── GENERIC PATTERN ENGINE ─────────────────────────────────────
// Finds patterns the same way for ANY file, using plain statistics —
// not hardcoded business rules. Works on whatever numeric/categorical/
// date/time columns the detector found, so it doesn't need to know in
// advance what "brand" or "operator" means.
function mean(arr) { return arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0; }
function stdevOf(arr) { const m = mean(arr); return Math.sqrt(mean(arr.map(v => (v - m) ** 2))); }
function pearson(xs, ys) {
  const n = xs.length;
  if (n < 3) return null;
  const mx = mean(xs), my = mean(ys);
  let num = 0, dx2 = 0, dy2 = 0;
  for (let i = 0; i < n; i++) { const dx = xs[i] - mx, dy = ys[i] - my; num += dx * dy; dx2 += dx * dx; dy2 += dy * dy; }
  const denom = Math.sqrt(dx2 * dy2);
  return denom ? num / denom : null;
}

function toggleGenericSection() {
  const section = document.getElementById('generic-section');
  const btn = document.getElementById('generic-toggle-btn');
  const show = section.style.display === 'none';
  section.style.display = show ? 'block' : 'none';
  btn.textContent = show
    ? '▴ Schowaj szybkie statystyki pliku (KPI, wykresy automatyczne)'
    : '▾ Pokaż szybkie statystyki pliku (KPI, wykresy automatyczne)';
}

function clearCustomCharts() {
  Object.keys(chartInstances).forEach(id => {
    if (id.startsWith('cc_')) {
      try { chartInstances[id].destroy(); } catch(e){}
      delete chartInstances[id];
    }
  });
  const area = document.getElementById('custom-charts-area');
  if (area) area.innerHTML = '';
  customChartCount = 0;
}

// ─── BADGES ───────────────────────────────────────────────────
// ─── HEADLINE SUMMARY ───────────────────────────────────────────
// One plain-sentence paragraph at the very top — no jargon, no chart,
// just the 2-3 things worth knowing about this file, written like a
// person would say it out loud.
function renderHeadlineSummary() {
  const host = document.getElementById('headline-section');
  if (!host) return;
  const data = getScopedData();
  if (data.length < 4) { host.innerHTML = ''; return; }

  const numCols = cols.filter(c => colTypes[c] === 'numeric');
  const catCols = cols.filter(c => colTypes[c] === 'categorical');
  const dateCol = cols.filter(c => colTypes[c] === 'date')[0];
  const mainNum = numCols.find(c => /qty|ilo|sztuk|liczba/i.test(c)) || numCols[0] || null;
  const mainCat = catCols.find(c => /operator|pracownik|worker|user/i.test(c)) || catCols[0] || null;

  const sentences = [];
  sentences.push(`Plik zawiera ${fmt(data.length)} ${getRowEntityLabel()}${mainNum ? ` i kolumnę „${mainNum}”` : ''}.`);

  if (mainNum && mainCat) {
    const groups = {};
    data.forEach(r => { const k = String(r[mainCat] ?? ''); const v = toNum(r[mainNum]); if (v !== null && k) (groups[k] = groups[k] || []).push(v); });
    const sums = Object.entries(groups).map(([k, arr]) => [k, arr.reduce((s, v) => s + v, 0)]).sort((a, b) => b[1] - a[1]);
    if (sums.length >= 2) {
      sentences.push(`Najlepszy wynik „${mainNum}” ma „${sums[0][0]}” (${fmt(sums[0][1])}), najsłabszy — „${sums[sums.length - 1][0]}” (${fmt(sums[sums.length - 1][1])}).`);
    }
  }
  if (dateCol && mainNum) {
    const byDate = {};
    data.forEach(r => { const d = String(r[dateCol]).slice(0, 10); const v = toNum(r[mainNum]); if (v !== null && d) byDate[d] = (byDate[d] || 0) + v; });
    const days = Object.keys(byDate).sort();
    if (days.length >= 2) {
      const best = days.reduce((a, b) => byDate[a] >= byDate[b] ? a : b);
      const worst = days.reduce((a, b) => byDate[a] <= byDate[b] ? a : b);
      sentences.push(`Najlepszy dzień — ${best} (${fmt(byDate[best])}), najsłabszy — ${worst} (${fmt(byDate[worst])}).`);
    }
  }
  const statusCol = cols.find(c => /status|hospital/i.test(c) && colTypes[c] === 'categorical');
  if (statusCol) {
    const bad = data.filter(r => /hospital|błąd|error|fail|niedob/i.test(String(r[statusCol]))).length;
    if (bad > 0) sentences.push(`${fmt(bad)} ${getRowEntityLabel()} (${(bad / data.length * 100).toFixed(0)}%) ma status wymagający uwagi.`);
  }

  host.innerHTML = `<div class="chart-card" style="margin-bottom:16px;border-color:var(--amber);">
    <div style="font-size:16px;line-height:1.6;color:var(--ink);">💬 ${sentences.join(' ')}</div>
  </div>`;
}

// ─── TOPIC TILES ────────────────────────────────────────────────
// Big one-click "ready-made questions" instead of making a non-technical
// person assemble X/Y/aggregation themselves. Each tile is a plain-language
// question; clicking it immediately renders the answer (table+chart+summary)
// using the same engine as the manual relationship explorer.
function detectTopics() {
  const numCols = cols.filter(c => colTypes[c] === 'numeric');
  const mainNum = numCols.find(c => /qty|ilo|sztuk|liczba|waga|anzahl|menge/i.test(c)) || numCols[0] || null;
  const find = re => cols.find(c => re.test(c));
  const topics = [];

  const personCol = find(/operator|pracownik|worker|employee|packer|mitarbeiter|^user$/i) || find(/leader|kierownik/i);
  if (personCol && mainNum) topics.push({ icon: '👥', label: 'Kto jest najlepszym pracownikiem?', a: personCol, b: mainNum });

  const stationCol = find(/stacj|station|workstation|zone|strefa|bereich|zona/i);
  const errorLikeCol = find(/blad|error|wyjat|hospital|usterk|awari|fehler|defekt/i);
  const isErrorNumeric = errorLikeCol && colTypes[errorLikeCol] === 'numeric';
  if (stationCol && errorLikeCol && !isErrorNumeric && stationCol !== errorLikeCol) {
    topics.push({ icon: '⚠️', label: 'Która strefa ma najwięcej błędów?', a: stationCol, b: errorLikeCol });
  } else if (errorLikeCol && !isErrorNumeric) {
    const statusCol = find(/status/i);
    if (statusCol && statusCol !== errorLikeCol) topics.push({ icon: '⚠️', label: 'Gdzie najwięcej błędów?', a: errorLikeCol, b: statusCol });
    else if (mainNum) topics.push({ icon: '⚠️', label: 'Gdzie najwięcej błędów?', a: errorLikeCol, b: mainNum });
  } else if (stationCol && mainNum) {
    topics.push({ icon: '🏭', label: 'Która stacja jest najlepsza?', a: stationCol, b: mainNum });
  }

  const shiftCol = find(/zmiana|shift|schicht/i);
  if (shiftCol && mainNum) topics.push({ icon: '🕐', label: 'Jak porównują się zmiany?', a: shiftCol, b: mainNum });

  const timeLossCol = find(/przestoj|idle|downtime|stillstand|strata|opoznien|delay|verzog/i) ||
    (find(/czas|time|zeit|dauer/i) && colTypes[find(/czas|time|zeit|dauer/i)] === 'numeric' ? find(/czas|time|zeit|dauer/i) : null);
  if (timeLossCol && personCol && timeLossCol !== personCol) topics.push({ icon: '⏱️', label: 'Gdzie największa strata czasu?', a: personCol, b: timeLossCol });
  else if (timeLossCol && stationCol && timeLossCol !== stationCol) topics.push({ icon: '⏱️', label: 'Gdzie największa strata czasu?', a: stationCol, b: timeLossCol });

  const brandCol = find(/brand|marka|rband|marke/i);
  if (brandCol && mainNum) topics.push({ icon: '🏷️', label: 'Jakie marki dominują?', a: brandCol, b: mainNum });

  const dateCol = cols.find(c => colTypes[c] === 'date');
  if (dateCol && mainNum) topics.push({ icon: '📅', label: 'Jak zmienia się to w czasie?', a: dateCol, b: mainNum });

  const orderCol = find(/zamow|order|bestellung/i);
  if (orderCol && brandCol && orderCol !== brandCol) topics.push({ icon: '📦', label: 'Ile zamówień na markę?', a: brandCol, b: orderCol });

  // de-dupe by a+b pair, keep first occurrence (priority order above)
  const seen = new Set();
  return topics.filter(t => { const k = t.a + '|' + t.b; if (seen.has(k)) return false; seen.add(k); return true; }).slice(0, 8);
}

function renderTopicTiles() {
  const host = document.getElementById('topics-section');
  if (!host) return;
  const topics = detectTopics();
  host.innerHTML = `
    <div style="font-size:12px;color:var(--ink-dim);margin-bottom:10px;text-transform:uppercase;letter-spacing:.4px;">Kliknij, żeby zobaczyć</div>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:12px;margin-bottom:20px;">
      ${topics.map((t, i) => `
        <button onclick="runTopicTile(${i})" style="background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:20px 14px;cursor:pointer;text-align:center;transition:all .15s;" onmouseover="this.style.borderColor='var(--amber)'" onmouseout="this.style.borderColor='var(--line)'">
          <div style="font-size:32px;margin-bottom:8px;">${t.icon}</div>
          <div style="font-size:13.5px;color:var(--ink);font-weight:600;line-height:1.3;">${escAttr(t.label)}</div>
        </button>`).join('')}
      <button onclick="scrollToOtherSearch()" style="background:var(--panel);border:1px dashed var(--ink-dim);border-radius:10px;padding:20px 14px;cursor:pointer;text-align:center;transition:all .15s;" onmouseover="this.style.borderColor='var(--amber)'" onmouseout="this.style.borderColor='var(--ink-dim)'">
        <div style="font-size:32px;margin-bottom:8px;">❓</div>
        <div style="font-size:13.5px;color:var(--ink);font-weight:600;line-height:1.3;">Szukam czegoś innego</div>
      </button>
    </div>`;
  window.__topicTiles = topics;
}

function scrollToOtherSearch() {
  const box = document.getElementById('rel-search');
  box.closest('.chart-builder').scrollIntoView({ behavior: 'smooth', block: 'center' });
  setTimeout(() => box.focus(), 400);
}

function runTopicTile(i) {
  const t = (window.__topicTiles || [])[i];
  if (!t) return;
  renderAnswerCard(t.a, t.b);
  document.getElementById('custom-charts-area').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ─── ANSWER CARD (manager path) ─────────────────────────────────
// One sentence + one KPI number + one small chart — no raw table, no
// "Suma/Średnia" jargon. The number itself is always RELATIVE (% vs
// the average of the rest), because a bare sum or average answers
// "what" but not "is that good or bad", which is what a manager
// actually wants to know in one glance.
function renderAnswerCard(a, b) {
  const data = getBuilderFilteredData();
  if (!data.length) { showBuilderValidation('Brak danych — sprawdź filtry okresu/zmiany.'); return; }
  const ta = colTypes[a], tb = colTypes[b];
  const numCol = ta === 'numeric' ? a : (tb === 'numeric' ? b : null);
  const catCol = numCol === a ? b : a;

  let sentence = '', kpiValue = '', kpiLabel = '', chartLabels = [], chartData = [];

  if (numCol && catCol && catCol !== numCol) {
    // category × number -> compare groups, relative to the average of the rest
    const groups = {};
    data.forEach(r => {
      const k = String(r[catCol] ?? '').trim();
      const v = toNum(r[numCol]);
      if (v === null || !k) return;
      (groups[k] = groups[k] || []).push(v);
    });
    const unit = getColumnUnit(numCol);
    const fmtU = v => fmt(v) + (unit ? ' ' + unit : '');
    const entries = Object.entries(groups).map(([k, arr]) => [k, arr.reduce((s, v) => s + v, 0), arr.length]);
    if (entries.length < 2) { showBuilderValidation('Za mało kategorii do porównania.'); return; }
    entries.sort((x, y) => y[1] - x[1]);
    const [bestName, bestVal] = entries[0];
    const rest = entries.slice(1);
    const restAvg = mean(rest.map(e => e[1]));
    const pct = restAvg ? ((bestVal - restAvg) / restAvg * 100) : 0;
    const worst = entries[entries.length - 1];

    sentence = `🏆 <strong>${escAttr(bestName)}</strong> ma najlepszy wynik „${b === catCol ? a : b}”${pct !== 0 ? ` — o ${Math.abs(pct).toFixed(0)}% ${pct >= 0 ? 'wyżej' : 'niżej'} niż średnia reszty` : ''}. Najsłabszy: <strong>${escAttr(worst[0])}</strong> (${fmtU(worst[1])}).`;
    kpiValue = (pct >= 0 ? '+' : '') + pct.toFixed(0) + '%';
    kpiLabel = `${bestName} vs średnia reszty`;
    chartLabels = entries.slice(0, 8).map(e => e[0]);
    chartData = entries.slice(0, 8).map(e => Math.round(e[1] * 100) / 100);
  } else if (ta === 'date' || tb === 'date') {
    const dateCol = ta === 'date' ? a : b;
    const valCol = dateCol === a ? b : a;
    const byDate = {};
    data.forEach(r => {
      const d = String(r[dateCol]).slice(0, 10);
      const v = toNum(r[valCol]);
      if (v === null || !d) return;
      byDate[d] = (byDate[d] || 0) + v;
    });
    const days = Object.keys(byDate).sort();
    if (days.length < 2) { showBuilderValidation('Za mało dni do oceny trendu.'); return; }
    const mid = Math.floor(days.length / 2) || 1;
    const f = mean(days.slice(0, mid).map(d => byDate[d])), s = mean(days.slice(mid).map(d => byDate[d]));
    const pct = f ? ((s - f) / f * 100) : 0;
    sentence = `📅 Wynik „${valCol}” ${pct > 5 ? 'rośnie' : pct < -5 ? 'spada' : 'jest stabilny'} w czasie — druga połowa okresu różni się o ${pct.toFixed(0)}% względem pierwszej.`;
    kpiValue = (pct >= 0 ? '+' : '') + pct.toFixed(0) + '%';
    kpiLabel = 'zmiana: 1. połowa → 2. połowa okresu';
    chartLabels = days; chartData = days.map(d => Math.round(byDate[d] * 100) / 100);
  } else {
    // category × category fallback -> just delegate to the full explorer (rare path for tiles)
    exploreRelationship(a, b, ta, tb);
    return;
  }

  const id = `cc_${customChartCount++}`;
  const wrap = document.createElement('div');
  wrap.className = 'chart-card';
  wrap.id = id;
  wrap.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:10px;">
      <div style="font-size:15px;line-height:1.6;color:var(--ink);max-width:70%;">${sentence}</div>
      <button onclick="removeChart(this,'${id}')" style="background:none;border:none;cursor:pointer;font-size:16px;color:var(--ink-dim);flex-shrink:0;">✕</button>
    </div>
    <div style="display:flex;gap:24px;align-items:center;margin-top:14px;flex-wrap:wrap;">
      <div style="text-align:center;min-width:120px;">
        <div style="font-family:var(--mono);font-size:32px;font-weight:700;color:var(--amber);">${escAttr(kpiValue)}</div>
        <div style="font-size:11px;color:var(--ink-dim);margin-top:2px;">${escAttr(kpiLabel)}</div>
      </div>
      <div style="flex:1;min-width:220px;position:relative;height:140px;"><canvas id="${id}_chart"></canvas></div>
    </div>`;
  document.getElementById('custom-charts-area').insertBefore(wrap, document.getElementById('custom-charts-area').firstChild);
  setTimeout(() => createChart(`${id}_chart`, 'bar', chartLabels, chartData, b, false, true, catCol || a), 50);
}

function renderBadges() {
  const map = { numeric:'badge-num', categorical:'badge-cat', date:'badge-date', time:'badge-date', text:'badge-txt', unknown:'badge-txt' };
  const icons = { numeric:'#', categorical:'≡', date:'📅', time:'🕐', text:'Aa', unknown:'?' };
  document.getElementById('col-badges').innerHTML = cols.map(c =>
    `<span class="badge ${map[colTypes[c]]}" title="${colTypes[c]}">${icons[colTypes[c]]} ${c}</span>`
  ).join('');
}

// ─── KPIs ─────────────────────────────────────────────────────
function renderKPIs() {
  const allData = getScopedData();
  const numCols = cols.filter(c => colTypes[c] === 'numeric');
  const accents = ['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4','#f97316','#ec4899'];
  if (!numCols.length) {
    document.getElementById('kpi-section').style.display = 'none';
    return;
  }
  document.getElementById('kpi-section').style.display = 'block';
  document.getElementById('kpi-grid').innerHTML = numCols.slice(0, 8).map((col, i) => {
    const vals = allData.map(r => toNum(r[col])).filter(v => v !== null);
    const sum = vals.reduce((a,b) => a+b, 0);
    const avg = vals.length ? sum / vals.length : 0;
    const max = vals.length ? Math.max(...vals) : 0;
    const min = vals.length ? Math.min(...vals) : 0;
    const accent = accents[i % accents.length];
    return `<div class="kpi-card" style="--accent:${accent}">
      <div class="kpi-label">${col}</div>
      <div class="kpi-value" style="color:${accent}">${fmt(sum)}</div>
      <div class="kpi-meta">
        Avg: <span>${fmt(avg)}</span> &nbsp;|&nbsp;
        Max: <span>${fmt(max)}</span> &nbsp;|&nbsp;
        Min: <span>${fmt(min)}</span> <br>
        Rows with data: <span>${vals.length}</span>
      </div>
    </div>`;
  }).join('');
}

// ─── AUTO CHARTS ──────────────────────────────────────────────
function renderAutoCharts() {
  const allData = getScopedData();
  const catCols = cols.filter(c => colTypes[c] === 'categorical');
  const numCols = cols.filter(c => colTypes[c] === 'numeric');
  const dateCols = cols.filter(c => colTypes[c] === 'date');
  const grid = document.getElementById('auto-charts-grid');
  grid.innerHTML = '';

  // Destroy old charts
  Object.values(chartInstances).forEach(c => { try { c.destroy(); } catch(e){} });
  for(const k in chartInstances) delete chartInstances[k];

  let chartIdx = 0;

  // Cat frequency charts (first 3 cat cols)
  catCols.slice(0, 3).forEach((col, ci) => {
    const freq = {};
    allData.forEach(r => { const v = String(r[col]) || '(empty)'; freq[v] = (freq[v]||0) + 1; });
    const sorted = Object.entries(freq).sort((a,b) => b[1]-a[1]).slice(0, 12);
    const id = `ac_${chartIdx++}`;
    const isH = sorted.length > 5;
    const h = isH ? Math.max(180, sorted.length * 30) : 220;
    const card = document.createElement('div');
    card.className = 'chart-card';
    const entity = getRowEntityLabel();
    card.innerHTML = `<div class="chart-card-title">${col}</div>
      <div class="chart-card-sub">Oś X: „${col}” (np. ID operatora, stacji, marki) · Oś Y: liczba ${entity} przypadająca na tę wartość — to ${entity}, NIE suma sztuk towaru</div>
      <div style="position:relative;width:100%;height:${h}px;">
        <canvas id="${id}" role="img" aria-label="Frequency by ${col}"></canvas>
      </div>
      ${insightDiv(sorted.map(e=>e[0]), sorted.map(e=>e[1]))}`;
    grid.appendChild(card);
    setTimeout(() => createChart(id, sorted.length <= 5 ? 'doughnut' : 'bar', sorted.map(e=>e[0]), sorted.map(e=>e[1]), col, isH, false, col), 50);
  });

  // Numeric grouped by first cat col
  if (numCols.length && catCols.length) {
    const gCol = catCols[0], nCol = numCols[0];
    const groups = {};
    allData.forEach(r => {
      const k = String(r[gCol]) || '(empty)';
      const v = toNum(r[nCol]);
      if (v !== null) groups[k] = (groups[k]||0) + v;
    });
    const sorted = Object.entries(groups).sort((a,b) => b[1]-a[1]).slice(0, 12);
    if (sorted.length > 1) {
      const id = `ac_${chartIdx++}`;
      const isH = sorted.length > 5;
      const h = isH ? Math.max(180, sorted.length * 30) : 220;
      const card = document.createElement('div');
      card.className = 'chart-card';
      card.innerHTML = `<div class="chart-card-title">${nCol} według ${gCol}</div>
        <div class="chart-card-sub">Oś X: „${gCol}” · Oś Y: suma „${nCol}” — czyli łączny wynik pracy (np. sztuk towaru) dla tej wartości, a NIE liczba ${getRowEntityLabel()}</div>
        <div style="position:relative;width:100%;height:${h}px;">
          <canvas id="${id}" role="img" aria-label="${nCol} by ${gCol}"></canvas>
        </div>
        ${insightDiv(sorted.map(e=>e[0]), sorted.map(e=>Math.round(e[1]*100)/100))}`;
      grid.appendChild(card);
      setTimeout(() => createChart(id, 'bar', sorted.map(e=>e[0]), sorted.map(e=>Math.round(e[1]*100)/100), `${nCol} według ${gCol}`, isH, true, gCol), 50);
    }
  }

  // Second numeric by first cat
  if (numCols.length > 1 && catCols.length) {
    const gCol = catCols[0], nCol = numCols[1];
    const groups = {};
    allData.forEach(r => {
      const k = String(r[gCol]) || '(empty)';
      const v = toNum(r[nCol]);
      if (v !== null) groups[k] = (groups[k]||0) + v;
    });
    const sorted = Object.entries(groups).sort((a,b) => b[1]-a[1]).slice(0, 12);
    if (sorted.length > 1) {
      const id = `ac_${chartIdx++}`;
      const isH = sorted.length > 5;
      const h = isH ? Math.max(180, sorted.length * 30) : 220;
      const card = document.createElement('div');
      card.className = 'chart-card';
      card.innerHTML = `<div class="chart-card-title">${nCol} według ${gCol}</div>
        <div class="chart-card-sub">Oś X: „${gCol}” · Oś Y: suma „${nCol}” — czyli łączny wynik pracy (np. sztuk towaru) dla tej wartości, a NIE liczba ${getRowEntityLabel()}</div>
        <div style="position:relative;width:100%;height:${h}px;">
          <canvas id="${id}" role="img" aria-label="${nCol} by ${gCol}"></canvas>
        </div>
        ${insightDiv(sorted.map(e=>e[0]), sorted.map(e=>Math.round(e[1]*100)/100))}`;
      grid.appendChild(card);
      setTimeout(() => createChart(id, 'bar', sorted.map(e=>e[0]), sorted.map(e=>Math.round(e[1]*100)/100), `${nCol} według ${gCol}`, isH, true, gCol), 50);
    }
  }

  // Date trend
  if (dateCols.length && numCols.length) {
    const dCol = dateCols[0], nCol = numCols[0];
    const byDate = {};
    allData.forEach(r => {
      const d = String(r[dCol]).slice(0, 10);
      if (!d || d === '') return;
      const v = toNum(r[nCol]);
      if (v !== null) byDate[d] = (byDate[d]||0) + v;
    });
    const sorted = Object.entries(byDate).sort((a,b) => a[0].localeCompare(b[0])).slice(0, 60);
    if (sorted.length > 1) {
      const id = `ac_${chartIdx++}`;
      const card = document.createElement('div');
      card.className = 'chart-card';
      card.style.gridColumn = 'span 2';
      card.innerHTML = `<div class="chart-card-title">Trend: ${nCol} w czasie</div>
        <div class="chart-card-sub">Oś X: data („${dCol}”) · Oś Y: suma „${nCol}” danego dnia — czyli ile wyniosła praca/produkcja każdego dnia</div>
        <div style="position:relative;width:100%;height:200px;">
          <canvas id="${id}" role="img" aria-label="Date Trend"></canvas>
        </div>
        ${insightDiv(sorted.map(e=>e[0]), sorted.map(e=>Math.round(e[1]*100)/100), {trend:true})}`;
      grid.appendChild(card);
      setTimeout(() => createChart(id, 'line', sorted.map(e=>e[0]), sorted.map(e=>Math.round(e[1]*100)/100), `${nCol}`, false, true, dCol), 50);
    }
  }

  if (!grid.children.length) {
    grid.innerHTML = '<div class="empty">Not enough categorical or numeric columns for auto-charts.<br>Go to the <strong>Charts</strong> tab for manual setup.</div>';
  }
}

// ─── OUTBOUND PACK ONLINE METRICS ──────────────────────────────
function extractHour(v) {
  if (v == null || v === '') return null;
  // Try plain "HH:mm" or "HH:mm:ss" style strings first, with optional AM/PM
  const s = String(v).trim();
  const hm = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM|am|pm)?/);
  if (hm) {
    let h = parseInt(hm[1], 10);
    const ampm = hm[3] ? hm[3].toUpperCase() : null;
    if (ampm === 'PM' && h < 12) h += 12;
    if (ampm === 'AM' && h === 12) h = 0;
    if (h >= 0 && h <= 23) return h;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.getHours();
  return null;
}

function findCol(predRegex, fallbackTypes) {
  let match = cols.find(c => predRegex.test(c));
  if (match) return match;
  if (fallbackTypes) return cols.find(c => fallbackTypes.includes(colTypes[c]));
  return null;
}

function renderOutboundSection() {
  const allData = getScopedData();
  const section = document.getElementById('outbound-section');
  const kpiGrid = document.getElementById('outbound-kpi-grid');
  const chartWrap = document.getElementById('hourly-productivity-chart')?.closest('.chart-card');
  if (!section || !kpiGrid) return;

  if (chartInstances['outbound_hourly']) {
    try { chartInstances['outbound_hourly'].destroy(); } catch(e){}
    delete chartInstances['outbound_hourly'];
  }

  const numCols = cols.filter(c => colTypes[c] === 'numeric');
  const dateCols = cols.filter(c => colTypes[c] === 'date');
  const timeCol = findCol(/time|hour|hora|godzina/i) || dateCols[0] || null;
  const qtyCol = findCol(/qty|quantity|units?|pieces|pcs|szt|liczba/i, ['numeric']) || numCols[0] || null;
  const orderCol = findCol(/order|pedido|zamow|shipment|delivery|wysyłk/i);

  if (!qtyCol) {
    kpiGrid.innerHTML = '<div class="empty">No numeric column found to compute outbound metrics.</div>';
    if (chartWrap) chartWrap.style.display = 'none';
    return;
  }

  const vals = allData.map(r => toNum(r[qtyCol])).filter(v => v !== null);
  const totalUnits = vals.reduce((a,b) => a+b, 0);
  const totalRows = allData.length;
  const totalOrders = orderCol ? new Set(allData.map(r => String(r[orderCol]))).size : totalRows;

  // Hourly breakdown, if we have any parseable time info
  const byHour = {};
  if (timeCol) {
    allData.forEach(r => {
      const h = extractHour(r[timeCol]);
      const v = toNum(r[qtyCol]);
      if (h !== null && v !== null) byHour[h] = (byHour[h] || 0) + v;
    });
  }
  const activeHours = Object.keys(byHour).length;
  const avgPerHour = activeHours ? totalUnits / activeHours : null;

  const accents = ['#3b82f6','#10b981','#f59e0b','#ef4444'];
  const cards = [
    { label: `Total ${qtyCol}`, value: fmt(totalUnits) },
    { label: orderCol ? `Distinct ${orderCol}` : 'Total Rows', value: fmt(totalOrders) },
    { label: 'Active Hours', value: activeHours ? activeHours : '—' },
    { label: 'Avg per Active Hour', value: avgPerHour != null ? fmt(avgPerHour) : '—' },
  ];
  kpiGrid.innerHTML = cards.map((c, i) => `
    <div class="kpi-card" style="--accent:${accents[i % accents.length]}">
      <div class="kpi-label">${c.label}</div>
      <div class="kpi-value" style="color:${accents[i % accents.length]}">${c.value}</div>
    </div>`).join('');

  if (!chartWrap) return;
  if (!timeCol || activeHours < 2) {
    chartWrap.innerHTML = `<div class="chart-card-title">Suma „${qtyCol}” w podziale na godziny</div>
      <div class="empty">Nie znaleziono kolumny z godziną/czasem — nie można zbudować wykresu godzinowego.</div>`;
    return;
  }
  chartWrap.innerHTML = `<div class="chart-card-title">Suma „${qtyCol}” w podziale na godziny</div>
    <div class="chart-card-sub">Oś X: godzina · Oś Y: suma „${qtyCol}” w tej godzinie — czyli wynik pracy godzina po godzinie</div>
    <div style="position:relative;width:100%;height:280px;">
      <canvas id="hourly-productivity-chart"></canvas>
    </div>
    <div id="hourly-productivity-insight"></div>`;
  const hours = Object.keys(byHour).map(Number).sort((a,b) => a-b);
  const labels = hours.map(h => `${String(h).padStart(2,'0')}:00`);
  const data = hours.map(h => Math.round(byHour[h] * 100) / 100);
  const insightHost = document.getElementById('hourly-productivity-insight');
  if (insightHost) insightHost.innerHTML = insightDiv(labels, data, { trend: true });

  const canvas = document.getElementById('hourly-productivity-chart');
  chartInstances['outbound_hourly'] = new Chart(canvas, {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: `${qtyCol} per hour`,
        data,
        borderColor: '#3b82f6',
        backgroundColor: 'rgba(59,130,246,0.1)',
        fill: true,
        tension: 0.3,
        pointRadius: 3,
        borderWidth: 2,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => ' ' + fmt(ctx.parsed.y) } }
      },
      scales: {
        x: { grid: { display: false }, ticks: { font: { size: 11 } } },
        y: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 }, callback: v => fmt(v) } }
      }
    }
  });
}

// ─── DYNAMIC OPERATIONAL INSIGHTS ──────────────────────────────
// Replaces the old enhancements.js + accountability.js add-ons.
// Single source of truth: every number here is computed once and
// reused everywhere, so worker totals / productivity never disagree
// between sections. Group columns (brand, error reason, etc.) are
// detected by data shape (low cardinality), not by hardcoded names,
// because the real warehouse files (Fortna WES / SAP) aren't live yet
// and column naming isn't fixed.

let productivityTarget = 100; // default; adjustable via the input in the UI

function renderInsightsSection() {
  const allData = getScopedData();
  const host = document.getElementById('insights-section');
  if (!host) return;
  host.innerHTML = ''; // full rebuild every time -> no leftover blocks on re-upload

  const userCol = findCol(/user|operator|worker|packer|employee|pracownik/i);
  const qtyCol = findCol(/qty|quantity|units?|pieces|pcs|szt|liczba/i, ['numeric']);
  const timeCol = findCol(/time|hour|godzina/i) || cols.filter(c => colTypes[c] === 'date')[0] || null;
  const orderCol = findCol(/order|zamow|shipment|delivery|wysyłk/i);

  if (!qtyCol) return; // nothing numeric to build insights from

  // ── Worker productivity (single computation, used by table + charts + audit) ──
  let workerStats = null;
  if (userCol) {
    workerStats = {};
    allData.forEach(r => {
      const user = r[userCol] || 'Unknown';
      const qty = toNum(r[qtyCol]);
      if (qty == null) return;
      if (!workerStats[user]) workerStats[user] = { units: 0, rows: 0, orders: new Set() };
      workerStats[user].units += qty;
      workerStats[user].rows++;
      if (orderCol && r[orderCol]) workerStats[user].orders.add(r[orderCol]);
    });
  }

  // ── Dynamic group columns: any categorical column not already used
  // as user/time/order, ranked with a soft preference for likely
  // "brand" / "reason/error" semantics, but generic otherwise. ──
  const usedCols = new Set([userCol, timeCol, orderCol].filter(Boolean));
  const candidateGroups = cols.filter(c =>
    !usedCols.has(c) && c !== qtyCol &&
    (colTypes[c] === 'categorical') &&
    new Set(allData.map(r => String(r[c] ?? ''))).size >= 2
  );
  // Includes Fortna WES-specific terminology confirmed from the SRS:
  // exceptions go to "Hospital" (not "error"), missing items are "short",
  // brand field in the pick request is literally named RBAND.
  const priorityRe = /brand|rband|marka|brend|reason|error|fault|issue|awaria|przyczyna|kod|status|defect|reklamacj|hospital|short|damage|uszkodz/i;
  candidateGroups.sort((a, b) => (priorityRe.test(b) ? 1 : 0) - (priorityRe.test(a) ? 1 : 0));

  const block = document.createElement('div');

  // Section header + adjustable target
  const header = document.createElement('div');
  header.className = 'col-section';
  header.innerHTML = `
    <div class="col-section-label" style="display:flex;align-items:center;gap:10px;">
      Operational Insights
      <span style="display:flex;align-items:center;gap:6px;font-weight:500;text-transform:none;letter-spacing:0;color:var(--ink-dim);">
        · target/unit: <input type="number" id="prod-target-input" value="${productivityTarget}" style="width:70px;padding:2px 6px;border:1px solid var(--line);border-radius:5px;font-size:12px;">
      </span>
    </div>`;
  block.appendChild(header);
  setTimeout(() => {
    const inp = document.getElementById('prod-target-input');
    if (inp) inp.addEventListener('change', () => {
      productivityTarget = toNum(inp.value) || 100;
      renderInsightsSection();
    });
  }, 0);

  // Audit summary cards
  if (workerStats) {
    const totalUnits = Object.values(workerStats).reduce((s, w) => s + w.units, 0);
    const totalRows = Object.values(workerStats).reduce((s, w) => s + w.rows, 0);
    const avg = totalRows ? totalUnits / totalRows : 0;
    const best = Object.entries(workerStats).sort((a, b) => b[1].units - a[1].units)[0];
    const worst = Object.entries(workerStats).sort((a, b) => a[1].units - b[1].units)[0];

    const summary = document.createElement('div');
    summary.className = 'kpi-grid';
    summary.style.marginBottom = '16px';
    const accents = ['#3b82f6', avg >= productivityTarget ? '#10b981' : '#ef4444', '#f59e0b', '#8b5cf6'];
    summary.innerHTML = `
      <div class="kpi-card" style="--accent:${accents[0]}"><div class="kpi-label">Total Units</div><div class="kpi-value" style="color:${accents[0]}">${fmt(totalUnits)}</div></div>
      <div class="kpi-card" style="--accent:${accents[1]}"><div class="kpi-label">Avg Productivity / row</div><div class="kpi-value" style="color:${accents[1]}">${avg.toFixed(1)}</div><div class="kpi-meta">Target: ${productivityTarget}</div></div>
      <div class="kpi-card" style="--accent:${accents[2]}"><div class="kpi-label">🏆 Top Worker</div><div class="kpi-value" style="font-size:16px;color:${accents[2]}">${best ? best[0] : '—'}</div></div>
      <div class="kpi-card" style="--accent:${accents[3]}"><div class="kpi-label">⚠️ Needs Attention</div><div class="kpi-value" style="font-size:16px;color:${accents[3]}">${worst ? worst[0] : '—'}</div></div>
    `;
    block.appendChild(summary);

    if (avg < productivityTarget) {
      const alert = document.createElement('div');
      alert.style.background = 'rgba(217,96,92,.15)';
      alert.style.color = 'var(--red)';
      alert.style.border = '1px solid rgba(217,96,92,.35)';
      alert.style.borderRadius = '8px';
      alert.style.padding = '8px 14px';
      alert.style.fontSize = '12px';
      alert.style.marginBottom = '16px';
      alert.textContent = `⚠️ Average productivity (${avg.toFixed(1)}) is below target (${productivityTarget}).`;
      block.appendChild(alert);
    }
  }

  // Worker table
  if (workerStats) {
    const sorted = Object.entries(workerStats)
      .map(([name, d]) => ({ name, units: d.units, orders: d.orders.size, productivity: d.rows ? d.units / d.rows : 0 }))
      .sort((a, b) => b.units - a.units);

    const tableWrap = document.createElement('div');
    tableWrap.className = 'col-section';
    tableWrap.innerHTML = `
      <div class="col-section-label">Worker Breakdown (${userCol})</div>
      <div class="table-wrap" style="max-height:360px;">
        <table>
          <thead><tr><th>Name</th><th>Units</th>${orderCol ? '<th>Orders</th>' : ''}<th>Avg / row</th></tr></thead>
          <tbody>
            ${sorted.slice(0, 25).map(w => `
              <tr>
                <td>${w.name}</td>
                <td>${fmt(w.units)}</td>
                ${orderCol ? `<td>${w.orders}</td>` : ''}
                <td style="color:${w.productivity >= productivityTarget ? 'var(--green)' : 'var(--red)'}">${w.productivity.toFixed(1)}</td>
              </tr>`).join('')}
          </tbody>
        </table>
      </div>`;
    block.appendChild(tableWrap);

    // Top 5 / Bottom 5 — explicitly calling out who's leading and who
    // might need support, side by side, automatically (no chart-building
    // required). This is the one thing the old enhancements.js add-on did
    // that the generic breakdown below doesn't: a low-cardinality "Top N"
    // chart only ever shows leaders, never flags who's struggling.
    if (sorted.length >= 4) {
      const top5 = sorted.slice(0, 5);
      const low5 = sorted.slice(-5).reverse();
      const chartsRow = document.createElement('div');
      chartsRow.className = 'charts-auto-grid';
      chartsRow.style.marginTop = '14px';
      const topId = 'wk_top', lowId = 'wk_low';
      chartsRow.innerHTML = `
        <div class="chart-card">
          <div class="chart-card-title">🏆 Top 5 — ${userCol}</div>
          <div class="chart-card-sub">Najwyższa suma „${qtyCol}”</div>
          <div style="position:relative;width:100%;height:220px;"><canvas id="${topId}"></canvas></div>
        </div>
        <div class="chart-card">
          <div class="chart-card-title">⚠️ Najsłabsze 5 — ${userCol}</div>
          <div class="chart-card-sub">Najniższa suma „${qtyCol}” — kandydaci do wsparcia/coachingu</div>
          <div style="position:relative;width:100%;height:220px;"><canvas id="${lowId}"></canvas></div>
        </div>`;
      block.appendChild(chartsRow);
      setTimeout(() => {
        createChart(topId, 'bar', top5.map(w => w.name), top5.map(w => Math.round(w.units * 100) / 100), 'Top 5', true, true, userCol);
        createChart(lowId, 'bar', low5.map(w => w.name), low5.map(w => Math.round(w.units * 100) / 100), 'Najsłabsze 5', true, true, userCol);
      }, 50);
    }
  }

  // Dynamic category breakdown charts (brand / error-reason / anything else low-cardinality)
  if (candidateGroups.length) {
    const chartsWrap = document.createElement('div');
    chartsWrap.className = 'col-section';
    chartsWrap.innerHTML = `<div class="col-section-label">Breakdown by Detected Categories</div>`;
    const grid = document.createElement('div');
    grid.className = 'charts-auto-grid';
    chartsWrap.appendChild(grid);
    block.appendChild(chartsWrap);

    candidateGroups.slice(0, 4).forEach((gCol, idx) => {
      const groups = {};
      allData.forEach(r => {
        const k = String(r[gCol] || '(empty)');
        const v = toNum(r[qtyCol]);
        if (v !== null) groups[k] = (groups[k] || 0) + v;
      });
      const sorted = Object.entries(groups).sort((a, b) => b[1] - a[1]).slice(0, 12);
      if (sorted.length < 2) return;
      const id = `insight_${idx}`;
      const isH = sorted.length > 6;
      const h = isH ? Math.max(180, sorted.length * 28) : 220;
      const card = document.createElement('div');
      card.className = 'chart-card';
      card.innerHTML = `<div class="chart-card-title">${qtyCol} według ${gCol}</div>
        <div class="chart-card-sub">Oś X: „${gCol}” · Oś Y: suma „${qtyCol}” — wynik pracy dla tej wartości, top ${sorted.length}</div>
        <div style="position:relative;width:100%;height:${h}px;"><canvas id="${id}"></canvas></div>
        ${insightDiv(sorted.map(e => e[0]), sorted.map(e => Math.round(e[1] * 100) / 100))}`;
      grid.appendChild(card);
      setTimeout(() => createChart(id, 'bar', sorted.map(e => e[0]), sorted.map(e => Math.round(e[1] * 100) / 100), gCol, isH, true, gCol), 50);
    });
  }

  host.appendChild(block);
}

// ─── DRILL-DOWN ─────────────────────────────────────────────
// Click a bar/slice/pivot cell -> jump to the Table tab pre-filtered
// to exactly those rows. secondaryQuery additionally narrows by a
// second value (used by pivot cells, which have a row AND a column
// dimension but the Table tab only has one structured filter slot).
function drillToTable(colName, value, secondaryQuery) {
  const tabs = document.querySelectorAll('.tab');
  const tableTab = Array.from(tabs).find(t => t.getAttribute('onclick')?.includes("'table'"));
  if (tableTab) switchTab('table', tableTab);

  const fc = document.getElementById('filter-col');
  const hasCol = Array.from(fc.options).some(o => o.value === colName);
  fc.value = hasCol ? colName : '';
  fc.dispatchEvent(new Event('change'));

  setTimeout(() => {
    if (hasCol) {
      const fv = document.getElementById('filter-val');
      const match = Array.from(fv.options).find(o => o.value === String(value));
      if (match) fv.value = match.value;
    }
    document.getElementById('search-input').value = secondaryQuery || (hasCol ? '' : String(value ?? ''));
    filterTable();
  }, 0);
}

// Unit suffix purely from the column NAME (works even when values are
// already plain numbers, e.g. "Czas_reakcji_min" = 112 meaning 112
// minutes — isDurationColumn() alone can't catch this since the raw
// values aren't HH:MM strings).
function getColumnUnit(col) {
  if (isDurationColumn(col)) return 'min';
  if (/_min\b|minut/i.test(col)) return 'min';
  if (/_godz\b|godzin/i.test(col)) return 'godz';
  if (/_sek\b|sekund/i.test(col)) return 'sek';
  return null;
}

function createChart(id, type, labels, data, label, isHBar = false, isNumY = false, drillCol = null, unitSuffix = null) {
  const canvas = document.getElementById(id);
  if (!canvas) return;
  // Always destroy any previous Chart.js instance attached to this same
  // canvas first. Without this, switching configurations rapidly (e.g.
  // moving a field from "Oś Y" to "Podział kolorami" and back) leaves
  // multiple live chart instances drawing on top of each other on the
  // same <canvas>, producing corrupted/gappy renders that only resolve
  // on a full page reload.
  if (chartInstances[id]) { try { chartInstances[id].destroy(); } catch (e) {} delete chartInstances[id]; }
  const isDoughnut = type === 'doughnut' || type === 'pie';
  const bgColors = isDoughnut ? labels.map((_, i) => COLORS[i % COLORS.length]) : COLORS[0];
  const fmtU = v => fmt(v) + (unitSuffix ? ' ' + unitSuffix : '');

  const cfg = {
    type: isDoughnut ? type : 'bar',
    data: {
      labels,
      datasets: [{
        label,
        data,
        backgroundColor: bgColors,
        borderWidth: 0,
        borderRadius: isDoughnut ? 0 : 3,
        tension: type === 'line' ? 0.3 : undefined,
        borderColor: type === 'line' ? COLORS[0] : undefined,
        fill: type === 'line' ? false : undefined,
        pointRadius: type === 'line' ? 3 : undefined,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: isHBar ? 'y' : 'x',
      onClick: drillCol ? (evt, elements) => {
        if (!elements.length) return;
        const idx = elements[0].index;
        drillToTable(drillCol, labels[idx]);
      } : undefined,
      onHover: drillCol ? (evt, elements) => { evt.native.target.style.cursor = elements.length ? 'pointer' : 'default'; } : undefined,
      plugins: {
        legend: { display: isDoughnut },
        tooltip: { callbacks: { label: ctx => ' ' + fmtU(isDoughnut ? ctx.parsed : ctx.parsed[isHBar ? 'x' : 'y']) } }
      },
      scales: isDoughnut ? {} : {
        [isHBar ? 'x' : 'y']: {
          grid: { color: 'rgba(0,0,0,0.05)' },
          ticks: { font: { size: 11 }, callback: v => isNumY ? fmtU(v) : v, precision: 0 }
        },
        [isHBar ? 'y' : 'x']: {
          grid: { display: false },
          ticks: { font: { size: 10 }, maxRotation: isHBar ? 0 : 45 }
        }
      }
    }
  };

  if (type === 'line') {
    cfg.type = 'line';
    cfg.options.scales = {
      x: { grid: { display: false }, ticks: { font: { size: 10 }, maxRotation: 45, autoSkip: true, maxTicksLimit: 12 } },
      y: { grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 }, callback: v => fmt(v) } }
    };
  }

  const instance = new Chart(canvas, cfg);
  chartInstances[id] = instance;
}

// ─── BUILDER SELECTS ──────────────────────────────────────────
const TYPE_LABELS_PL = { numeric: 'liczbowa', categorical: 'kategoria', date: 'data', time: 'czas', text: 'tekst', unknown: '?' };

// ─── EXCEL-STYLE PIVOT BUILDER ──────────────────────────────────
// This is the ONLY visible way to build a view now. Everything else
// (tiles, headline, relationship explorer, manual axis/agg builder)
// stays defined above as working functions — kept in code on purpose,
// just not wired into the UI — in case we want to bring any of it back.
let pivotState = { filters: [], columns: [], rows: [], values: [] };

function renderPivotFieldList() {
  const host = document.getElementById('pivot-field-list');
  if (!host) return;
  const icon = { numeric: '🔢', categorical: '🏷️', date: '📅', time: '🕐', text: '🔤' };
  host.innerHTML = cols.map(c => `
    <div class="field-chip" draggable="true" data-col="${escAttr(c)}" ondragstart="onFieldDragStart(event)">
      <span>${icon[colTypes[c]] || '❔'}</span><span>${escAttr(c)}</span>
    </div>`).join('');
}

function onFieldDragStart(event) {
  event.dataTransfer.setData('text/plain', event.currentTarget.dataset.col);
}

function onFieldDrop(event, zone) {
  event.preventDefault();
  event.currentTarget.classList.remove('drag-over');
  const col = event.dataTransfer.getData('text/plain');
  if (!col || !cols.includes(col)) return;
  if (zone === 'rows' || zone === 'columns') {
    pivotState[zone] = [{ col }]; // single field, like a simplified Excel pivot
  } else if (zone === 'values') {
    if (colTypes[col] !== 'numeric') {
      // A text/category field dropped on "Oś Y" can't be summed/averaged —
      // forcing it to "LICZBA" would silently ignore its actual values
      // (just counting rows). What the person almost always wants instead
      // is to see that field's breakdown, which is exactly what "Podział
      // kolorami" does — so redirect it there instead of giving a
      // technically-valid but meaningless count.
      if (!pivotState.columns.some(f => f.col === col)) pivotState.columns = [{ col }];
      showBuilderValidation(`„${col}” to kolumna tekstowa, więc przeniesiono ją do „Podział kolorami” — tam jej wartości faktycznie się pokażą (w „Oś Y” zliczałaby tylko wiersze, ignorując treść).`);
    } else {
      if (pivotState.values.some(v => v.col === col)) return;
      pivotState.values.push({ col, agg: 'sum' });
    }
  } else if (zone === 'filters') {
    if (pivotState.filters.some(f => f.col === col)) return;
    const allVals = [...new Set(allData.map(r => String(r[col] ?? '')))];
    pivotState.filters.push({ col, excluded: new Set() }); // nothing excluded by default = show all
  }
  renderPivotZones();
  renderExcelPivot();
}

const AGG_CYCLE = ['sum', 'count', 'avg', 'max', 'min'];
const AGG_SHORT = { sum: 'SUMA', count: 'LICZBA', avg: 'ŚREDNIA', max: 'MAX', min: 'MIN' };

function cycleValueAgg(idx) {
  const v = pivotState.values[idx];
  const next = AGG_CYCLE[(AGG_CYCLE.indexOf(v.agg) + 1) % AGG_CYCLE.length];
  v.agg = next;
  renderPivotZones();
  renderExcelPivot();
}
function removeFromZone(zone, idx) {
  pivotState[zone].splice(idx, 1);
  renderPivotZones();
  renderExcelPivot();
}

function renderPivotZones() {
  document.getElementById('zone-rows').innerHTML = pivotState.rows.map((f, i) => `
    <div class="field-chip in-zone"><span>${escAttr(f.col)}</span><span class="chip-remove" onclick="removeFromZone('rows',${i})">✕</span></div>`).join('') || '';
  document.getElementById('zone-columns').innerHTML = pivotState.columns.map((f, i) => `
    <div class="field-chip in-zone"><span>${escAttr(f.col)}</span><span class="chip-remove" onclick="removeFromZone('columns',${i})">✕</span></div>`).join('') || '';
  document.getElementById('zone-values').innerHTML = pivotState.values.map((f, i) => `
    <div class="field-chip in-zone"><span>${escAttr(f.col)}<span class="chip-agg" onclick="cycleValueAgg(${i})">${AGG_SHORT[f.agg]}</span></span><span class="chip-remove" onclick="removeFromZone('values',${i})">✕</span></div>`).join('') || '';
  document.getElementById('zone-filters').innerHTML = pivotState.filters.map((f, i) => `
    <div class="field-chip in-zone" onclick="openFilterPopup(event,${i})"><span>${escAttr(f.col)}${f.excluded.size ? ` (${f.excluded.size} wykl.)` : ''}</span><span class="chip-remove" onclick="event.stopPropagation();removeFromZone('filters',${i})">✕</span></div>`).join('') || '';
}

function openFilterPopup(evt, idx) {
  const f = pivotState.filters[idx];
  const popup = document.getElementById('pivot-filter-popup');
  const values = [...new Set(allData.map(r => String(r[f.col] ?? '')))].sort();
  popup.innerHTML = `
    <div style="font-size:11px;font-weight:700;color:var(--ink-dim);margin-bottom:6px;text-transform:uppercase;">${escAttr(f.col)}</div>
    ${values.map(v => `<label style="display:flex;align-items:center;gap:6px;font-size:12px;padding:3px 0;cursor:pointer;">
      <input type="checkbox" style="width:auto;" ${f.excluded.has(v) ? '' : 'checked'} onchange="toggleFilterValue(${idx},${JSON.stringify(v)},this.checked)"> ${escAttr(v) || '(puste)'}
    </label>`).join('')}
    <button class="btn btn-ghost btn-sm" style="margin-top:8px;width:100%;" onclick="document.getElementById('pivot-filter-popup').style.display='none'">Zamknij</button>`;
  const rect = evt.currentTarget.getBoundingClientRect();
  popup.style.left = rect.left + 'px';
  popup.style.top = (rect.bottom + 4) + 'px';
  popup.style.display = 'block';
}
function toggleFilterValue(idx, val, checked) {
  const f = pivotState.filters[idx];
  if (checked) f.excluded.delete(val); else f.excluded.add(val);
  renderExcelPivot();
}
document.addEventListener('click', e => {
  const popup = document.getElementById('pivot-filter-popup');
  if (popup && popup.style.display === 'block' && !popup.contains(e.target) && !e.target.closest('#zone-filters')) popup.style.display = 'none';
});

function clearPivotZones() {
  pivotState = { filters: [], columns: [], rows: [], values: [] };
  renderPivotZones();
  renderExcelPivot();
}

function renderExcelPivot() {
  const tableHost = document.getElementById('pivot-table-wrap');
  const chartWrap = document.getElementById('pivot-chart-wrap');
  if (!tableHost) return;

  // Safety net independent of the drop event: if the state somehow ended
  // up with a non-numeric field in "Wartości" (e.g. state set before this
  // check existed, or restored some other way), fix it here too — not
  // only at the moment of dropping — so it can never silently stick.
  const badValueFields = pivotState.values.filter(v => colTypes[v.col] !== 'numeric');
  if (badValueFields.length) {
    badValueFields.forEach(v => { if (!pivotState.columns.some(f => f.col === v.col)) pivotState.columns = [{ col: v.col }]; });
    pivotState.values = pivotState.values.filter(v => colTypes[v.col] === 'numeric');
    renderPivotZones();
  }

  let data = getScopedData();
  pivotState.filters.forEach(f => { if (f.excluded.size) data = data.filter(r => !f.excluded.has(String(r[f.col] ?? ''))); });

  if (!pivotState.rows.length && !pivotState.columns.length) {
    tableHost.innerHTML = '<div class="empty">Przeciągnij pola po lewej do obszarów powyżej, aby zbudować tabelę.</div>';
    chartWrap.style.display = 'none';
    return;
  }
  const rowCol0 = pivotState.rows[0]?.col || null;
  const colCol0 = pivotState.columns[0]?.col || null;
  // If nothing was dropped into "Wartości", default to counting rows —
  // the most common pivot use case (e.g. "Status × Typ_problemu: ile
  // przypadków każdej kombinacji") shouldn't require an extra manual step.
  const values = pivotState.values.length ? pivotState.values : [{ col: (rowCol0 || colCol0), agg: 'count' }];
  const rowCol = rowCol0, colCol = colCol0;

  const rawByCell = {}; // "row||col" -> {valueCol: [numbers]}
  const rowSet = new Set(), colSet = new Set();
  data.forEach(r => {
    const rv = rowCol ? String(r[rowCol] ?? '(puste)') : '__all__';
    const cv = colCol ? String(r[colCol] ?? '(puste)') : '__all__';
    rowSet.add(rv); colSet.add(cv);
    const key = rv + '||' + cv;
    if (!rawByCell[key]) rawByCell[key] = {};
    values.forEach(v => {
      const num = v.agg === 'count' ? 1 : toNum(r[v.col]);
      if (num === null) return;
      (rawByCell[key][v.col] = rawByCell[key][v.col] || []).push(num);
    });
  });

  const MAX_ROWS = 500, MAX_COLS = 12;
  let rowVals = [...rowSet];
  let colVals = [...colSet];
  // sort rows by total of first value field, descending
  const rowTotal = {};
  rowVals.forEach(rv => {
    let t = 0;
    colVals.forEach(cv => { const cell = rawByCell[rv + '||' + cv]; if (cell && cell[values[0].col]) t += aggregateValues(cell[values[0].col], values[0].agg === 'count' ? 'sum' : values[0].agg); });
    rowTotal[rv] = t;
  });
  rowVals.sort((a, b) => rowTotal[b] - rowTotal[a]);
  const truncatedRows = rowVals.length > MAX_ROWS;
  if (truncatedRows) rowVals = rowVals.slice(0, MAX_ROWS);
  const truncatedCols = colVals.length > MAX_COLS;
  if (truncatedCols) colVals = colVals.slice(0, MAX_COLS);

  // build header: for each colVal, one sub-column per value-field (Excel does this when >1 value field)
  const subCols = colVals.length ? colVals : ['__all__'];
  const headerRow2 = subCols.map(cv => values.map(v => `${AGG_SHORT[v.agg]} ${escAttr(v.agg === 'count' ? getRowEntityLabel() : v.col)}`).join('</th><th style="padding:6px 10px;text-align:right;font-size:10.5px;color:var(--ink-dim);border-bottom:1px solid var(--line);">')).join('</th><th style="padding:6px 10px;text-align:right;font-size:10.5px;color:var(--ink-dim);border-bottom:1px solid var(--line);border-left:2px solid var(--line);">');

  let html = `<table style="width:100%;border-collapse:collapse;font-size:12.5px;">
    <thead>
      <tr>
        <th rowspan="2" style="position:sticky;left:0;top:0;background:var(--panel-2);padding:8px 10px;text-align:left;border-bottom:1px solid var(--line);z-index:3;">${escAttr(rowCol || '')}</th>
        ${subCols.map(cv => `<th colspan="${values.length}" style="padding:6px 10px;text-align:center;background:var(--panel-2);border-bottom:1px solid var(--line);position:sticky;top:0;z-index:2;border-left:2px solid var(--line);">${colCol ? escAttr(cv) : 'Wartość'}</th>`).join('')}
      </tr>
      <tr>${subCols.map(() => values.map(v => `<th style="padding:6px 10px;text-align:right;font-size:10.5px;color:var(--ink-dim);border-bottom:1px solid var(--line);position:sticky;top:26px;background:var(--panel-2);z-index:2;">${AGG_SHORT[v.agg]} ${escAttr(v.agg === 'count' ? getRowEntityLabel() : v.col)}</th>`).join('')).join('<th style="border-left:2px solid var(--line);width:0;padding:0;"></th>')}</tr>
    </thead>
    <tbody>`;

  rowVals.forEach(rv => {
    html += `<tr><td style="position:sticky;left:0;background:var(--panel);padding:7px 10px;border-bottom:1px solid var(--line);font-weight:600;">${escAttr(rv)}</td>`;
    subCols.forEach((cv, ci) => {
      const cell = rawByCell[rv + '||' + cv] || {};
      values.forEach(v => {
        const arr = cell[v.col] || [];
        const val = arr.length ? aggregateValues(arr, v.agg === 'count' ? 'sum' : v.agg) : null;
        const unit = v.agg === 'count' ? '' : (getColumnUnit(v.col) ? ' ' + getColumnUnit(v.col) : '');
        html += `<td style="padding:7px 10px;text-align:right;border-bottom:1px solid var(--line);${ci > 0 ? 'border-left:2px solid var(--line);' : ''}">${val == null ? '—' : fmt(val) + unit}</td>`;
      });
    });
    html += `</tr>`;
  });
  html += `</tbody></table>`;
  if (truncatedRows || truncatedCols) html += `<div style="padding:8px 10px;font-size:11px;color:var(--ink-dim);">Pokazano tylko najliczniejsze kategorie (limit: ${MAX_ROWS} wierszy × ${MAX_COLS} kolumn).</div>`;
  tableHost.innerHTML = html;

  // optional chart: first value field, by rows (and stacked by column field if present)
  const showChart = document.getElementById('pivot-show-chart')?.checked;
  if (showChart) {
    chartWrap.style.display = 'block';
    const v0 = values[0];
    const v0Label = v0.agg === 'count' ? `Liczba ${getRowEntityLabel()}` : `${AGG_SHORT[v0.agg]} ${v0.col}`;
    if (colCol && colVals.length > 1) {
      const datasets = colVals.map((cv, i) => ({
        label: cv,
        data: rowVals.map(rv => { const cell = rawByCell[rv + '||' + cv]; const arr = cell?.[v0.col] || []; return arr.length ? Math.round(aggregateValues(arr, v0.agg === 'count' ? 'sum' : v0.agg) * 100) / 100 : 0; }),
        backgroundColor: COLORS[i % COLORS.length],
      }));
      setTimeout(() => {
        if (chartInstances['pivot-chart']) { try { chartInstances['pivot-chart'].destroy(); } catch (e) {} }
        const canvasEl = document.getElementById('pivot-chart');
        if (canvasEl) canvasEl.style.width = Math.max(rowVals.length * 42, 200) + 'px';
        chartInstances['pivot-chart'] = new Chart(document.getElementById('pivot-chart'), {
          type: 'bar', data: { labels: rowVals, datasets },
          options: { responsive: true, maintainAspectRatio: false, indexAxis: 'x',
            scales: { x: { stacked: true, ticks: { font: { size: 10 } } }, y: { stacked: true, ticks: { font: { size: 10 } } } },
            plugins: { legend: { display: true, position: 'bottom', labels: { font: { size: 10 } } } } } });
      }, 30);
    } else {
      const chartData = rowVals.map(rv => { const cell = rawByCell[rv + '||__all__'] || rawByCell[rv + '||' + colVals[0]]; const arr = cell?.[v0.col] || []; return arr.length ? Math.round(aggregateValues(arr, v0.agg === 'count' ? 'sum' : v0.agg) * 100) / 100 : 0; });
      const canvasEl = document.getElementById('pivot-chart');
      if (canvasEl) canvasEl.style.width = Math.max(rowVals.length * 42, 200) + 'px';
      setTimeout(() => createChart('pivot-chart', 'bar', rowVals, chartData, v0Label, false, true, rowCol), 30);
    }
  } else {
    chartWrap.style.display = 'none';
  }
}

function renderBuilderSelects() {
  const bx = document.getElementById('b-x');
  const by = document.getElementById('b-y');
  const bs = document.getElementById('b-stack');
  bx.innerHTML = cols.map(c => `<option value="${c}">${c} [${TYPE_LABELS_PL[colTypes[c]] || colTypes[c]}]</option>`).join('');
  by.innerHTML = `<optgroup label="— tylko licz, bez konkretnej kolumny —"><option value="_count_">Liczba ${getRowEntityLabel()}</option></optgroup>` +
    `<optgroup label="Kolumny liczbowe (suma/średnia/min/max)">${cols.filter(c => colTypes[c] === 'numeric').map(c => `<option value="${c}">${c}</option>`).join('')}</optgroup>` +
    `<optgroup label="Kolumny tekstowe/kategorie (rozkład kolorami)">${cols.filter(c => colTypes[c] !== 'numeric').map(c => `<option value="${c}">${c}</option>`).join('')}</optgroup>`;
  bs.innerHTML = `<option value="">Brak</option>` +
    cols.filter(c => colTypes[c] === 'categorical' || colTypes[c] === 'text').map(c => `<option value="${c}">${c}</option>`).join('');
  by.dispatchEvent(new Event('change'));
  bs.dispatchEvent(new Event('change'));
  updateBuilderLockState();
  renderBuilderFilters();
  renderCalculatedFieldsList();
  renderRelationshipSelect();
  renderCalcFieldColumnPicker();
  const countOpt = document.querySelector('#b-agg option[value="count"]');
  if (countOpt) countOpt.textContent = 'Liczba ' + getRowEntityLabel();
}

function updateBuilderLockState() {
  const stackVal = document.getElementById('b-stack').value;
  const chartType = document.getElementById('b-type').value;
  const isPivot = chartType === 'pivot';
  // Stacking only actually changes anything for bar charts — for pie/line/doughnut
  // "Stack By" is simply ignored, so locking Y/Aggregation in that case would just
  // confuse the user with greyed-out fields that have no visible cause.
  const stacking = !isPivot && !!stackVal && (chartType === 'bar' || chartType === 'barH');
  document.getElementById('b-y').disabled = stacking;
  document.getElementById('b-agg').disabled = stacking;
  const hint = document.getElementById('b-y-lock-hint');
  if (hint) hint.style.display = stacking ? 'block' : 'none';

  // In pivot mode, "Stack By" becomes the column dimension of the matrix —
  // rename it and surface the advanced panel automatically, since it's
  // essential here rather than a rarely-used extra.
  const stackLabel = document.getElementById('b-stack-label');
  if (stackLabel) stackLabel.textContent = isPivot ? 'Kolumny (wymiar 2, opcjonalnie)' : 'Stack By (opcjonalnie)';
  if (isPivot) {
    const adv = document.getElementById('builder-advanced');
    if (adv && adv.style.display === 'none') toggleBuilderAdvanced();
  }
}
document.getElementById('b-stack')?.addEventListener('change', updateBuilderLockState);
document.getElementById('b-type')?.addEventListener('change', updateBuilderLockState);

let customChartCount = 0;

document.getElementById('b-y')?.addEventListener('change', function() {
  const agg = document.getElementById('b-agg');
  const isNumeric = this.value === '_count_' || colTypes[this.value] === 'numeric';
  Array.from(agg.options).forEach(opt => {
    opt.disabled = !isNumeric && opt.value !== 'count';
  });
  if (!isNumeric) agg.value = 'count';
  const hint = document.getElementById('b-agg-lock-hint');
  if (hint) {
    hint.style.display = isNumeric ? 'none' : 'block';
    hint.textContent = `⚠️ Tylko „Liczba ${getRowEntityLabel()}" działa, bo „Wartość Y" to kolumna tekstowa — Suma/Średnia/Max/Min wymagają kolumny liczbowej`;
  }
});

// If X axis changes to the same column currently selected as Y, that
// combination is meaningless ("group X by itself") and is almost always
// an artifact of X being changed after Y was left on a default/previous
// selection — reset Y back to Row Count so Aggregation doesn't end up
// stuck on "Count only" without an obvious reason.
document.getElementById('b-x')?.addEventListener('change', function() {
  const by = document.getElementById('b-y');
  if (by.value === this.value) {
    by.value = '_count_';
    by.dispatchEvent(new Event('change'));
  }
});

// ─── BUILDER FILTERS ───────────────────────────────────────────
// Optional row of column/operator/value filters applied only to the
// custom chart being built (Charts tab), independent of the Table
// tab's own search/filter. Lets you e.g. restrict a chart to a date
// range, one brand, or operators above a quantity threshold.
let builderFilters = [];
let builderFilterSeq = 0;

function defaultOpFor(col) {
  const t = colTypes[col];
  if (t === 'numeric') return 'gte';
  if (t === 'date') return 'after';
  return 'eq';
}

function addBuilderFilter() {
  const col = cols[0] || '';
  builderFilters.push({ id: ++builderFilterSeq, col, op: defaultOpFor(col), val: '', val2: '' });
  renderBuilderFilters();
}

function removeBuilderFilter(id) {
  builderFilters = builderFilters.filter(f => f.id !== id);
  renderBuilderFilters();
}

function updateBuilderFilter(id, field, value) {
  const f = builderFilters.find(f => f.id === id);
  if (!f) return;
  f[field] = value;
  if (field === 'col') { f.op = defaultOpFor(value); f.val = ''; f.val2 = ''; }
  renderBuilderFilters();
}

function escAttr(s) {
  return String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function renderBuilderFilters() {
  const host = document.getElementById('b-filters-list');
  if (!host) return;
  if (!builderFilters.length) { host.innerHTML = '<div style="font-size:11px;color:var(--ink-dim);">Brak filtrów — wykres używa wszystkich wierszy z pliku.</div>'; return; }

  host.innerHTML = builderFilters.map(f => {
    const type = colTypes[f.col];
    const colOptions = cols.map(c => `<option value="${escAttr(c)}" ${c===f.col?'selected':''}>${escAttr(c)}</option>`).join('');
    let opOptions = '', valueInput = '';

    if (type === 'numeric') {
      opOptions = `
        <option value="gte" ${f.op==='gte'?'selected':''}>≥</option>
        <option value="lte" ${f.op==='lte'?'selected':''}>≤</option>
        <option value="eq" ${f.op==='eq'?'selected':''}>=</option>
        <option value="between" ${f.op==='between'?'selected':''}>zakres od–do</option>`;
      valueInput = f.op === 'between'
        ? `<input type="number" value="${f.val}" oninput="updateBuilderFilterValueOnly(${f.id},'val',this.value)" style="width:90px;" placeholder="od">
           <input type="number" value="${f.val2}" oninput="updateBuilderFilterValueOnly(${f.id},'val2',this.value)" style="width:90px;" placeholder="do">`
        : `<input type="number" value="${f.val}" oninput="updateBuilderFilterValueOnly(${f.id},'val',this.value)" style="width:110px;">`;
    } else if (type === 'date') {
      opOptions = `
        <option value="after" ${f.op==='after'?'selected':''}>po (włącznie)</option>
        <option value="before" ${f.op==='before'?'selected':''}>przed (włącznie)</option>
        <option value="between" ${f.op==='between'?'selected':''}>zakres od–do</option>`;
      valueInput = f.op === 'between'
        ? `<input type="date" value="${f.val}" oninput="updateBuilderFilterValueOnly(${f.id},'val',this.value)">
           <input type="date" value="${f.val2}" oninput="updateBuilderFilterValueOnly(${f.id},'val2',this.value)">`
        : `<input type="date" value="${f.val}" oninput="updateBuilderFilterValueOnly(${f.id},'val',this.value)">`;
    } else {
      opOptions = `
        <option value="eq" ${f.op==='eq'?'selected':''}>jest równe</option>
        <option value="neq" ${f.op==='neq'?'selected':''}>jest różne od</option>`;
      const uniqVals = [...new Set(allData.map(r => String(r[f.col] ?? '')))].filter(v=>v!=='').sort().slice(0, 500);
      valueInput = `<select onchange="updateBuilderFilterValueOnly(${f.id},'val',this.value)" style="min-width:140px;">
        <option value="">— wybierz wartość —</option>
        ${uniqVals.map(v => `<option value="${escAttr(v)}" ${v===f.val?'selected':''}>${escAttr(v)}</option>`).join('')}
      </select>`;
    }

    return `<div style="display:flex;gap:8px;align-items:center;margin-bottom:8px;flex-wrap:wrap;">
      <select onchange="updateBuilderFilter(${f.id},'col',this.value)" style="min-width:140px;">${colOptions}</select>
      <select onchange="updateBuilderFilter(${f.id},'op',this.value)" style="min-width:130px;">${opOptions}</select>
      ${valueInput}
      <button onclick="removeBuilderFilter(${f.id})" style="background:none;border:none;cursor:pointer;color:var(--ink-dim);font-size:15px;" title="Usuń filtr">✕</button>
    </div>`;
  }).join('');
}

function updateBuilderFilterValueOnly(id, field, value) {
  const f = builderFilters.find(f => f.id === id);
  if (f) f[field] = value;
}

function getBuilderFilteredData() {
  const base = getScopedData();
  const active = builderFilters.filter(f => f.col && (f.val !== '' || f.op === 'between' && (f.val !== '' || f.val2 !== '')));
  if (!active.length) return base;
  return base.filter(row => active.every(f => {
    const raw = row[f.col];
    const type = colTypes[f.col];
    if (type === 'numeric') {
      const v = toNum(raw);
      if (v === null) return false;
      if (f.op === 'gte') return f.val === '' || v >= parseFloat(f.val);
      if (f.op === 'lte') return f.val === '' || v <= parseFloat(f.val);
      if (f.op === 'eq') return f.val === '' || v === parseFloat(f.val);
      if (f.op === 'between') return (f.val === '' || v >= parseFloat(f.val)) && (f.val2 === '' || v <= parseFloat(f.val2));
      return true;
    }
    if (type === 'date') {
      const d = String(raw).slice(0, 10);
      if (!d) return false;
      if (f.op === 'after') return f.val === '' || d >= f.val;
      if (f.op === 'before') return f.val === '' || d <= f.val;
      if (f.op === 'between') return (f.val === '' || d >= f.val) && (f.val2 === '' || d <= f.val2);
      return true;
    }
    const s = String(raw ?? '');
    if (f.op === 'neq') return f.val === '' || s !== f.val;
    return f.val === '' || s === f.val;
  }));
}

function showBuilderValidation(msg, actionLabel, actionFn) {
  const el = document.getElementById('b-validation-hint');
  if (!el) { alert(msg); return; }
  el.innerHTML = '⚠️ ' + msg;
  if (actionLabel && actionFn) {
    window.__builderValidationAction = actionFn;
    el.innerHTML += ` <button onclick="window.__builderValidationAction()" style="margin-left:8px;background:var(--amber-dim);color:var(--amber);border:1px solid var(--amber);border-radius:5px;padding:3px 10px;font-size:11px;cursor:pointer;">${actionLabel}</button>`;
  }
  el.style.display = 'block';
}
function clearBuilderValidation() {
  const el = document.getElementById('b-validation-hint');
  if (el) el.style.display = 'none';
}

// ─── CALCULATED FIELDS ──────────────────────────────────────
// Power BI / Excel-style "measure": define a formula referencing
// existing columns as {ColumnName}, and it becomes a real virtual
// column usable anywhere (builder, filters, pivot, table) just like
// any column that was actually in the file.
let calculatedFields = [];

// ─── EXCEL-LIKE FORMULA ENGINE ──────────────────────────────────
// Supports {ColumnName} references plus a whitelisted set of Excel-style
// functions (IF, ROUND, ABS, MIN, MAX, AND, OR, NOT, CONCAT, LEN, ROUNDUP,
// ROUNDDOWN, UPPER, LOWER, LEFT, RIGHT) and comparisons using Excel syntax
// ("=" for equals, "<>" for not-equal) on top of standard operators.
// Sandboxed: only these named functions are reachable, nothing from the
// page's global scope (window, fetch, etc.) is exposed to the formula.
const FORMULA_FUNCS = {
  IF: (c, t, f) => (c ? t : f),
  ROUND: (x, n = 0) => Math.round(x * 10 ** n) / 10 ** n,
  ROUNDUP: (x, n = 0) => Math.ceil(x * 10 ** n) / 10 ** n,
  ROUNDDOWN: (x, n = 0) => Math.floor(x * 10 ** n) / 10 ** n,
  ABS: Math.abs,
  MIN: (...a) => Math.min(...a),
  MAX: (...a) => Math.max(...a),
  AND: (...a) => a.every(Boolean),
  OR: (...a) => a.some(Boolean),
  NOT: a => !a,
  CONCAT: (...a) => a.join(''),
  LEN: s => String(s).length,
  UPPER: s => String(s).toUpperCase(),
  LOWER: s => String(s).toLowerCase(),
  LEFT: (s, n) => String(s).slice(0, n),
  RIGHT: (s, n) => String(s).slice(-n),
};
const FORMULA_FUNC_NAMES = Object.keys(FORMULA_FUNCS);

function evalFormulaForRow(formula, row) {
  // 1) Substitute {ColumnName} with the row's actual value — numbers stay
  // raw, everything else becomes a quoted, escaped JS string literal so
  // text comparisons like {Status}="OK" work correctly.
  let expr = formula.replace(/\{([^}]+)\}/g, (_, colName) => {
    const raw = row[colName];
    const n = toNum(raw);
    if (n !== null) return `(${n})`;
    const s = String(raw ?? '').replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    return `"${s}"`;
  });
  // 2) Excel-style comparison syntax -> JS: "<>"->"!=", bare "="->"==".
  expr = expr.replace(/<>/g, '!=');
  expr = expr.replace(/([^=!<>])=(?!=)/g, '$1==');

  // 3) Pull out string literals so the whitelist check below doesn't have
  // to special-case arbitrary quoted content.
  const literals = [];
  let masked = expr.replace(/"(?:[^"\\]|\\.)*"/g, m => { literals.push(m); return `§${literals.length - 1}§`; });

  // 4) Whitelist: only digits, operators, parens, commas, whitespace,
  // the §N§ literal placeholders, and known function names survive.
  const bareWords = masked.match(/[A-Za-z_]\w*/g) || [];
  const allowedWords = new Set([...FORMULA_FUNC_NAMES, 'true', 'false', 'NaN']);
  if (bareWords.some(w => !allowedWords.has(w))) return null;
  if (!/^[\d\s.+\-*/(),§\d!<>=&|A-Za-z_]+$/.test(masked)) return null;

  const restored = masked.replace(/§(\d+)§/g, (_, i) => literals[i]);
  try {
    const fn = new Function(...FORMULA_FUNC_NAMES, '"use strict"; return (' + restored + ');');
    const result = fn(...FORMULA_FUNC_NAMES.map(k => FORMULA_FUNCS[k]));
    if (typeof result === 'boolean') return result;
    if (typeof result === 'number') return isFinite(result) ? result : null;
    return result == null ? null : String(result);
  } catch (e) { return null; }
}

function applyCalculatedField(name, formula) {
  allData.forEach(row => { row[name] = evalFormulaForRow(formula, row); });
  const sample = allData.slice(0, 50).map(r => r[name]).filter(v => v !== null && v !== undefined);
  const allNum = sample.length && sample.every(v => typeof v === 'number');
  if (!cols.includes(name)) cols.push(name);
  colTypes[name] = allNum ? 'numeric' : detectType(name, allData);
}

function addCalculatedField() {
  const nameInput = document.getElementById('calc-field-name');
  const formulaInput = document.getElementById('calc-field-formula');
  const name = nameInput.value.trim();
  const formula = formulaInput.value.trim();
  if (!name || !formula) { showBuilderValidation('Podaj nazwę pola i formułę, np. nazwa „Efektywność %”, formuła „{Qty_Faktyczna} / {Qty_Oczekiwana} * 100”.'); return; }
  if (cols.includes(name)) { showBuilderValidation(`Kolumna „${name}” już istnieje w pliku — wybierz inną nazwę dla pola obliczonego.`); return; }
  const refs = [...formula.matchAll(/\{([^}]+)\}/g)].map(m => m[1]);
  if (!refs.length) { showBuilderValidation('Formuła musi odwoływać się do przynajmniej jednej kolumny w nawiasach klamrowych, np. {Qty_Faktyczna}.'); return; }
  const missing = refs.filter(r => !cols.includes(r));
  if (missing.length) { showBuilderValidation(`Formuła odwołuje się do nieistniejącej kolumny: „${missing[0]}”. Sprawdź dokładną nazwę (wielkość liter też ma znaczenie).`); return; }

  calculatedFields.push({ name, formula });
  applyCalculatedField(name, formula);
  nameInput.value = ''; formulaInput.value = '';
  clearBuilderValidation();
  renderCalculatedFieldsList();
  renderBuilderSelects();
  renderFilterSelects();
}

function removeCalculatedField(name) {
  calculatedFields = calculatedFields.filter(f => f.name !== name);
  allData.forEach(row => { delete row[name]; });
  cols = cols.filter(c => c !== name);
  delete colTypes[name];
  renderCalculatedFieldsList();
  renderBuilderSelects();
  renderFilterSelects();
}

function renderCalculatedFieldsList() {
  const host = document.getElementById('calc-fields-list');
  if (!host) return;
  if (!calculatedFields.length) { host.innerHTML = '<div style="font-size:11px;color:var(--ink-dim);">Brak pól obliczonych.</div>'; return; }
  host.innerHTML = calculatedFields.map(f => `
    <div style="display:flex;gap:8px;align-items:center;margin-bottom:6px;">
      <span style="background:rgba(232,162,61,.1);color:var(--amber);padding:3px 10px;border-radius:6px;font-size:12px;border:1px solid rgba(232,162,61,.3);">${escAttr(f.name)} = ${escAttr(f.formula)}</span>
      <button onclick="removeCalculatedField('${f.name.replace(/'/g, "\\'")}')" style="background:none;border:none;cursor:pointer;color:var(--ink-dim);font-size:14px;" title="Usuń pole">✕</button>
    </div>`).join('');
}

function insertColumnToken(targetId, colName) {
  const el = document.getElementById(targetId);
  if (!el || !colName) return;
  const token = `{${colName}}`;
  const start = el.selectionStart ?? el.value.length;
  el.value = el.value.slice(0, start) + token + el.value.slice(start);
  el.focus();
}

function renderCalcFieldColumnPicker() {
  const sel = document.getElementById('calc-field-col-picker');
  if (!sel) return;
  sel.innerHTML = `<option value="">+ Wstaw kolumnę…</option>` +
    cols.filter(c => colTypes[c] === 'numeric').map(c => `<option value="${escAttr(c)}">${escAttr(c)}</option>`).join('');
}

// ─── COLUMN RELATIONSHIP EXPLORER ───────────────────────────────
// Auto-generates every meaningful column-pair combination (number×number,
// number×text, text×text, date×number...) so you can pick from a list
// instead of manually configuring each one. Picking a pair renders a
// single card with a table, a chart, AND a plain-language summary —
// the same three views, regardless of which data types are involved.
const REL_PRIORITY = { numeric: 3, date: 2, categorical: 1, time: 1, text: 0, unknown: -1 };

function generateAllRelationships() {
  const combos = [];
  for (let i = 0; i < cols.length; i++) {
    for (let j = i + 1; j < cols.length; j++) {
      const a = cols[i], b = cols[j];
      const ta = colTypes[a], tb = colTypes[b];
      if (ta === 'unknown' || tb === 'unknown') continue;
      const score = (REL_PRIORITY[ta] || 0) + (REL_PRIORITY[tb] || 0);
      combos.push({ a, b, ta, tb, score, label: `${a} × ${b}  (${TYPE_LABELS_PL[ta]} × ${TYPE_LABELS_PL[tb]})` });
    }
  }
  combos.sort((x, y) => y.score - x.score);
  return combos;
}

function renderRelationshipSelect() {
  const sel = document.getElementById('rel-select');
  if (!sel) return;
  const combos = generateAllRelationships();
  if (!combos.length) { sel.innerHTML = '<option value="">Za mało kolumn do porównania</option>'; return; }
  sel.dataset.allCombos = JSON.stringify(combos);
  sel.innerHTML = combos.map((c, i) => `<option value="${i}">${escAttr(c.label)}</option>`).join('');
  sel.dataset.combos = JSON.stringify(combos);
  const search = document.getElementById('rel-search');
  if (search) search.value = '';
}

function filterRelationshipOptions() {
  const query = document.getElementById('rel-search').value.trim().toLowerCase();
  const sel = document.getElementById('rel-select');
  const all = JSON.parse(sel.dataset.allCombos || '[]');
  const filtered = query ? all.filter(c => c.label.toLowerCase().includes(query)) : all;
  if (!filtered.length) { sel.innerHTML = '<option value="">— nic nie znaleziono, spróbuj innego słowa —</option>'; sel.dataset.combos = '[]'; return; }
  // dataset.combos must use the SAME indices as the option values we render below
  sel.innerHTML = filtered.map((c, i) => `<option value="${i}">${escAttr(c.label)}</option>`).join('');
  sel.dataset.combos = JSON.stringify(filtered);
}

function toggleManualBuilder() {
  const wrap = document.getElementById('manual-builder-wrap');
  const btn = document.getElementById('manual-builder-toggle-btn');
  const show = wrap.style.display === 'none';
  wrap.style.display = show ? 'block' : 'none';
  btn.textContent = show
    ? '▴ Schowaj zaawansowany konstruktor'
    : '▾ Zaawansowane — zbuduj własny wykres ręcznie (dla Advanced User / Team Leadera)';
}

function buildSelectedRelationship() {
  const sel = document.getElementById('rel-select');
  const combos = JSON.parse(sel.dataset.combos || '[]');
  const combo = combos[parseInt(sel.value, 10)];
  if (!combo) return;
  exploreRelationship(combo.a, combo.b, combo.ta, combo.tb);
}

// Detects whether a column's values look like Excel clock-duration
// strings (HH:MM[:SS]) — i.e. the kind toNum() converts to total minutes.
// Used purely for DISPLAY: a sum of "minutes" should say so, not show as
// an unexplained "1.8K".
function isDurationColumn(col) {
  const vals = allData.slice(0, 200).map(r => r[col]).filter(v => v !== '' && v != null);
  if (!vals.length) return false;
  const hits = vals.filter(v => /^\d{1,2}:\d{2}(:\d{2})?$/.test(String(v).trim())).length;
  return hits / vals.length > 0.6;
}
function fmtDuration(totalMinutes) {
  if (totalMinutes == null || isNaN(totalMinutes)) return '—';
  const sign = totalMinutes < 0 ? '-' : '';
  totalMinutes = Math.abs(totalMinutes);
  const h = Math.floor(totalMinutes / 60);
  const m = Math.round(totalMinutes % 60);
  if (h === 0) return `${sign}${m} min`;
  return `${sign}${h} godz ${m} min`;
}

function exploreRelationship(a, b, ta, tb) {
  const data = getBuilderFilteredData();
  if (!data.length) { showBuilderValidation('Brak danych do analizy — sprawdź filtry.'); return; }

  let tableRows = [], tableHeaders = null, chartLabels = [], chartData = [], datasets = null, summary = '', chartType = 'bar', valueLabel = '';

  const isNumNum = ta === 'numeric' && tb === 'numeric';
  const isDateInvolved = ta === 'date' || tb === 'date';
  const numCol = ta === 'numeric' ? a : (tb === 'numeric' ? b : null);
  const otherCol = numCol === a ? b : a;
  const otherType = numCol === a ? tb : ta;
  const numIsDuration = numCol ? isDurationColumn(numCol) : false;
  const fmtNum = numIsDuration ? fmtDuration : fmt;

  if (isNumNum) {
    // number × number -> correlation + binned average chart
    const aIsDur = isDurationColumn(a), bIsDur = isDurationColumn(b);
    const xs = [], ys = [];
    data.forEach(r => { const x = toNum(r[a]), y = toNum(r[b]); if (x !== null && y !== null) { xs.push(x); ys.push(y); } });
    const r = pearson(xs, ys);
    const BINS = 8;
    const minX = Math.min(...xs), maxX = Math.max(...xs);
    const span = (maxX - minX) || 1;
    const bins = Array.from({ length: BINS }, () => []);
    xs.forEach((x, i) => { let idx = Math.floor((x - minX) / span * BINS); if (idx >= BINS) idx = BINS - 1; if (idx < 0) idx = 0; bins[idx].push(ys[i]); });
    const fmtA = aIsDur ? fmtDuration : fmt, fmtB = bIsDur ? fmtDuration : fmt;
    chartLabels = bins.map((_, i) => `${fmtA(minX + span * i / BINS)}–${fmtA(minX + span * (i + 1) / BINS)}`);
    chartData = bins.map(arr => arr.length ? mean(arr) : 0);
    valueLabel = `Średnia „${b}” w przedziale „${a}”`;
    tableHeaders = [`Przedział „${a}”`, 'Liczba', `Średnia „${b}”`];
    tableRows = chartLabels.map((l, i) => [l, bins[i].length, fmtB(chartData[i])]);
    summary = r === null ? `Za mało danych, aby ocenić zależność między „${a}” i „${b}”.`
      : `Współczynnik korelacji r = ${r.toFixed(2)} — ${Math.abs(r) >= 0.5 ? (r > 0 ? 'wyraźna zależność wprost: gdy „' + a + '” rośnie, „' + b + '” zwykle też.' : 'wyraźna zależność odwrotna: gdy „' + a + '” rośnie, „' + b + '” zwykle spada.') : 'brak silnej liniowej zależności między tymi kolumnami.'}`;
  } else if (numCol && (otherType === 'categorical' || otherType === 'text' || otherType === 'time')) {
    // number × category -> sum/avg/count grouped by category
    const groups = {};
    data.forEach(r => {
      const k = String(r[otherCol] ?? '(empty)');
      const v = toNum(r[numCol]);
      if (v === null) return;
      (groups[k] = groups[k] || []).push(v);
    });
    const entries = Object.entries(groups).sort((x, y) => y[1].reduce((s, v) => s + v, 0) - x[1].reduce((s, v) => s + v, 0)).slice(0, 20);
    chartLabels = entries.map(e => e[0]);
    chartData = entries.map(e => Math.round(e[1].reduce((s, v) => s + v, 0) * 100) / 100);
    valueLabel = `Suma „${numCol}”${numIsDuration ? ' (czas)' : ''} według „${otherCol}”`;
    tableHeaders = [otherCol, `Liczba ${getRowEntityLabel()}`, `Suma „${numCol}”`, `Średnia „${numCol}” (na 1 rekord)`];
    tableRows = entries.map(([k, arr]) => [k, arr.length, fmtNum(arr.reduce((s, v) => s + v, 0)), fmtNum(mean(arr))]);
    const top = entries[0], low = entries[entries.length - 1];
    summary = `Najwyższa suma „${numCol}” — „${top?.[0]}” (${fmtNum(top?.[1].reduce((s,v)=>s+v,0))}). Najniższa — „${low?.[0]}” (${fmtNum(low?.[1].reduce((s,v)=>s+v,0))}). ${entries.length} kategorii w „${otherCol}”.`;
  } else if (isDateInvolved && numCol) {
    // date × number -> trend over time
    const dateCol = ta === 'date' ? a : b;
    const byDate = {};
    data.forEach(r => {
      const d = String(r[dateCol]).slice(0, 10);
      const v = toNum(r[numCol]);
      if (v === null || !d) return;
      byDate[d] = (byDate[d] || 0) + v;
    });
    const days = Object.keys(byDate).sort();
    chartLabels = days; chartData = days.map(d => Math.round(byDate[d] * 100) / 100);
    chartType = 'line';
    valueLabel = `Suma „${numCol}”${numIsDuration ? ' (czas)' : ''} dziennie`;
    tableHeaders = [dateCol, `Suma „${numCol}”`];
    tableRows = days.map(d => [d, fmtNum(byDate[d])]);
    const mid = Math.floor(days.length / 2) || 1;
    const f = mean(days.slice(0, mid).map(d => byDate[d])), s = mean(days.slice(mid).map(d => byDate[d]));
    summary = f ? `Trend ${s > f * 1.1 ? 'wzrostowy' : s < f * 0.9 ? 'spadkowy' : 'stabilny'} — druga połowa okresu: ${((s - f) / f * 100).toFixed(0)}% względem pierwszej.` : 'Za mało dni do oceny trendu.';
  } else {
    // category × category (or date × category, text × text) -> REAL crosstab,
    // not just a count-by-"a" that silently drops "b". Rendered as a proper
    // stacked bar (one series per "b" value) plus a grid table of the actual
    // intersection counts, so the breakdown by "b" is genuinely visible.
    const crosstab = {}; const rowTotals = {}; const colTotals = {};
    data.forEach(r => {
      const ra = String(r[a] ?? '(empty)'), rb = String(r[b] ?? '(empty)');
      crosstab[ra] = crosstab[ra] || {};
      crosstab[ra][rb] = (crosstab[ra][rb] || 0) + 1;
      rowTotals[ra] = (rowTotals[ra] || 0) + 1;
      colTotals[rb] = (colTotals[rb] || 0) + 1;
    });
    const rows = Object.keys(crosstab).sort((x, y) => rowTotals[y] - rowTotals[x]).slice(0, 12);
    const columns = Object.keys(colTotals).sort((x, y) => colTotals[y] - colTotals[x]).slice(0, 6);
    chartLabels = rows;
    datasets = columns.map((c, i) => ({
      label: c,
      data: rows.map(r => (crosstab[r] && crosstab[r][c]) || 0),
      backgroundColor: COLORS[i % COLORS.length],
    }));
    valueLabel = `Liczba ${getRowEntityLabel()} — „${a}” podzielone kolorami według „${b}”`;
    tableHeaders = [a, ...columns, 'Razem'];
    tableRows = rows.map(r => [r, ...columns.map(c => (crosstab[r] && crosstab[r][c]) || 0), rowTotals[r]]);
    const top = rows[0];
    summary = `Najczęstsza wartość „${a}” — „${top}” (${rowTotals[top]} ${getRowEntityLabel()}). Łącznie ${Object.keys(crosstab).length} wartości „${a}” × ${Object.keys(colTotals).length} wartości „${b}”${rows.length < Object.keys(crosstab).length ? ` (pokazano top ${rows.length})` : ''}.`;
  }

  renderRelationshipCard(a, b, chartLabels, chartData, chartType, valueLabel, tableRows, summary, tableHeaders, datasets);
}

function renderRelationshipCard(a, b, labels, data, chartType, valueLabel, tableRows, summary, tableHeaders, datasets) {
  const id = `cc_${customChartCount++}`;
  const wrap = document.createElement('div');
  wrap.className = 'chart-card';
  wrap.id = id;
  const isH = labels.length > 6 && chartType !== 'line';
  const h = isH ? Math.max(180, labels.length * 26) : 240;
  const theadHtml = tableHeaders
    ? `<thead><tr>${tableHeaders.map((hh, ci) => `<th style="padding:5px 8px;text-align:${ci===0?'left':'right'};font-size:10.5px;color:var(--ink-dim);border-bottom:1px solid var(--line);white-space:nowrap;">${escAttr(hh)}</th>`).join('')}</tr></thead>`
    : '';
  wrap.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:10px;">
      <div>
        <div class="chart-card-title">${escAttr(a)} × ${escAttr(b)}</div>
        <div class="chart-card-sub">${escAttr(valueLabel)}</div>
      </div>
      <button onclick="removeChart(this,'${id}')" style="background:none;border:none;cursor:pointer;font-size:16px;color:var(--ink-dim);flex-shrink:0;">✕</button>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:10px;align-items:start;">
      <div style="position:relative;width:100%;height:${h}px;"><canvas id="${id}_chart"></canvas></div>
      <div style="overflow:auto;max-height:${h}px;border:1px solid var(--line);border-radius:8px;">
        <table style="width:100%;border-collapse:collapse;font-size:11.5px;">
          ${theadHtml}
          <tbody>
            ${tableRows.slice(0, 30).map(row => `<tr>${row.map((cell, ci) => `<td style="padding:5px 8px;border-bottom:1px solid var(--line);text-align:${ci===0?'left':'right'};white-space:nowrap;">${escAttr(cell)}</td>`).join('')}</tr>`).join('')}
          </tbody>
        </table>
      </div>
    </div>
    <div class="chart-insight">📌 ${summary}</div>
  `;
  document.getElementById('custom-charts-area').insertBefore(wrap, document.getElementById('custom-charts-area').firstChild);
  setTimeout(() => {
    if (datasets) {
      new Chart(document.getElementById(`${id}_chart`), {
        type: 'bar',
        data: { labels, datasets },
        options: {
          responsive: true, maintainAspectRatio: false, indexAxis: isH ? 'y' : 'x',
          scales: { x: { stacked: true, grid: { display: false }, ticks: { font: { size: 10 } } }, y: { stacked: true, grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 10 } } } },
          plugins: { legend: { display: true, position: 'bottom', labels: { font: { size: 10 }, boxWidth: 10 } } },
        }
      });
    } else {
      createChart(`${id}_chart`, chartType === 'line' ? 'line' : 'bar', labels, data, valueLabel, isH, true, a);
    }
  }, 50);
}


function aggregateValues(values, agg) {
  if (!values.length) return null;
  if (agg === 'count') return values.length;
  if (agg === 'avg') return values.reduce((a, b) => a + b, 0) / values.length;
  if (agg === 'max') return Math.max(...values);
  if (agg === 'min') return Math.min(...values);
  return values.reduce((a, b) => a + b, 0); // sum (default)
}

// ─── PIVOT TABLE ────────────────────────────────────────────
// The literal "pivot table" concept from Excel/Power BI: pick a row
// dimension, an optional column dimension, and a value — get a real
// matrix with row/column/grand totals, not just a single-axis chart.
let pivotInstances = {};

function buildPivotTable(rowsCol, colsCol, yCol, agg, titleInput, existingId, sortKey) {
  if (!rowsCol) { showBuilderValidation('Wybierz kolumnę w „Oś X / Kategoria” — to będą wiersze tabeli przestawnej.'); return; }
  const sourceData = getBuilderFilteredData();
  if (!sourceData.length) {
    showBuilderValidation('Ustawione filtry usunęły wszystkie wiersze — nie ma czego pokazać. Sprawdź wartości w filtrach.');
    return;
  }
  if (colsCol && colsCol === rowsCol) {
    showBuilderValidation('„Oś X” i „Kolumny” to ta sama kolumna — wybierz dwie różne, albo wyczyść „Kolumny”.');
    return;
  }
  clearBuilderValidation();

  const isCountAgg = yCol === '_count_' || agg === 'count';
  const effAgg = isCountAgg ? 'count' : agg;

  const cellRaw = {};   // "row||col" -> array of numbers (or count placeholders)
  const rowFreq = {};   // row -> number of source rows (for top-N sorting)
  const rowSet = new Set(), colSet = new Set();

  sourceData.forEach(row => {
    const r = String(row[rowsCol] ?? '(empty)');
    const c = colsCol ? String(row[colsCol] ?? '(empty)') : '__value__';
    rowSet.add(r); colSet.add(c);
    rowFreq[r] = (rowFreq[r] || 0) + 1;
    const key = r + '||' + c;
    if (isCountAgg) {
      (cellRaw[key] = cellRaw[key] || []).push(1);
    } else {
      const v = toNum(row[yCol]);
      if (v !== null) (cellRaw[key] = cellRaw[key] || []).push(v);
    }
  });

  const MAX_ROWS = 40, MAX_COLS = 12;
  let rows = [...rowSet].sort((a, b) => (rowFreq[b] || 0) - (rowFreq[a] || 0));
  const wasTruncatedRows = rows.length > MAX_ROWS;
  if (wasTruncatedRows) rows = rows.slice(0, MAX_ROWS);

  let columns = colsCol ? [...colSet] : ['__value__'];
  let wasTruncatedCols = false;
  if (colsCol && columns.length > MAX_COLS) {
    const colTotal = {};
    columns.forEach(c => { colTotal[c] = Object.keys(cellRaw).filter(k => k.endsWith('||' + c)).reduce((s, k) => s + cellRaw[k].length, 0); });
    columns = columns.sort((a, b) => colTotal[b] - colTotal[a]).slice(0, MAX_COLS);
    wasTruncatedCols = true;
  }

  const matrix = {}, rowTotals = {}, colTotals = {};
  rows.forEach(r => {
    matrix[r] = {};
    let rowVals = [];
    columns.forEach(c => {
      const vals = cellRaw[r + '||' + c] || [];
      matrix[r][c] = vals.length ? aggregateValues(vals, effAgg) : null;
      rowVals = rowVals.concat(vals);
    });
    rowTotals[r] = rowVals.length ? aggregateValues(rowVals, effAgg) : null;
  });
  let grandVals = [];
  columns.forEach(c => {
    let colVals = [];
    rows.forEach(r => { colVals = colVals.concat(cellRaw[r + '||' + c] || []); });
    colTotals[c] = colVals.length ? aggregateValues(colVals, effAgg) : null;
    grandVals = grandVals.concat(colVals);
  });
  const grandTotal = grandVals.length ? aggregateValues(grandVals, effAgg) : null;

  if (sortKey) {
    rows = [...rows].sort((a, b) => {
      const va = sortKey === '__ROWTOTAL__' ? rowTotals[a] : matrix[a][sortKey];
      const vb = sortKey === '__ROWTOTAL__' ? rowTotals[b] : matrix[b][sortKey];
      return (vb ?? -Infinity) - (va ?? -Infinity);
    });
  }

  const aggLabelsPl = { sum: 'Suma', avg: 'Średnia', max: 'Maksimum', min: 'Minimum', count: `Liczba ${getRowEntityLabel()}` };
  const valueLabel = aggLabelsPl[effAgg] + (isCountAgg ? '' : ` „${yCol}”`);
  const chartTitle = titleInput || `Tabela przestawna: ${rowsCol}${colsCol ? ' × ' + colsCol : ''}`;

  const allNums = [];
  rows.forEach(r => columns.forEach(c => { const v = matrix[r][c]; if (v != null) allNums.push(v); }));
  const maxV = allNums.length ? Math.max(...allNums) : 0, minV = allNums.length ? Math.min(...allNums) : 0;
  const cellBg = v => {
    if (v == null || maxV === minV) return 'transparent';
    const t = (v - minV) / (maxV - minV);
    return `rgba(59,130,246,${(0.06 + t * 0.34).toFixed(2)})`;
  };

  const id = existingId || `cc_${customChartCount++}`;
  pivotInstances[id] = { rowsCol, colsCol, yCol, agg, titleInput, sortKey };
  let wrap = existingId ? document.getElementById(existingId) : null;
  const isNewCard = !wrap;
  if (isNewCard) {
    wrap = document.createElement('div');
    wrap.className = 'chart-card';
    wrap.id = id;
  }
  wrap.innerHTML = `
    <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:10px;">
      <div>
        <div class="chart-card-title">${chartTitle}</div>
        <div class="chart-card-sub">Wiersze: „${rowsCol}”${colsCol ? ` · Kolumny: „${colsCol}”` : ''} · Wartości: ${valueLabel}${wasTruncatedRows || wasTruncatedCols ? ' · pokazano tylko najliczniejsze kategorie' : ''}</div>
      </div>
      <button onclick="removeChart(this,'${id}')" style="background:none;border:none;cursor:pointer;font-size:16px;color:var(--ink-dim);flex-shrink:0;">✕</button>
    </div>
    <div style="overflow:auto;max-height:480px;margin-top:10px;border:1px solid var(--line);border-radius:8px;">
      <table style="width:100%;border-collapse:collapse;font-size:12px;">
        <thead>
          <tr>
            <th style="position:sticky;top:0;left:0;z-index:3;background:var(--panel-2);padding:8px 10px;text-align:left;border-bottom:1px solid var(--line);">${escAttr(rowsCol)}</th>
            ${columns.map(c => {
              const safeC = String(c).replace(/'/g, "\\'");
              const active = sortKey === c;
              return `<th onclick="resortPivotTable('${id}','${safeC}')" style="position:sticky;top:0;z-index:2;background:${active?'rgba(232,162,61,.22)':'var(--panel-2)'};padding:8px 10px;text-align:right;border-bottom:1px solid var(--line);white-space:nowrap;cursor:pointer;" title="Kliknij, aby sortować wg tej kolumny">${escAttr(colsCol ? c : valueLabel)}${active?' ↓':''}</th>`;
            }).join('')}
            ${columns.length > 1 ? `<th onclick="resortPivotTable('${id}','__ROWTOTAL__')" style="position:sticky;top:0;z-index:2;background:${sortKey==='__ROWTOTAL__'?'rgba(232,162,61,.3)':'rgba(232,162,61,.12)'};padding:8px 10px;text-align:right;border-bottom:1px solid var(--line);font-weight:700;cursor:pointer;" title="Kliknij, aby sortować wg sumy wiersza">Razem${sortKey==='__ROWTOTAL__'?' ↓':''}</th>` : ''}
          </tr>
        </thead>
        <tbody>
          ${rows.map(r => `
            <tr>
              <td style="position:sticky;left:0;background:var(--panel);padding:7px 10px;border-bottom:1px solid var(--line);font-weight:600;">${escAttr(r)}</td>
              ${columns.map(c => {
                const v = matrix[r][c];
                const safeR = String(r).replace(/'/g, "\\'"); const safeC = String(c).replace(/'/g, "\\'");
                return `<td style="padding:7px 10px;border-bottom:1px solid var(--line);text-align:right;background:${cellBg(v)};${v != null ? 'cursor:pointer;' : ''}" ${v != null ? `onclick="drillToTable('${rowsCol.replace(/'/g, "\\'")}','${safeR}','${colsCol ? safeC : ''}')"` : ''}>${v == null ? '—' : fmt(v)}</td>`;
              }).join('')}
              ${columns.length > 1 ? `<td style="padding:7px 10px;border-bottom:1px solid var(--line);text-align:right;font-weight:700;background:rgba(232,162,61,.12);">${rowTotals[r] == null ? '—' : fmt(rowTotals[r])}</td>` : ''}
            </tr>`).join('')}
        </tbody>
        <tfoot>
          <tr>
            <td style="position:sticky;left:0;padding:7px 10px;font-weight:700;background:rgba(232,162,61,.12);">Razem</td>
            ${columns.map(c => `<td style="padding:7px 10px;text-align:right;font-weight:700;background:rgba(232,162,61,.12);">${colTotals[c] == null ? '—' : fmt(colTotals[c])}</td>`).join('')}
            ${columns.length > 1 ? `<td style="padding:7px 10px;text-align:right;font-weight:700;background:rgba(232,162,61,.25);">${grandTotal == null ? '—' : fmt(grandTotal)}</td>` : ''}
          </tr>
        </tfoot>
      </table>
    </div>
    <div style="font-size:10px;color:var(--ink-dim);margin-top:6px;">💡 Kliknij komórkę, aby zobaczyć dokładne wiersze w zakładce Tabela</div>`;
  if (isNewCard) document.getElementById('custom-charts-area').insertBefore(wrap, document.getElementById('custom-charts-area').firstChild);
}

function resortPivotTable(id, sortKey) {
  const cfg = pivotInstances[id];
  if (!cfg) return;
  const newSortKey = cfg.sortKey === sortKey ? null : sortKey; // click again to clear sort
  buildPivotTable(cfg.rowsCol, cfg.colsCol, cfg.yCol, cfg.agg, cfg.titleInput, id, newSortKey);
}

function buildCustomChart() {
  const type = document.getElementById('b-type').value;
  const xCol = document.getElementById('b-x').value;
  const yCol = document.getElementById('b-y').value;
  const agg = document.getElementById('b-agg').value;
  const stackCol = document.getElementById('b-stack').value;
  const limit = parseInt(document.getElementById('b-limit').value) || 0;
  const titleInput = document.getElementById('b-title').value.trim();

  if (type === 'pivot') {
    buildPivotTable(xCol, stackCol, yCol, agg, titleInput);
    return;
  }

  // Stacked crosstab path: only meaningful for bar/horizontal-bar charts
  if (stackCol && (type === 'bar' || type === 'barH')) {
    if (stackCol === xCol) {
      showBuilderValidation('„Oś X” i „Stack By” to ta sama kolumna — wybierz dwie różne kolumny, bo podział kolorami tej samej wartości nie ma sensu.');
      return;
    }
    buildStackedChart(xCol, stackCol, limit, type, titleInput);
    return;
  }

  // If Y is a non-numeric column (and not the explicit "Row Count" option),
  // a plain count-by-X would silently throw away the column you actually
  // picked. On a bar chart, treat it as an implicit "Stack By" instead, so
  // picking e.g. Y = "Status" on X = "Error type" gives you the breakdown
  // you'd expect, not just a row count that ignores Status entirely.
  const yIsNonNumericGroup = yCol !== '_count_' && colTypes[yCol] !== 'numeric';
  if (yIsNonNumericGroup && yCol !== xCol && !stackCol) {
    if (type === 'bar' || type === 'barH') {
      buildStackedChart(xCol, yCol, limit, type, titleInput);
      return;
    }
    showBuilderValidation(
      `„Wartość Y” („${yCol}”) to kolumna tekstowa, więc na wykresie ${type === 'line' ? 'liniowym' : 'kołowym'} pokazujemy tylko liczbę ${getRowEntityLabel()} — sama treść „${yCol}” się nie zmieści.`,
      'Przełącz na słupkowy →',
      () => { document.getElementById('b-type').value = 'bar'; buildCustomChart(); }
    );
  }
  clearBuilderValidation();
  const isCountAgg = yCol === '_count_' || agg === 'count';
  const sourceData = getBuilderFilteredData();
  if (!sourceData.length) {
    showBuilderValidation('Ustawione filtry usunęły wszystkie wiersze — nie ma czego pokazać. Sprawdź wartości w filtrach (np. literówkę, zły zakres dat) albo usuń jeden z nich.');
    return;
  }
  clearBuilderValidation();

  const groups = {};
  sourceData.forEach(r => {
    const k = String(r[xCol] || '(empty)');
    if (isCountAgg) {
      groups[k] = (groups[k]||0) + 1;
    } else {
      const v = toNum(r[yCol]);
      if (v !== null) {
        if (!groups[k]) groups[k] = { sum:0, count:0, max:-Infinity, min:Infinity };
        groups[k].sum += v;
        groups[k].count++;
        if (v > groups[k].max) groups[k].max = v;
        if (v < groups[k].min) groups[k].min = v;
      }
    }
  });

  let entries = Object.entries(groups).map(([k, v]) => {
    let val;
    if (isCountAgg) { val = v; }
    else {
      if (!v) { val = 0; }
      else if (agg === 'sum') val = v.sum;
      else if (agg === 'avg') val = v.sum / v.count;
      else if (agg === 'max') val = v.max;
      else if (agg === 'min') val = v.min;
      else val = v.sum;
    }
    return [k, Math.round(val * 100) / 100];
  }).sort((a,b) => b[1]-a[1]);

  if (limit > 0) entries = entries.slice(0, limit);

  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);

  const aggLabels = { sum:'Sum', count:'Count', avg:'Average', max:'Maximum', min:'Minimum' };
  const aggLabelsPl = { sum:'Suma', count:'Liczba ' + getRowEntityLabel(), avg:'Średnia', max:'Maksimum', min:'Minimum' };
  const entity = getRowEntityLabel();
  const autoTitle = isCountAgg ? `Liczba ${entity} według ${xCol}` : `${aggLabelsPl[agg]} kolumny „${yCol}” według ${xCol}`;
  const chartTitle = titleInput || autoTitle;
  const unit = isCountAgg ? '' : (getColumnUnit(yCol) ? ' ' + getColumnUnit(yCol) : '');
  const countsNote = (!isCountAgg && agg === 'avg')
    ? ' · liczba rekordów na kategorię: ' + entries.map(([k]) => `${k}=${groups[k]?.count ?? 0}`).slice(0, 6).join(', ') + (entries.length > 6 ? '...' : '')
    : '';
  const plainSub = isCountAgg
    ? `Oś X: „${xCol}” · Oś Y: liczba ${entity} — to ${entity}, NIE sztuki towaru · ${labels.length} kategorii`
    : `Oś X: „${xCol}” · Oś Y: ${aggLabelsPl[agg].toLowerCase()} kolumny „${yCol}”${unit} — wynik pracy dla tej wartości · ${labels.length} kategorii${countsNote}`;

  const id = `cc_${customChartCount++}`;
  const isH = type === 'barH';
  const isDoughnut = type === 'pie' || type === 'doughnut';
  const isLine = type === 'line';
  const h = isH ? Math.max(200, labels.length * 32 + 60) : 260;

  const area = document.getElementById('custom-charts-area');
  const wrap = document.createElement('div');
  wrap.className = 'chart-card';
  wrap.style.marginBottom = '16px';
  wrap.innerHTML = `<div style="display:flex;align-items:start;justify-content:space-between;margin-bottom:4px;">
    <div>
      <div class="chart-card-title">${chartTitle}</div>
      <div class="chart-card-sub">${plainSub}</div>
    </div>
    <button onclick="removeChart(this,'${id}')" style="background:none;border:none;cursor:pointer;font-size:16px;color:var(--ink-dim);">✕</button>
  </div>
  <div style="position:relative;width:100%;height:${h}px;">
    <canvas id="${id}" role="img" aria-label="${chartTitle}"></canvas>
  </div>
  ${insightDiv(labels, data, { trend: isLine })}
  <div style="font-size:10px;color:var(--ink-dim);margin-top:4px;">💡 Kliknij słupek/segment, aby zobaczyć dokładne wiersze w Tabeli</div>`;
  area.insertBefore(wrap, area.firstChild);

  setTimeout(() => createChart(id, type, labels, data, chartTitle, isH, !isDoughnut, xCol, isCountAgg ? null : getColumnUnit(yCol)), 50);
}

function buildStackedChart(xCol, stackCol, limit, type, titleInput) {
  const sourceData = getBuilderFilteredData();
  if (!sourceData.length) {
    showBuilderValidation('Ustawione filtry usunęły wszystkie wiersze — nie ma czego pokazać. Sprawdź wartości w filtrach (np. literówkę, zły zakres dat) albo usuń jeden z nich.');
    return;
  }
  clearBuilderValidation();
  // Build crosstab: counts[xValue][stackValue] = count
  const xTotals = {};
  const stackTotals = {};
  const crosstab = {};
  sourceData.forEach(r => {
    const x = String(r[xCol] || '(empty)');
    const s = String(r[stackCol] || '(empty)');
    if (!crosstab[x]) crosstab[x] = {};
    crosstab[x][s] = (crosstab[x][s] || 0) + 1;
    xTotals[x] = (xTotals[x] || 0) + 1;
    stackTotals[s] = (stackTotals[s] || 0) + 1;
  });

  let xLabels = Object.entries(xTotals).sort((a,b) => b[1]-a[1]).map(e => e[0]);
  if (limit > 0) xLabels = xLabels.slice(0, limit);

  // Cap number of stack series to keep the legend readable; fold the rest into "Other"
  const MAX_STACKS = 8;
  let stackKeys = Object.entries(stackTotals).sort((a,b) => b[1]-a[1]).map(e => e[0]);
  let otherKeys = [];
  if (stackKeys.length > MAX_STACKS) {
    otherKeys = stackKeys.slice(MAX_STACKS);
    stackKeys = stackKeys.slice(0, MAX_STACKS);
  }

  const datasets = stackKeys.map((s, i) => ({
    label: s,
    data: xLabels.map(x => (crosstab[x] && crosstab[x][s]) || 0),
    backgroundColor: COLORS[i % COLORS.length],
    borderWidth: 0,
  }));
  if (otherKeys.length) {
    datasets.push({
      label: 'Other',
      data: xLabels.map(x => otherKeys.reduce((sum, s) => sum + ((crosstab[x] && crosstab[x][s]) || 0), 0)),
      backgroundColor: '#94a3b8',
      borderWidth: 0,
    });
  }

  const chartTitle = titleInput || `${xCol} w podziale na ${stackCol}`;
  const id = `cc_${customChartCount++}`;
  const isH = type === 'barH';
  const h = isH ? Math.max(200, xLabels.length * 32 + 60) : 280;

  const area = document.getElementById('custom-charts-area');
  const wrap = document.createElement('div');
  wrap.className = 'chart-card';
  wrap.style.marginBottom = '16px';
  wrap.innerHTML = `<div style="display:flex;align-items:start;justify-content:space-between;margin-bottom:4px;">
    <div>
      <div class="chart-card-title">${chartTitle}</div>
      <div class="chart-card-sub">Oś X: „${xCol}” · Oś Y: liczba ${getRowEntityLabel()}, podzielona kolorami według „${stackCol}” · ${xLabels.length} kategorii</div>
    </div>
    <button onclick="removeChart(this,'${id}')" style="background:none;border:none;cursor:pointer;font-size:16px;color:var(--ink-dim);">✕</button>
  </div>
  <div style="position:relative;width:100%;height:${h}px;">
    <canvas id="${id}" role="img" aria-label="${chartTitle}"></canvas>
  </div>
  ${insightDiv(xLabels, xLabels.map(x => xTotals[x] || 0))}`;
  area.insertBefore(wrap, area.firstChild);

  setTimeout(() => {
    const canvas = document.getElementById(id);
    if (!canvas) return;
    const axisKey = isH ? 'x' : 'y';
    const catKey = isH ? 'y' : 'x';
    const instance = new Chart(canvas, {
      type: 'bar',
      data: { labels: xLabels, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: isH ? 'y' : 'x',
        plugins: {
          legend: { display: true, position: 'top', labels: { font: { size: 11 } } },
          tooltip: { callbacks: { label: ctx => ` ${ctx.dataset.label}: ${fmt(ctx.parsed[isH ? 'x' : 'y'])}` } }
        },
        scales: {
          [axisKey]: { stacked: true, grid: { color: 'rgba(0,0,0,0.05)' }, ticks: { font: { size: 11 }, precision: 0 } },
          [catKey]: { stacked: true, grid: { display: false }, ticks: { font: { size: 10 }, maxRotation: isH ? 0 : 45 } }
        }
      }
    });
    chartInstances[id] = instance;
  }, 50);
}

function removeChart(btn, id) {
  if (chartInstances[id]) { chartInstances[id].destroy(); delete chartInstances[id]; }
  btn.closest('.chart-card').remove();
}

// ─── TABLE ────────────────────────────────────────────────────
function renderFilterSelects() {
  const fc = document.getElementById('filter-col');
  const filterable = cols.filter(c => {
    if (colTypes[c] === 'categorical' || colTypes[c] === 'date') return true;
    if (colTypes[c] === 'text') {
      const uniq = new Set(allData.map(r => String(r[c] ?? ''))).size;
      return uniq > 0 && uniq <= 200;
    }
    return false;
  });
  fc.innerHTML = `<option value="">Filter by column — all</option>` +
    filterable.map(c => `<option value="${escAttr(c)}">${escAttr(c)}</option>`).join('');
  document.getElementById('filter-val').innerHTML = '<option value="">All values</option>';
}

document.getElementById('filter-col').addEventListener('change', function() {
  const col = this.value;
  const fv = document.getElementById('filter-val');
  if (!col) { 
    fv.innerHTML = '<option value="">All values</option>'; 
    filterTable(); 
    return; 
  }
  const uniq = [...new Set(allData.map(r => String(r[col] || '')))].sort();
  fv.innerHTML = `<option value="">All values</option>` + uniq.map(v => `<option value="${escAttr(v)}">${escAttr(v)}</option>`).join('');
  filterTable();
});

function filterTable() {
  const q = document.getElementById('search-input').value.toLowerCase().trim();
  const fc = document.getElementById('filter-col').value;
  const fv = document.getElementById('filter-val').value;
  filteredData = allData.filter(row => {
    if (fc && fv && String(row[fc] || '') !== fv) return false;
    if (q) return cols.some(c => String(row[c] || '').toLowerCase().includes(q));
    return true;
  });
  currentPage = 1;
  renderTable();
}

function renderTable() {
  const head = document.getElementById('table-head');
  const body = document.getElementById('table-body');
  const dispCols = cols.slice(0, 20);

  head.innerHTML = `<tr>${dispCols.map(c => {
    const cls = sortCol === c ? (sortDir === 1 ? 'sort-asc' : 'sort-desc') : '';
    return `<th class="${cls}" onclick="sortByCol('${c.replace(/'/g, "\\'")}')">${c}</th>`;
  }).join('')}</tr>`;

  const start = (currentPage - 1) * PAGE_SIZE;
  const page = filteredData.slice(start, start + PAGE_SIZE);

  body.innerHTML = page.map(row =>
    `<tr>${dispCols.map(c => `<td title="${String(row[c] || '').replace(/"/g,'&quot;')}">${String(row[c] || '').slice(0,60)}</td>`).join('')}</tr>`
  ).join('') || `<tr><td colspan="${dispCols.length}" style="text-align:center;color:var(--ink-dim);padding:32px;">No results found</td></tr>`;

  const totalPages = Math.ceil(filteredData.length / PAGE_SIZE);
  document.getElementById('page-info').textContent =
    `${filteredData.length.toLocaleString('en-US')} rows${filteredData.length < allData.length ? ` (of ${allData.length.toLocaleString('en-US')})` : ''} · Page ${currentPage}/${totalPages||1}`;

  const pb = document.getElementById('page-btns');
  let btns = '';
  const pages = getPageRange(currentPage, totalPages);
  pages.forEach(p => {
    if (p === '…') btns += `<span class="page-btn" style="cursor:default;color:var(--ink-dim);">…</span>`;
    else btns += `<button class="page-btn ${p===currentPage?'active':''}" onclick="goPage(${p})">${p}</button>`;
  });
  pb.innerHTML = (currentPage > 1 ? `<button class="page-btn" onclick="goPage(${currentPage-1})">‹</button>` : '') +
    btns +
    (currentPage < totalPages ? `<button class="page-btn" onclick="goPage(${currentPage+1})">›</button>` : '');
}

function getPageRange(cur, total) {
  if (total <= 7) return Array.from({length:total},(_,i)=>i+1);
  const r = new Set([1, total, cur]);
  if (cur > 1) r.add(cur-1);
  if (cur < total) r.add(cur+1);
  const sorted = [...r].sort((a,b)=>a-b);
  const result = [];
  sorted.forEach((p,i) => {
    if (i > 0 && p - sorted[i-1] > 1) result.push('…');
    result.push(p);
  });
  return result;
}

function goPage(p) { currentPage = p; renderTable(); }

function parseDateVal(v) {
  if (v == null || v === '') return null;
  const s = String(v).trim();
  let d = new Date(s);
  if (!isNaN(d.getTime())) return d.getTime();
  // Fallback for dd.mm.yyyy / dd-mm-yyyy / dd/mm/yyyy (common outside the US)
  const m = s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{2,4})/);
  if (m) {
    let [, dd, mm, yy] = m;
    if (yy.length === 2) yy = (parseInt(yy, 10) < 50 ? '20' : '19') + yy;
    d = new Date(`${yy}-${mm.padStart(2,'0')}-${dd.padStart(2,'0')}`);
    if (!isNaN(d.getTime())) return d.getTime();
  }
  return null;
}

function sortByCol(col) {
  if (sortCol === col) sortDir *= -1; else { sortCol = col; sortDir = 1; }
  const isNumericCol = colTypes[col] === 'numeric';
  const isDateCol = colTypes[col] === 'date';
  filteredData.sort((a, b) => {
    let av, bv;
    if (isNumericCol) {
      av = toNum(a[col]); bv = toNum(b[col]);
      if (av === null && bv === null) return 0;
      if (av === null) return 1;
      if (bv === null) return -1;
    } else if (isDateCol) {
      av = parseDateVal(a[col]); bv = parseDateVal(b[col]);
      if (av === null && bv === null) return 0;
      if (av === null) return 1;
      if (bv === null) return -1;
    } else {
      av = String(a[col] ?? '').toLowerCase();
      bv = String(b[col] ?? '').toLowerCase();
    }
    return av < bv ? -sortDir : av > bv ? sortDir : 0;
  });
  currentPage = 1;
  renderTable();
}

// ─── EXPORT ───────────────────────────────────────────────────
function exportFiltered() {
  const ws = XLSX.utils.json_to_sheet(filteredData);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, 'Filtered');
  XLSX.writeFile(wb2, 'export_' + new Date().toISOString().slice(0,10) + '.xlsx');
}

// ─── TABS ─────────────────────────────────────────────────────
function switchTab(name, el) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('tab-' + name).classList.add('active');
  if (typeof renderOpsCenter === 'function') renderOpsCenter();
}

// ─── PACK OPERATIONS CENTER (frontend-only, localStorage) ──────
const OPS_STORE_KEY = 'pack_ops_center_v1';

function defaultOpsStore() {
  return {
    version: 'hybrid-v1.3',
    updatedAt: new Date().toISOString(),
    excel: { fileName: '', sheetName: '', rows: [], cols: [], colTypes: {} },
    incidents: [],
    knowledge: [],
    actions: [],
    trainings: [],
    settings: {},
    ui: {}
  };
}

function normalizeOpsStore(parsed) {
  const base = defaultOpsStore();
  const incoming = parsed && typeof parsed === 'object' ? parsed : {};
  return {
    ...base,
    ...incoming,
    excel: { ...base.excel, ...(incoming.excel || {}) },
    incidents: Array.isArray(incoming.incidents) ? incoming.incidents : [],
    knowledge: Array.isArray(incoming.knowledge) ? incoming.knowledge : [],
    actions: Array.isArray(incoming.actions) ? incoming.actions : [],
    trainings: Array.isArray(incoming.trainings) ? incoming.trainings : [],
    settings: incoming.settings && typeof incoming.settings === 'object' ? incoming.settings : {},
    ui: incoming.ui && typeof incoming.ui === 'object' ? incoming.ui : {}
  };
}

function loadOpsStore() {
  try {
    const raw = localStorage.getItem(OPS_STORE_KEY);
    return raw ? normalizeOpsStore(JSON.parse(raw)) : defaultOpsStore();
  } catch (e) {
    return defaultOpsStore();
  }
}

function saveOpsStore(store) {
  const normalized = normalizeOpsStore(store);
  normalized.updatedAt = new Date().toISOString();
  localStorage.setItem(OPS_STORE_KEY, JSON.stringify(normalized));
}

function saveCurrentWorkspaceSnapshot() {
  const store = loadOpsStore();
  store.excel = {
    fileName: currentFileName || (document.getElementById('ib-name') ? document.getElementById('ib-name').textContent : '') || '',
    sheetName: currentSheetName || '',
    rows: Array.isArray(allData) ? allData : [],
    cols: Array.isArray(cols) ? cols : [],
    colTypes: colTypes || {}
  };
  saveOpsStore(store);
}

function applyWorkspaceExcel(store) {
  const excel = store && store.excel ? store.excel : null;
  if (!excel || !Array.isArray(excel.rows) || !excel.rows.length) return false;
  allData = excel.rows;
  filteredData = [...allData];
  cols = Array.isArray(excel.cols) && excel.cols.length ? excel.cols : Object.keys(allData[0] || {});
  colTypes = excel.colTypes && Object.keys(excel.colTypes).length ? excel.colTypes : {};
  if (!Object.keys(colTypes).length) cols.forEach(c => colTypes[c] = detectType(c, allData));
  currentFileName = excel.fileName || 'Workspace JSON';
  currentSheetName = excel.sheetName || '';

  const nameEl = document.getElementById('ib-name');
  const headEl = document.getElementById('header-filename');
  const rowsEl = document.getElementById('ib-rows');
  const colsEl = document.getElementById('ib-cols');
  if (nameEl) nameEl.textContent = currentFileName;
  if (headEl) headEl.textContent = currentFileName;
  if (rowsEl) rowsEl.textContent = allData.length.toLocaleString('en-US');
  if (colsEl) colsEl.textContent = cols.length;

  const hasTimeCol = !!(findCol(/time|hour|godzina/i) || cols.filter(c => colTypes[c] === 'date')[0]);
  const hasDateCol = cols.some(c => colTypes[c] === 'date');
  const shiftWrap = document.getElementById('shift-wrap');
  const periodWrap = document.getElementById('period-wrap');
  const sheetWrap = document.getElementById('ib-sheet-wrap');
  if (shiftWrap) shiftWrap.style.display = hasTimeCol ? 'flex' : 'none';
  if (periodWrap) periodWrap.style.display = hasDateCol ? 'flex' : 'none';
  if (sheetWrap) sheetWrap.style.display = 'none';

  const upload = document.getElementById('upload-screen');
  const app = document.getElementById('app');
  if (upload) upload.style.display = 'none';
  if (app) app.style.display = 'block';
  return true;
}

function restoreWorkspaceOnStartup() {
  const store = loadOpsStore();
  const restoredExcel = applyWorkspaceExcel(store);
  if (restoredExcel) {
    try { renderAll(); } catch (e) { try { renderOpsCenter(); } catch (_) {} }
    return true;
  }
  try { renderOpsCenter(); } catch(e) {}
  return false;
}

function opsNowDate() { return new Date().toISOString().slice(0, 10); }
function opsId(prefix) { return prefix + '-' + new Date().getFullYear() + '-' + Math.random().toString(36).slice(2, 7).toUpperCase(); }
function opsVal(id) { const el = document.getElementById(id); return el ? el.value.trim() : ''; }
function opsNum(id) { const n = parseFloat(opsVal(id)); return isNaN(n) ? null : n; }
function opsNorm(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }

function findOpsColumn(patterns) {
  return cols.find(c => patterns.some(re => re.test(c))) || null;
}

function inferExcelIncidents() {
  if (!allData || !allData.length) return [];
  const categoryCol = findOpsColumn([/category|kategoria|typ|type|problem|error|błąd|blad|usterk|awari|status/i]);
  const areaCol = findOpsColumn([/area|obszar|strefa|zone|station|stacja|workstation/i]);
  const deviceCol = findOpsColumn([/device|urząd|urzad|printer|drukark|scanner|skaner|panel|fortna|sap/i]);
  const statusCol = findOpsColumn([/status|state|stan/i]);
  const dateCol = cols.find(c => colTypes[c] === 'date') || findOpsColumn([/date|data/i]);
  const responseCol = findOpsColumn([/response|reakcj|czas.?reakcj/i]);
  const resolutionCol = findOpsColumn([/resolution|napraw|czas.?rozwiąz|czas.?rozwiaz|mttr/i]);
  const priorityCol = findOpsColumn([/priority|priorytet|critical|severity/i]);
  const ownerCol = findOpsColumn([/owner|assigned|au|pracownik|operator|user/i]);

  const looksLikeIncidentFile = categoryCol || statusCol || responseCol || resolutionCol || deviceCol;
  if (!looksLikeIncidentFile) return [];

  return allData.map((r, i) => ({
    id: 'XLS-' + String(i + 1).padStart(5, '0'),
    source: 'Excel',
    date: dateCol ? String(r[dateCol] || '').slice(0, 10) : '',
    area: areaCol ? String(r[areaCol] || '') : '',
    device: deviceCol ? String(r[deviceCol] || '') : '',
    category: categoryCol ? String(r[categoryCol] || '') : 'Excel row',
    priority: priorityCol ? String(r[priorityCol] || '') : '',
    status: statusCol ? String(r[statusCol] || '') : '',
    owner: ownerCol ? String(r[ownerCol] || '') : '',
    response: responseCol ? toNum(r[responseCol]) : null,
    resolution: resolutionCol ? toNum(r[resolutionCol]) : null,
    description: '', root: '', solution: ''
  })).filter(x => x.category || x.device || x.area || x.status);
}

function getAllOpsIncidents() {
  const store = loadOpsStore();
  return [...store.incidents, ...inferExcelIncidents()];
}

function addIncident() {
  const store = loadOpsStore();
  const incident = {
    id: opsId('PACK'),
    source: 'Local',
    date: opsNowDate(),
    area: opsVal('inc-area'),
    device: opsVal('inc-device'),
    priority: opsVal('inc-priority') || 'Medium',
    status: opsVal('inc-status') || 'Open',
    category: opsVal('inc-category') || 'General',
    owner: opsVal('inc-owner'),
    response: opsNum('inc-response'),
    resolution: opsNum('inc-resolution'),
    root: opsVal('inc-root'),
    solution: opsVal('inc-solution'),
    description: opsVal('inc-description')
  };
  store.incidents.unshift(incident);
  saveOpsStore(store);
  ['inc-area','inc-device','inc-category','inc-owner','inc-response','inc-resolution','inc-root','inc-solution','inc-description'].forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
  renderOpsCenter();
}

function deleteIncident(id) {
  const store = loadOpsStore();
  store.incidents = store.incidents.filter(x => x.id !== id);
  saveOpsStore(store);
  renderOpsCenter();
}

function addKnowledgeArticle() {
  const store = loadOpsStore();
  const art = {
    id: opsId('KB'),
    date: opsNowDate(),
    title: opsVal('kb-title') || 'Untitled article',
    area: opsVal('kb-area'),
    type: opsVal('kb-type') || 'Instruction',
    tags: opsVal('kb-tags'),
    problem: opsVal('kb-problem'),
    solution: opsVal('kb-solution')
  };
  store.knowledge.unshift(art);
  saveOpsStore(store);
  ['kb-title','kb-area','kb-tags','kb-problem','kb-solution'].forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
  renderOpsCenter();
}

function deleteKnowledgeArticle(id) {
  const store = loadOpsStore();
  store.knowledge = store.knowledge.filter(x => x.id !== id);
  saveOpsStore(store);
  renderOpsCenter();
}

function avg(vals) {
  const a = vals.filter(v => typeof v === 'number' && !isNaN(v));
  return a.length ? a.reduce((s, v) => s + v, 0) / a.length : null;
}

function fmtMin(v) { return v == null ? '—' : Math.round(v) + ' min'; }
function statusClass(s) { s = opsNorm(s); return /critical|high|open|escalated/.test(s) ? 'red' : /progress|medium|waiting/.test(s) ? 'amber' : 'green'; }

function recurringIssues() {
  const map = {};
  getAllOpsIncidents().forEach(x => {
    const key = [x.category, x.device, x.area].map(opsNorm).filter(Boolean).join(' · ') || opsNorm(x.description) || 'general';
    if (!map[key]) map[key] = { key, count: 0, last: '', response: [], resolution: [], solution: false, items: [] };
    const g = map[key];
    g.count++;
    if (x.date && x.date > g.last) g.last = x.date;
    if (typeof x.response === 'number') g.response.push(x.response);
    if (typeof x.resolution === 'number') g.resolution.push(x.resolution);
    if (x.solution) g.solution = true;
    g.items.push(x);
  });
  return Object.values(map).sort((a, b) => b.count - a.count).filter(g => g.count > 1).slice(0, 20);
}

function renderOpsKpis() {
  const host = document.getElementById('ops-kpi-grid');
  if (!host) return;
  const incidents = getAllOpsIncidents();
  const kb = loadOpsStore().knowledge;
  const open = incidents.filter(x => !/closed|resolved|zamk/i.test(String(x.status))).length;
  const escalated = incidents.filter(x => /escal/i.test(String(x.status)) || /fortna|vendor/i.test(String(x.owner))).length;
  const rec = recurringIssues().length;
  const response = avg(incidents.map(x => x.response));
  const resolution = avg(incidents.map(x => x.resolution));
  host.innerHTML = [
    ['Incidents', incidents.length, 'Excel + local register'],
    ['Open', open, 'do obsługi / sprawdzenia'],
    ['Escalated', escalated, 'AU → Fortna / vendor'],
    ['Recurring', rec, 'powtarzalne problemy'],
    ['Avg response', fmtMin(response), 'czas reakcji'],
    ['Avg resolution', fmtMin(resolution), 'czas rozwiązania'],
    ['KB articles', kb.length, 'instrukcje / checklisty'],
    ['Readiness', readinessScore() + '%', 'gotowość procesu']
  ].map(k => `<div class="kpi-card"><div class="kpi-label">${escAttr(k[0])}</div><div class="kpi-value">${escAttr(k[1])}</div><div class="kpi-meta">${escAttr(k[2])}</div></div>`).join('');
}

function renderRecurringIssues() {
  const host = document.getElementById('recurring-issues-list');
  if (!host) return;
  const list = recurringIssues();
  host.innerHTML = list.length ? list.map(g => `
    <div class="ops-list-item">
      <div class="ops-list-title"><span>${escAttr(g.key || 'Problem')}</span><span class="ops-tag amber">${g.count}x</span></div>
      <div class="ops-muted">Last seen: ${escAttr(g.last || '—')} · Avg response: ${fmtMin(avg(g.response))} · Avg resolution: ${fmtMin(avg(g.resolution))}</div>
      <div>${g.solution ? '<span class="ops-tag green">Solution exists</span>' : '<span class="ops-tag red">Need KB article</span>'}</div>
    </div>`).join('') : '<div class="empty">Brak powtarzalnych problemów. Po imporcie danych lub dodaniu incydentów lista utworzy się automatycznie.</div>';
}

function renderIncidentTable() {
  const host = document.getElementById('incident-table-wrap');
  if (!host) return;
  const localIds = new Set(loadOpsStore().incidents.map(x => x.id));
  const rows = getAllOpsIncidents().slice(0, 300);
  host.innerHTML = `<table><thead><tr>
    <th>ID</th><th>Date</th><th>Status</th><th>Priority</th><th>Area</th><th>Device</th><th>Category</th><th>Response</th><th>Resolution</th><th>Root cause</th><th>Solution</th><th></th>
  </tr></thead><tbody>${rows.map(x => `<tr>
    <td>${escAttr(x.id)}</td><td>${escAttr(x.date || '—')}</td><td><span class="ops-tag ${statusClass(x.status)}">${escAttr(x.status || '—')}</span></td><td>${escAttr(x.priority || '—')}</td><td>${escAttr(x.area || '—')}</td><td>${escAttr(x.device || '—')}</td><td>${escAttr(x.category || '—')}</td><td>${fmtMin(x.response)}</td><td>${fmtMin(x.resolution)}</td><td>${escAttr(x.root || '—')}</td><td>${escAttr(x.solution || '—')}</td><td>${localIds.has(x.id) ? `<button class="ops-delete" onclick="deleteIncident('${escAttr(x.id)}')">✕</button>` : ''}</td>
  </tr>`).join('') || `<tr><td colspan="12" style="text-align:center;color:var(--ink-dim);padding:32px;">No incidents yet</td></tr>`}</tbody></table>`;
}

function renderKnowledgeBase() {
  const host = document.getElementById('kb-list');
  if (!host) return;
  const store = loadOpsStore();
  const q = opsNorm(opsVal('kb-search'));
  const list = store.knowledge.filter(a => !q || [a.title,a.area,a.type,a.tags,a.problem,a.solution].some(v => opsNorm(v).includes(q)));
  host.innerHTML = list.length ? list.map(a => `
    <div class="ops-card">
      <div style="display:flex;justify-content:space-between;gap:8px;"><div class="ops-card-title">${escAttr(a.title)}</div><button class="ops-delete" onclick="deleteKnowledgeArticle('${escAttr(a.id)}')">✕</button></div>
      <div><span class="ops-tag amber">${escAttr(a.type)}</span><span class="ops-tag">${escAttr(a.area || 'Pack')}</span><span class="ops-tag">${escAttr(a.date)}</span></div>
      <div class="ops-card-body"><strong>Problem:</strong> ${escAttr(a.problem || '—')}\n<strong>Solution:</strong> ${escAttr(a.solution || '—')}</div>
      <div class="ops-muted">Tags: ${escAttr(a.tags || '—')}</div>
    </div>`).join('') : '<div class="empty">Knowledge Base jest pusta. Dodaj instrukcje, checklisty i rozwiązania typowych problemów.</div>';
  renderKbProgress();
}

function renderKbProgress() {
  const host = document.getElementById('kb-progress');
  if (!host) return;
  const kb = loadOpsStore().knowledge;
  const types = ['Instruction','Checklist','Quick reaction','Solution','Training material'];
  host.innerHTML = types.map(t => {
    const count = kb.filter(a => a.type === t).length;
    const pct = Math.min(100, count * 25);
    return `<div class="ops-progress-row"><span>${escAttr(t)}</span><div class="ops-bar"><span style="width:${pct}%"></span></div><strong>${count}</strong></div>`;
  }).join('');
}

function readinessScore() {
  const incidents = getAllOpsIncidents();
  const kb = loadOpsStore().knowledge;
  const rec = recurringIssues();
  let score = 0;
  score += kb.some(a => a.type === 'Instruction') ? 18 : 0;
  score += kb.some(a => a.type === 'Checklist') ? 18 : 0;
  score += kb.some(a => a.type === 'Quick reaction' || a.type === 'Solution') ? 18 : 0;
  score += incidents.length ? 16 : 0;
  score += incidents.some(x => typeof x.response === 'number') ? 10 : 0;
  score += incidents.some(x => /escal/i.test(String(x.status)) || /fortna/i.test(String(x.owner))) ? 8 : 0;
  score += rec.length ? 7 : 0;
  score += kb.some(a => a.type === 'Training material') ? 5 : 0;
  return Math.min(100, score);
}

function renderPerformance() {
  const goals = document.getElementById('goal-tracker');
  const ready = document.getElementById('readiness-box');
  const summary = document.getElementById('manager-summary');
  if (!goals || !ready || !summary) return;
  const incidents = getAllOpsIncidents();
  const store = loadOpsStore();
  const kb = store.knowledge;
  const rec = recurringIssues();
  const resolved = incidents.filter(x => /closed|resolved|zamk/i.test(String(x.status))).length;
  const avgResp = fmtMin(avg(incidents.map(x => x.response)));
  const progress = {
    bg1: Math.min(100, (incidents.length ? 35 : 0) + (avg(incidents.map(x => x.response)) != null ? 25 : 0) + (rec.length ? 20 : 0) + (incidents.some(x => /escal/i.test(String(x.status))) ? 20 : 0)),
    bg2: Math.min(100, kb.length * 12 + (kb.some(a => a.type === 'Checklist') ? 20 : 0) + (kb.some(a => a.type === 'Training material') ? 20 : 0)),
    dev: Math.min(100, (incidents.length ? 25 : 0) + (rec.length ? 25 : 0) + (kb.length ? 25 : 0) + (allData.length ? 25 : 0))
  };
  const goalCards = [
    ['Business Goal 1', progress.bg1, ['Incident Center','Response Time','Recurring Issues','Escalation evidence']],
    ['Business Goal 2', progress.bg2, ['Knowledge Base','Instructions','Checklists','Training materials']],
    ['Development Goal', progress.dev, ['Excel analytics','Problem analysis','Documentation standard','AU evidence dashboard']]
  ];
  goals.innerHTML = goalCards.map(g => `<div class="ops-goal-card"><h3>${g[0]}</h3><div class="ops-progress-row" style="grid-template-columns:1fr 46px;"><div class="ops-bar"><span style="width:${g[1]}%"></span></div><strong>${g[1]}%</strong></div><div class="ops-evidence">${g[2].map(x => '✔ ' + escAttr(x)).join('<br>')}</div></div>`).join('');
  const r = readinessScore();
  ready.innerHTML = `<div style="font-family:var(--mono);font-size:44px;font-weight:800;color:var(--amber);line-height:1;">${r}%</div><div class="ops-muted" style="margin:8px 0 14px;">Frontend-only readiness score. Dane są lokalne: Excel + localStorage.</div>
    ${[
      ['Documentation', kb.length ? Math.min(100, kb.length * 20) : 0],
      ['Incident process', incidents.length ? 85 : 20],
      ['Response metrics', avg(incidents.map(x => x.response)) != null ? 90 : 15],
      ['Recurring analysis', rec.length ? 90 : 30],
      ['Training materials', kb.some(a => a.type === 'Training material') ? 100 : 25]
    ].map(x => `<div class="ops-progress-row"><span>${x[0]}</span><div class="ops-bar"><span style="width:${x[1]}%"></span></div><strong>${x[1]}%</strong></div>`).join('')}`;
  summary.textContent = `Do startu magazynu przygotowałem frontendowy Pack Operations Center bez backendu i bez bazy danych. Narzędzie działa lokalnie w przeglądarce, analizuje pliki Excel oraz pozwala prowadzić lokalny rejestr incydentów, bazę wiedzy, checklisty i materiały szkoleniowe.\n\nAktualnie system pokazuje: ${incidents.length} incydentów / wpisów, ${resolved} rozwiązanych, ${rec.length} powtarzalnych problemów, ${kb.length} artykułów KB oraz średni czas reakcji: ${avgResp}. Dzięki temu po Go-Live proces zgłaszania, analizy, eskalacji i dokumentowania problemów będzie gotowy od pierwszego dnia.`;
}

function renderOpsCenter() {
  renderOpsKpis();
  renderRecurringIssues();
  renderIncidentTable();
  renderKnowledgeBase();
  renderPerformance();
}

function exportOpsData() {
  if (allData && allData.length) saveCurrentWorkspaceSnapshot();
  const data = loadOpsStore();
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'pack_ops_center_' + opsNowDate() + '.json';
  a.click();
  URL.revokeObjectURL(a.href);
}

function triggerOpsImport() {
  const input = document.getElementById('ops-import-input');
  if (input) input.click();
}

function importOpsData(event) {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const parsed = JSON.parse(e.target.result);
      const store = normalizeOpsStore(parsed);
      saveOpsStore(store);
      const restoredExcel = applyWorkspaceExcel(store);
      if (restoredExcel) renderAll();
      else renderOpsCenter();
      alert('Workspace imported successfully.');
    } catch (err) {
      alert('Could not import Workspace JSON: ' + err.message);
    }
  };
  reader.readAsText(file);
  event.target.value = '';
}

function clearOpsDataConfirm() {
  if (confirm('Clear local Pack Operations data from this browser? Excel file data will not be changed.')) {
    localStorage.removeItem(OPS_STORE_KEY);
    renderOpsCenter();
  }
}

// Initial render / restore Workspace if a previous session exists.
try { restoreWorkspaceOnStartup(); } catch(e) { try { renderOpsCenter(); } catch(_) {} }
