let planWorkbookData = null;
let toolWorkbookData = null;
let allRows = [];
let filteredRows = [];
let lineChart = null;
let totalChart = null;

const planFileInput = document.getElementById('planFile');
const toolFileInput = document.getElementById('toolFile');
const lineFilter = document.getElementById('lineFilter');
const toolFilter = document.getElementById('toolFilter');
const resetFilterBtn = document.getElementById('resetFilterBtn');
const messageBox = document.getElementById('messageBox');
const totalQtyEl = document.getElementById('totalQty');
const pivotTable = document.getElementById('pivotTable');
const detailTable = document.getElementById('detailTable');
const unmappedSummary = document.getElementById('unmappedSummary');
const unmappedList = document.getElementById('unmappedList');
const downloadPivotBtn = document.getElementById('downloadPivotBtn');
const downloadDetailBtn = document.getElementById('downloadDetailBtn');

planFileInput.addEventListener('change', async (event) => {
  planWorkbookData = await readExcelFile(event.target.files[0]);
  processIfReady();
});

toolFileInput.addEventListener('change', async (event) => {
  toolWorkbookData = await readExcelFile(event.target.files[0]);
  processIfReady();
});

lineFilter.addEventListener('change', applyFilters);
toolFilter.addEventListener('change', applyFilters);
resetFilterBtn.addEventListener('click', () => {
  [...lineFilter.options].forEach(opt => opt.selected = false);
  [...toolFilter.options].forEach(opt => opt.selected = false);
  applyFilters();
});

downloadPivotBtn.addEventListener('click', () => downloadTableAsCsv('pivotTable', 'production_pivot.csv'));
downloadDetailBtn.addEventListener('click', () => downloadTableAsCsv('detailTable', 'production_detail.csv'));

function setMessage(text, type = 'info') {
  messageBox.textContent = text;
  messageBox.className = `message ${type}`;
}

function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    if (!file) return resolve(null);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        resolve(workbook);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function processIfReady() {
  if (!planWorkbookData || !toolWorkbookData) {
    setMessage('생산계획 파일과 Tool 매핑 파일을 모두 업로드하세요.', 'info');
    return;
  }

  try {
    const planRows = parsePlanWorkbook(planWorkbookData);
    const toolMap = parseToolWorkbook(toolWorkbookData);
    allRows = mapTools(planRows, toolMap);
    buildFilterOptions(allRows);
    applyFilters();
    setMessage('업로드 및 분석이 완료되었습니다.', 'success');
  } catch (error) {
    console.error(error);
    setMessage(`처리 중 오류가 발생했습니다: ${error.message}`, 'error');
  }
}

function sheetToJson(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
}

function parsePlanWorkbook(workbook) {
  const firstSheetName = workbook.SheetNames[0];
  const rows = sheetToJson(workbook, firstSheetName);

  if (!rows.length) throw new Error('생산계획 엑셀에 데이터가 없습니다.');

  const headers = Object.keys(rows[0]);
  const lineCol = findHeader(headers, ['Line']);
  const modelCol = findHeader(headers, ['Model.Suffix', 'Model Suffix', 'ModelSuffix', 'suffix']);

  if (!lineCol) throw new Error('생산계획 파일에서 Line 컬럼을 찾을 수 없습니다.');
  if (!modelCol) throw new Error('생산계획 파일에서 Model.Suffix 컬럼을 찾을 수 없습니다.');

  const dateCols = headers.filter(col => String(col).includes('/'));
  if (!dateCols.length) throw new Error('생산계획 파일에서 날짜 컬럼을 찾을 수 없습니다. 예: 1/1, 1/2');

  const allowedLines = new Set(['KR1', 'KR2', 'KR3']);
  const melted = [];

  rows.forEach(row => {
    const line = String(row[lineCol] ?? '').trim();
    if (!allowedLines.has(line)) return;

    const model = normalizeModel(row[modelCol]);
    if (!model) return;

    dateCols.forEach(dateCol => {
      melted.push({
        Line: line,
        'Model.Suffix': model,
        Date: addDayLabel(String(dateCol).trim()),
        Qty: extractNumber(row[dateCol])
      });
    });
  });

  return melted;
}

function parseToolWorkbook(workbook) {
  const candidates = [];

  workbook.SheetNames.forEach(sheetName => {
    const rows = sheetToJson(workbook, sheetName);
    if (!rows.length) return;

    const headers = Object.keys(rows[0]).map(String);
    const suffixCol = headers.find(h => h.toLowerCase().includes('suffix'));
    const toolCol = headers.find(h => h.toLowerCase().includes('tool'));

    if (!suffixCol || !toolCol) return;

    rows.forEach(row => {
      const suffix = normalizeModel(row[suffixCol]);
      const tool = String(row[toolCol] ?? '').trim();
      if (suffix && tool) {
        candidates.push({ suffix, tool });
      }
    });
  });

  if (!candidates.length) throw new Error('Tool 매핑 파일에서 suffix / tool 컬럼을 찾을 수 없습니다.');

  // Model.Suffix 중복 제거: 먼저 나온 값 유지
  const unique = new Map();
  candidates.forEach(item => {
    if (!unique.has(item.suffix)) unique.set(item.suffix, item.tool);
  });

  // 부분 매칭 정확도 향상: 긴 suffix 우선
  return [...unique.entries()]
    .map(([suffix, tool]) => ({ suffix, tool }))
    .sort((a, b) => b.suffix.length - a.suffix.length);
}

function mapTools(planRows, toolMap) {
  return planRows.map(row => {
    const matched = toolMap.find(item => row['Model.Suffix'].includes(item.suffix));
    return {
      ...row,
      Tool: matched ? matched.tool : '미정'
    };
  });
}

function findHeader(headers, candidates) {
  const lowerMap = new Map(headers.map(h => [String(h).toLowerCase().replace(/\s+/g, ''), h]));
  for (const candidate of candidates) {
    const key = String(candidate).toLowerCase().replace(/\s+/g, '');
    if (lowerMap.has(key)) return lowerMap.get(key);
  }
  return null;
}

function normalizeModel(value) {
  return String(value ?? '').trim().toUpperCase();
}

function extractNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  const match = String(value).match(/-?\d+/);
  return match ? parseInt(match[0], 10) : 0;
}

function addDayLabel(dateStr) {
  try {
    const normalized = dateStr.replace(/^2026\//, '');
    const [month, day] = normalized.split('/').map(v => parseInt(v, 10));
    if (!month || !day) return dateStr;

    const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
    const dt = new Date(2026, month - 1, day);
    return `${month}/${day}(${dayNames[dt.getDay()]})`;
  } catch {
    return dateStr;
  }
}

function buildFilterOptions(rows) {
  fillMultiSelect(lineFilter, uniqueSorted(rows.map(r => r.Line)));
  fillMultiSelect(toolFilter, uniqueSorted(rows.map(r => r.Tool)));
}

function fillMultiSelect(selectEl, values) {
  selectEl.innerHTML = '';
  values.forEach(value => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    selectEl.appendChild(option);
  });
}

function selectedValues(selectEl) {
  return [...selectEl.selectedOptions].map(opt => opt.value);
}

function uniqueSorted(values) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), 'ko'));
}

function applyFilters() {
  if (!allRows.length) return;

  const selectedLines = selectedValues(lineFilter);
  const selectedTools = selectedValues(toolFilter);

  filteredRows = allRows.filter(row => {
    const lineOk = !selectedLines.length || selectedLines.includes(row.Line);
    const toolOk = !selectedTools.length || selectedTools.includes(row.Tool);
    return lineOk && toolOk;
  });

  renderAll();
}

function renderAll() {
  totalQtyEl.textContent = formatNumber(sum(filteredRows.map(r => r.Qty)));
  renderPivotTable(filteredRows);
  renderDetailTable(filteredRows);
  renderUnmapped(allRows);
  renderCharts(filteredRows);
}

function getDateOrder(rows) {
  return uniqueSorted(rows.map(r => r.Date)).sort(compareDateLabels);
}

function compareDateLabels(a, b) {
  const pa = parseDateLabel(a);
  const pb = parseDateLabel(b);
  return pa - pb;
}

function parseDateLabel(label) {
  const match = String(label).match(/(\d{1,2})\/(\d{1,2})/);
  if (!match) return 999999;
  return parseInt(match[1], 10) * 100 + parseInt(match[2], 10);
}

function renderPivotTable(rows) {
  const dates = getDateOrder(rows);
  const group = new Map();

  rows.forEach(row => {
    const key = `${row.Line}||${row.Tool}`;
    if (!group.has(key)) group.set(key, { Line: row.Line, Tool: row.Tool, values: {} });
    group.get(key).values[row.Date] = (group.get(key).values[row.Date] || 0) + row.Qty;
  });

  const sortedGroups = [...group.values()].sort((a, b) =>
    a.Line.localeCompare(b.Line, 'ko') || a.Tool.localeCompare(b.Tool, 'ko')
  );

  const thead = `<thead><tr><th>Line</th><th>Tool</th>${dates.map(d => `<th>${escapeHtml(d)}</th>`).join('')}<th>합계</th></tr></thead>`;
  const tbody = `<tbody>${sortedGroups.map(item => {
    const total = sum(dates.map(d => item.values[d] || 0));
    return `<tr><td>${escapeHtml(item.Line)}</td><td>${escapeHtml(item.Tool)}</td>${dates.map(d => `<td>${formatNumber(item.values[d] || 0)}</td>`).join('')}<td><strong>${formatNumber(total)}</strong></td></tr>`;
  }).join('')}</tbody>`;

  pivotTable.innerHTML = thead + tbody;
}

function renderDetailTable(rows) {
  const sorted = [...rows].sort((a, b) =>
    compareDateLabels(a.Date, b.Date) || a.Line.localeCompare(b.Line, 'ko') || a.Tool.localeCompare(b.Tool, 'ko')
  );

  const thead = '<thead><tr><th>Line</th><th>Model.Suffix</th><th>Date</th><th>Qty</th><th>Tool</th></tr></thead>';
  const tbody = `<tbody>${sorted.map(row => `
    <tr>
      <td>${escapeHtml(row.Line)}</td>
      <td>${escapeHtml(row['Model.Suffix'])}</td>
      <td>${escapeHtml(row.Date)}</td>
      <td>${formatNumber(row.Qty)}</td>
      <td>${escapeHtml(row.Tool)}</td>
    </tr>`).join('')}</tbody>`;

  detailTable.innerHTML = thead + tbody;
}

function renderUnmapped(rows) {
  const unmapped = uniqueSorted(rows.filter(r => r.Tool === '미정').map(r => r['Model.Suffix']));

  if (unmapped.length) {
    unmappedSummary.textContent = `${unmapped.length}개 모델 미매핑`;
    unmappedList.innerHTML = unmapped.map(model => `<span class="chip">${escapeHtml(model)}</span>`).join('');
  } else {
    unmappedSummary.textContent = '✅ Tool 매핑 완료';
    unmappedList.innerHTML = '';
  }
}

function renderCharts(rows) {
  const dates = getDateOrder(rows);
  const lineValues = uniqueSorted(rows.map(r => r.Line));

  const lineDatasets = lineValues.map(line => ({
    label: line,
    data: dates.map(date => sum(rows.filter(r => r.Line === line && r.Date === date).map(r => r.Qty))),
    tension: 0.25,
    pointRadius: 4
  }));

  const totalData = dates.map(date => sum(rows.filter(r => r.Date === date).map(r => r.Qty)));

  if (lineChart) lineChart.destroy();
  if (totalChart) totalChart.destroy();

  lineChart = new Chart(document.getElementById('lineChart'), {
    type: 'line',
    data: { labels: dates, datasets: lineDatasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom' } },
      scales: { x: { ticks: { maxRotation: 45, minRotation: 25 } }, y: { beginAtZero: true } }
    }
  });

  totalChart = new Chart(document.getElementById('totalChart'), {
    type: 'bar',
    data: { labels: dates, datasets: [{ label: '전체 생산량', data: totalData }] },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { x: { ticks: { maxRotation: 45, minRotation: 25 } }, y: { beginAtZero: true } }
    }
  });
}

function sum(values) {
  return values.reduce((acc, cur) => acc + (Number(cur) || 0), 0);
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString('ko-KR');
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function downloadTableAsCsv(tableId, filename) {
  const table = document.getElementById(tableId);
  if (!table || !table.rows.length) return;

  const rows = [...table.rows].map(row =>
    [...row.cells].map(cell => `"${cell.innerText.replaceAll('"', '""')}"`).join(',')
  );

  const bom = '\uFEFF';
  const blob = new Blob([bom + rows.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
