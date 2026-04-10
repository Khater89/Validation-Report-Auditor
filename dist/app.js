const state = {
  raw: [],
  filtered: [],
  fileMeta: null,
};

const el = id => document.getElementById(id);
const dropzone = el('dropzone');
const fileInput = el('fileInput');
const cards = el('cards');
const toolbarPanel = el('toolbarPanel');
const tablePanel = el('tablePanel');
const resultsBody = el('resultsBody');
const searchInput = el('searchInput');
const ticketTypeFilter = el('ticketTypeFilter');
const specificDateFilter = el('specificDateFilter');
const fromDateFilter = el('fromDateFilter');
const toDateFilter = el('toDateFilter');
const statusFilter = el('statusFilter');
const scopeChip = el('scopeChip');
const runtimeBanner = el('runtimeBanner');

function normalize(value) {
  return (value ?? '').toString().trim().replace(/\s+/g, ' ');
}

function normalizeName(value) {
  return normalize(value).toUpperCase();
}

function safeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function pad2(value) {
  return String(value).padStart(2, '0');
}

function showBanner(message) {
  if (!runtimeBanner) return;
  runtimeBanner.textContent = message;
  runtimeBanner.classList.remove('hidden');
}

function hideBanner() {
  if (!runtimeBanner) return;
  runtimeBanner.textContent = '';
  runtimeBanner.classList.add('hidden');
}

function formatDateParts(year, month, day) {
  return `${year}-${pad2(month)}-${pad2(day)}`;
}

function formatDateFromLocal(date) {
  return formatDateParts(date.getFullYear(), date.getMonth() + 1, date.getDate());
}

function formatDateFromUTC(date) {
  return formatDateParts(date.getUTCFullYear(), date.getUTCMonth() + 1, date.getUTCDate());
}

function excelSerialToDate(value) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return null;
  const days = Math.floor(value);
  const adjusted = days >= 60 ? days - 1 : days;
  const epoch = Date.UTC(1899, 11, 31);
  return new Date(epoch + adjusted * 86400000);
}

function parseDatePart(value) {
  if (value === undefined || value === null || value === '') return '';

  if (value instanceof Date) {
    return formatDateFromLocal(value);
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const parsed = excelSerialToDate(value);
    return parsed ? formatDateFromUTC(parsed) : '';
  }

  const text = normalize(value);
  if (!text) return '';

  const isoMatch = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if (isoMatch) {
    return formatDateParts(isoMatch[1], isoMatch[2], isoMatch[3]);
  }

  const slashMatch = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (slashMatch) {
    let year = slashMatch[3];
    if (year.length === 2) year = `20${year}`;
    const first = Number(slashMatch[1]);
    const second = Number(slashMatch[2]);
    const month = first > 12 ? second : first;
    const day = first > 12 ? first : second;
    return formatDateParts(year, month, day);
  }

  const monthWordMatch = text.match(/^([A-Za-z]{3,9})\s+(\d{1,2}),\s*(\d{4})$/);
  if (monthWordMatch) {
    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? text : formatDateFromLocal(parsed);
  }

  return text;
}

function parseTimePart(value) {
  if (value === undefined || value === null || value === '') return '';

  if (value instanceof Date) {
    return `${String(value.getHours()).padStart(2, '0')}:${String(value.getMinutes()).padStart(2, '0')}`;
  }

  if (typeof value === 'string') {
    const text = normalize(value).toUpperCase();
    const ampmMatch = text.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/);
    if (ampmMatch) {
      let hour = Number(ampmMatch[1]);
      const minute = ampmMatch[2];
      const suffix = ampmMatch[3];
      if (suffix === 'PM' && hour < 12) hour += 12;
      if (suffix === 'AM' && hour === 12) hour = 0;
      return `${String(hour).padStart(2, '0')}:${minute}`;
    }
    const timeMatch = text.match(/(\d{1,2}):(\d{2})/);
    if (timeMatch) {
      return `${String(Number(timeMatch[1])).padStart(2, '0')}:${timeMatch[2]}`;
    }
    return text;
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const totalMinutes = Math.round(value * 24 * 60);
    const hour = Math.floor(totalMinutes / 60) % 24;
    const minute = totalMinutes % 60;
    return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
  }

  return normalize(value);
}

function defaultTracking() {
  return {
    correctedMarked: false,
    correctedAt: '',
  };
}

function nowStamp() {
  return new Date().toLocaleString();
}

function recomputeRow(row) {
  const statusOk = normalize(row.ieFnStatus).toUpperCase() === 'OK';
  const dateComparable = Boolean(row.ieDueDate && row.fnScheduledDate);
  const timeComparable = Boolean(row.ieDueTime && row.fnScheduledTime);
  const nameComparable = Boolean(row.ieFeName || row.fnFeName);

  const rawDateMatch = !dateComparable ? true : row.ieDueDate === row.fnScheduledDate;
  const rawTimeMatch = !timeComparable ? true : row.ieDueTime === row.fnScheduledTime;

  const dateException = !rawDateMatch && row.ticketType === 'ITD';
  const timeException = !rawTimeMatch && (row.ticketType === 'ITD' || row.ticketType === 'ITR');

  const dateValid = !dateComparable ? true : (rawDateMatch || dateException);
  const timeValid = !timeComparable ? true : (rawTimeMatch || timeException);
  const nameValid = !nameComparable ? true : normalizeName(row.ieFeName) === normalizeName(row.fnFeName);

  row.statusOk = statusOk;
  row.dateValid = dateValid;
  row.timeValid = timeValid;
  row.nameValid = nameValid;
  row.dateException = dateException;
  row.timeException = timeException;
  row.exceptionApplied = dateException || timeException;
  row.exceptionType = [
    dateException ? 'ITD-Date' : '',
    timeException ? `${row.ticketType}-Time` : '',
  ].filter(Boolean).join(', ');

  const currentErrors = [
    !statusOk ? 'IE & FN Status is not OK' : '',
    !dateValid ? 'Due Date ≠ FN Scheduled Date' : '',
    !timeValid ? 'Due Time ≠ FN Scheduled Time' : '',
    !nameValid ? 'IE FE Name ≠ FN FE Name' : '',
  ].filter(Boolean);

  row.validationResult = currentErrors.length ? 'Error' : 'Matched';
  row.baseIssueDetails = currentErrors.length
    ? currentErrors.join('; ')
    : (row.exceptionApplied ? 'Matched by allowed ticket-type exception' : 'All checks passed');

  if (row.validationResult === 'Error' && row.tracking.correctedMarked) {
    row.finalResult = 'Corrected';
    row.trackingNote = `Marked corrected by user on ${row.tracking.correctedAt || '-'}`;
    row.issueDetails = `${row.baseIssueDetails}; Externally corrected and marked done by user`;
  } else {
    row.finalResult = row.validationResult;
    row.trackingNote = row.tracking.correctedMarked ? `Marked corrected by user on ${row.tracking.correctedAt || '-'}` : 'Pending follow-up';
    row.issueDetails = row.baseIssueDetails;
  }

  if (row.exceptionApplied && !row.issueDetails.includes(row.exceptionType)) {
    row.issueDetails += `; ${row.exceptionType}`;
  }
}

function getStorageKey() {
  if (!state.fileMeta) return null;
  const { name, size, lastModified } = state.fileMeta;
  return `validation-report-tracking:${name}:${size}:${lastModified}`;
}

function persistState() {
  const key = getStorageKey();
  if (!key) return;
  const payload = {
    emailTo: el('emailTo').value || '',
    emailCc: el('emailCc').value || '',
    emailSubject: el('emailSubject').value || '',
    filters: {
      search: searchInput.value || '',
      ticketType: ticketTypeFilter.value || 'all',
      specificDate: specificDateFilter.value || '',
      fromDate: fromDateFilter.value || '',
      toDate: toDateFilter.value || '',
      status: statusFilter.value || 'all',
    },
    tracking: state.raw.reduce((acc, row) => {
      if (row.tracking?.correctedMarked) {
        acc[row.rowId] = row.tracking;
      }
      return acc;
    }, {}),
  };
  localStorage.setItem(key, JSON.stringify(payload));
}

function loadState() {
  const key = getStorageKey();
  if (!key) return;
  const raw = localStorage.getItem(key);
  if (!raw) return;

  try {
    const payload = JSON.parse(raw);
    state.raw.forEach(row => {
      row.tracking = payload.tracking?.[row.rowId]
        ? { ...defaultTracking(), ...payload.tracking[row.rowId] }
        : defaultTracking();
      recomputeRow(row);
    });

    el('emailTo').value = payload.emailTo || '';
    el('emailCc').value = payload.emailCc || '';
    el('emailSubject').value = payload.emailSubject || '';

    searchInput.value = payload.filters?.search || '';
    ticketTypeFilter.value = payload.filters?.ticketType || 'all';
    specificDateFilter.value = payload.filters?.specificDate || '';
    fromDateFilter.value = payload.filters?.fromDate || '';
    toDateFilter.value = payload.filters?.toDate || '';
    statusFilter.value = payload.filters?.status || 'all';
  } catch (error) {
    console.error('Failed to load stored state', error);
  }
}

function defaultEmailSubject() {
  const fileName = state.fileMeta?.name ? state.fileMeta.name.replace(/\.[^.]+$/, '') : 'Validation Report';
  return `Final Validation Report - ${fileName}`;
}

function pill(type, label) {
  return `<span class="pill ${type}">${safeHtml(label)}</span>`;
}

function finalResultPill(row) {
  if (row.finalResult === 'Error') return pill('err', 'Error');
  if (row.finalResult === 'Corrected') return pill('corrected', 'Corrected');
  if (row.exceptionApplied) return pill('warn', 'Matched via exception');
  return pill('ok', 'Matched');
}

function actionButton(row) {
  if (row.finalResult === 'Corrected') {
    return `<button class="action-btn done" data-correct-row="${safeHtml(row.rowId)}" disabled>Corrected</button>`;
  }
  if (row.validationResult === 'Error') {
    return `<button class="action-btn mark" data-correct-row="${safeHtml(row.rowId)}">Correct</button>`;
  }
  return '<span class="muted">—</span>';
}

function getRowClass(row) {
  if (row.finalResult === 'Corrected') return 'corrected';
  if (row.finalResult === 'Error') return 'error';
  if (row.exceptionApplied) return 'exception';
  return 'matched';
}

function computeStats(rows) {
  const total = rows.length;
  const matched = rows.filter(row => row.finalResult === 'Matched').length;
  const corrected = rows.filter(row => row.finalResult === 'Corrected').length;
  const errors = rows.filter(row => row.finalResult === 'Error').length;
  const exceptions = rows.filter(row => row.exceptionApplied).length;
  const completion = total ? (((matched + corrected) / total) * 100).toFixed(1) : '0.0';
  const statusErrors = rows.filter(row => !row.statusOk).length;
  const dateErrors = rows.filter(row => !row.dateValid).length;
  const timeErrors = rows.filter(row => !row.timeValid).length;
  const nameErrors = rows.filter(row => !row.nameValid).length;
  return { total, matched, corrected, errors, exceptions, completion, statusErrors, dateErrors, timeErrors, nameErrors };
}

function buildCards(rows) {
  const stats = computeStats(rows);
  cards.innerHTML = `
    <div class="card"><div class="label">Visible Tickets</div><div class="value">${stats.total}</div></div>
    <div class="card success"><div class="label">Matched</div><div class="value">${stats.matched}</div></div>
    <div class="card corrected"><div class="label">Corrected</div><div class="value">${stats.corrected}</div></div>
    <div class="card error"><div class="label">Open Errors</div><div class="value">${stats.errors}</div></div>
    <div class="card warn"><div class="label">Exceptions Applied</div><div class="value">${stats.exceptions}</div></div>
    <div class="card"><div class="label">Completion</div><div class="value">${stats.completion}%</div></div>
    <div class="card"><div class="label">Status Errors</div><div class="value">${stats.statusErrors}</div></div>
    <div class="card"><div class="label">Date Errors</div><div class="value">${stats.dateErrors}</div></div>
    <div class="card"><div class="label">Time Errors</div><div class="value">${stats.timeErrors}</div></div>
    <div class="card"><div class="label">Name Errors</div><div class="value">${stats.nameErrors}</div></div>
  `;
  cards.classList.remove('hidden');
}

function populateTicketTypes(rows) {
  const currentValue = ticketTypeFilter.value || 'all';
  const types = [...new Set(rows.map(row => row.ticketType).filter(Boolean))].sort();
  ticketTypeFilter.innerHTML = '<option value="all">All ticket types</option>' +
    types.map(type => `<option value="${safeHtml(type)}">${safeHtml(type)}</option>`).join('');
  ticketTypeFilter.value = types.includes(currentValue) ? currentValue : 'all';
}

function renderTable(rows) {
  if (!rows.length) {
    resultsBody.innerHTML = `
      <tr>
        <td colspan="23" class="empty-state">No rows match the selected filters.</td>
      </tr>
    `;
  } else {
    resultsBody.innerHTML = rows.map(row => `
      <tr class="${getRowClass(row)}">
        <td>${safeHtml(row.sourceRowNumber)}</td>
        <td>${safeHtml(row.ticket)}</td>
        <td>${safeHtml(row.ticketType || '')}</td>
        <td>${safeHtml(row.assignedCompany || '')}</td>
        <td>${safeHtml(row.filterDate || '')}</td>
        <td>${safeHtml(row.ieFnStatus || '')}</td>
        <td>${safeHtml(row.ieFeName || '')}</td>
        <td>${safeHtml(row.ieDueDate || '')}</td>
        <td>${safeHtml(row.ieDueTime || '')}</td>
        <td>${safeHtml([row.ieScheduledDate, row.ieScheduledTime].filter(Boolean).join(' '))}</td>
        <td>${safeHtml(row.fnStatus || '')}</td>
        <td>${safeHtml(row.fnFeName || '')}</td>
        <td>${safeHtml(row.fnScheduledDate || '')}</td>
        <td>${safeHtml(row.fnScheduledTime || '')}</td>
        <td>${row.statusOk ? pill('ok', 'OK') : pill('err', 'Error')}</td>
        <td>${row.dateValid ? pill(row.dateException ? 'warn' : 'ok', row.dateException ? 'ITD' : 'OK') : pill('err', 'Error')}</td>
        <td>${row.timeValid ? pill(row.timeException ? 'warn' : 'ok', row.timeException ? row.ticketType : 'OK') : pill('err', 'Error')}</td>
        <td>${row.nameValid ? pill('ok', 'OK') : pill('err', 'Error')}</td>
        <td>${row.exceptionApplied ? pill('warn', row.exceptionType) : pill('ok', 'None')}</td>
        <td>${finalResultPill(row)}</td>
        <td class="details-cell">${safeHtml(row.issueDetails || '')}</td>
        <td class="note-cell">${safeHtml(row.trackingNote || '')}</td>
        <td>${actionButton(row)}</td>
      </tr>
    `).join('');
  }

  toolbarPanel.classList.remove('hidden');
  tablePanel.classList.remove('hidden');
}

function updateScopeChip() {
  scopeChip.textContent = `Showing ${state.filtered.length} of ${state.raw.length} tickets`;
}

function applyFilters() {
  const search = normalize(searchInput.value).toLowerCase();
  const type = ticketTypeFilter.value;
  const specificDate = specificDateFilter.value;
  const fromDate = fromDateFilter.value;
  const toDate = toDateFilter.value;
  const status = statusFilter.value;

  state.filtered = state.raw.filter(row => {
    const hitSearch = !search || [row.ticket, row.ieFeName, row.fnFeName].join(' ').toLowerCase().includes(search);
    const hitType = type === 'all' || row.ticketType === type;
    const rowDate = row.filterDate || '';
    const hitSpecificDate = !specificDate || rowDate === specificDate;
    const hitFromDate = !fromDate || (rowDate && rowDate >= fromDate);
    const hitToDate = !toDate || (rowDate && rowDate <= toDate);

    let hitStatus = true;
    if (status === 'errors') hitStatus = row.finalResult === 'Error';
    if (status === 'corrected') hitStatus = row.finalResult === 'Corrected';
    if (status === 'matched') hitStatus = row.finalResult === 'Matched';
    if (status === 'exception') hitStatus = row.exceptionApplied && row.finalResult === 'Matched';

    return hitSearch && hitType && hitSpecificDate && hitFromDate && hitToDate && hitStatus;
  });

  buildCards(state.filtered);
  renderTable(state.filtered);
  updateScopeChip();
  persistState();
}

function clearFilters() {
  searchInput.value = '';
  ticketTypeFilter.value = 'all';
  specificDateFilter.value = '';
  fromDateFilter.value = '';
  toDateFilter.value = '';
  statusFilter.value = 'all';
  applyFilters();
}

function parseReportSheet(rows) {
  const mapped = [];

  for (let index = 2; index < rows.length; index += 1) {
    const row = rows[index] || [];
    const ticket = row[0];
    if (!ticket) continue;

    const assignedCompany = normalize(row[12]);
    if (assignedCompany.toUpperCase() !== 'FIELD NATION') continue;

    const fnScheduledDate = parseDatePart(row[20]);
    const ieDueDate = parseDatePart(row[7]);

    const mappedRow = {
      rowId: `row-${index + 1}`,
      sourceRowNumber: index + 1,
      ticket,
      ticketType: normalize(row[5]).toUpperCase(),
      assignedCompany,
      ieFnStatus: normalize(row[26]),
      ieFeName: normalize(row[13]),
      fnFeName: normalize(row[23]),
      ieDueDate,
      ieDueTime: parseTimePart(row[8]),
      ieScheduledDate: parseDatePart(row[9]),
      ieScheduledTime: parseTimePart(row[10]),
      fnScheduledDate,
      fnScheduledTime: parseTimePart(row[21]),
      fnStatus: normalize(row[16]),
      filterDate: fnScheduledDate || ieDueDate || '',
      tracking: defaultTracking(),
    };

    recomputeRow(mappedRow);
    mapped.push(mappedRow);
  }

  return mapped;
}

function setSheetMeta(file) {
  state.fileMeta = {
    name: file.name,
    size: file.size,
    lastModified: file.lastModified,
  };
  if (!normalize(el('emailSubject').value)) {
    el('emailSubject').value = defaultEmailSubject();
  }
}

function handleWorkbook(file) {
  if (typeof XLSX === 'undefined') {
    showBanner('Excel library failed to load. Refresh the page and try again.');
    return;
  }

  hideBanner();
  setSheetMeta(file);
  const reader = new FileReader();
  reader.onload = event => {
    try {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, {
        type: 'array',
        cellDates: false,
      });
      const sheetName = workbook.SheetNames.includes('Report') ? 'Report' : workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        defval: '',
        dateNF: 'yyyy-mm-dd',
      });

      state.raw = parseReportSheet(rows);
      if (!state.raw.length) {
        cards.classList.add('hidden');
        toolbarPanel.classList.add('hidden');
        tablePanel.classList.add('hidden');
        showBanner('No comparable Field Nation rows were found. Confirm that the workbook contains a Report sheet and that Assigned Company is Field Nation.');
        return;
      }

      populateTicketTypes(state.raw);
      loadState();
      populateTicketTypes(state.raw);
      if (!normalize(el('emailSubject').value)) {
        el('emailSubject').value = defaultEmailSubject();
      }
      hideBanner();
      applyFilters();
    } catch (error) {
      console.error('Failed to read workbook', error);
      showBanner(`Failed to load the workbook: ${error.message || 'Unknown error'}`);
    }
  };
  reader.onerror = () => showBanner('The selected file could not be read. Please try again.');
  reader.readAsArrayBuffer(file);
}

function getRowById(rowId) {
  return state.raw.find(row => row.rowId === rowId);
}

function markCorrected(rowId) {
  const row = getRowById(rowId);
  if (!row || row.validationResult !== 'Error' || row.tracking.correctedMarked) return;
  row.tracking.correctedMarked = true;
  row.tracking.correctedAt = nowStamp();
  recomputeRow(row);
  persistState();
  applyFilters();
}

function autoWidth(worksheet, rows) {
  const widths = [];
  rows.forEach(row => {
    Object.values(row).forEach((value, index) => {
      const length = String(value ?? '').length;
      widths[index] = Math.min(Math.max(widths[index] || 12, length + 2), 40);
    });
  });
  worksheet['!cols'] = widths.map(width => ({ wch: width }));
}

function makeExportRows() {
  return state.raw.map(row => ({
    'Excel Row': row.sourceRowNumber,
    'Ticket': row.ticket,
    'Ticket Type': row.ticketType,
    'Assigned Company': row.assignedCompany,
    'Filter Date': row.filterDate,
    'IE & FN Status': row.ieFnStatus,
    'IE FE Name': row.ieFeName,
    'IE Due Date': row.ieDueDate,
    'IE Due Time': row.ieDueTime,
    'FN Status': row.fnStatus,
    'FN FE Name': row.fnFeName,
    'FN Scheduled Date': row.fnScheduledDate,
    'FN Scheduled Time': row.fnScheduledTime,
    'Status Check': row.statusOk ? 'OK' : 'Error',
    'Date Check': row.dateValid ? (row.dateException ? 'OK via ITD exception' : 'OK') : 'Error',
    'Time Check': row.timeValid ? (row.timeException ? `OK via ${row.ticketType} exception` : 'OK') : 'Error',
    'Name Check': row.nameValid ? 'OK' : 'Error',
    'Exception Applied': row.exceptionType || 'None',
    'Validation Result': row.validationResult,
    'Tracking Status': row.finalResult,
    'Tracking Note': row.trackingNote,
    'Details': row.issueDetails,
  }));
}

function buildWorkbook() {
  const workbook = XLSX.utils.book_new();
  const stats = computeStats(state.raw);
  const visibleStats = computeStats(state.filtered);
  const summaryRows = [
    ['Metric', 'All Data', 'Current Filter View'],
    ['Comparable Tickets', stats.total, visibleStats.total],
    ['Matched', stats.matched, visibleStats.matched],
    ['Corrected', stats.corrected, visibleStats.corrected],
    ['Open Errors', stats.errors, visibleStats.errors],
    ['Exceptions Applied', stats.exceptions, visibleStats.exceptions],
    ['Completion %', stats.completion, visibleStats.completion],
    ['Status Errors', stats.statusErrors, visibleStats.statusErrors],
    ['Date Errors', stats.dateErrors, visibleStats.dateErrors],
    ['Time Errors', stats.timeErrors, visibleStats.timeErrors],
    ['Name Errors', stats.nameErrors, visibleStats.nameErrors],
    ['Workbook', state.fileMeta?.name || '', state.fileMeta?.name || ''],
    ['Applied Search', searchInput.value || '-', searchInput.value || '-'],
    ['Specific Date Filter', specificDateFilter.value || '-', specificDateFilter.value || '-'],
    ['From Date Filter', fromDateFilter.value || '-', fromDateFilter.value || '-'],
    ['To Date Filter', toDateFilter.value || '-', toDateFilter.value || '-'],
    ['Status Filter', statusFilter.value || 'all', statusFilter.value || 'all'],
    ['Ticket Type Filter', ticketTypeFilter.value || 'all', ticketTypeFilter.value || 'all'],
    ['Generated At', new Date().toLocaleString(), new Date().toLocaleString()],
  ];

  const exportRows = makeExportRows();
  const visibleRows = state.filtered.map(row => ({
    'Excel Row': row.sourceRowNumber,
    'Ticket': row.ticket,
    'Ticket Type': row.ticketType,
    'Assigned Company': row.assignedCompany,
    'Filter Date': row.filterDate,
    'Tracking Status': row.finalResult,
    'Validation Result': row.validationResult,
    'Details': row.issueDetails,
    'Tracking Note': row.trackingNote,
  }));
  const openErrorsRows = exportRows.filter(row => row['Tracking Status'] === 'Error');
  const correctedRows = exportRows.filter(row => row['Tracking Status'] === 'Corrected');

  const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
  const finalSheet = XLSX.utils.json_to_sheet(exportRows);
  const visibleSheet = XLSX.utils.json_to_sheet(visibleRows.length ? visibleRows : [{ Info: 'No rows in current filter view' }]);
  const errorsSheet = XLSX.utils.json_to_sheet(openErrorsRows.length ? openErrorsRows : [{ Info: 'No open errors remaining' }]);
  const correctedSheet = XLSX.utils.json_to_sheet(correctedRows.length ? correctedRows : [{ Info: 'No corrected rows yet' }]);

  autoWidth(finalSheet, exportRows);
  autoWidth(visibleSheet, visibleRows.length ? visibleRows : [{ Info: 'No rows in current filter view' }]);
  autoWidth(errorsSheet, openErrorsRows.length ? openErrorsRows : [{ Info: 'No open errors remaining' }]);
  autoWidth(correctedSheet, correctedRows.length ? correctedRows : [{ Info: 'No corrected rows yet' }]);
  summarySheet['!cols'] = [{ wch: 30 }, { wch: 22 }, { wch: 22 }];

  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');
  XLSX.utils.book_append_sheet(workbook, finalSheet, 'Final_Report');
  XLSX.utils.book_append_sheet(workbook, visibleSheet, 'Current_Filter_View');
  XLSX.utils.book_append_sheet(workbook, correctedSheet, 'Corrected_Rows');
  XLSX.utils.book_append_sheet(workbook, errorsSheet, 'Open_Errors');

  return workbook;
}

function timestampStamp() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const mm = String(now.getMinutes()).padStart(2, '0');
  return `${y}${m}${d}_${hh}${mm}`;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 400);
}

function exportWorkbook() {
  if (!state.raw.length || typeof XLSX === 'undefined') return;
  const workbook = buildWorkbook();
  const blob = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  const filename = `validation_final_report_${timestampStamp()}.xlsx`;
  downloadBlob(new Blob([blob], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), filename);
  persistState();
}

function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const chunk = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode.apply(null, chunk);
  }
  return btoa(binary);
}

function wrapBase64(base64) {
  return base64.match(/.{1,76}/g)?.join('\r\n') || base64;
}

function buildEmailBody(stats, attachmentFilename) {
  return [
    'Hello,',
    '',
    'Please find attached the final validation report.',
    '',
    'Current visible summary:',
    `- Visible Tickets: ${stats.total}`,
    `- Matched: ${stats.matched}`,
    `- Corrected: ${stats.corrected}`,
    `- Open Errors: ${stats.errors}`,
    `- Exceptions Applied: ${stats.exceptions}`,
    `- Completion: ${stats.completion}%`,
    '',
    'Applied filters:',
    `- Specific Date: ${specificDateFilter.value || 'All'}`,
    `- From Date: ${fromDateFilter.value || 'All'}`,
    `- To Date: ${toDateFilter.value || 'All'}`,
    `- Status Filter: ${statusFilter.value || 'all'}`,
    `- Ticket Type: ${ticketTypeFilter.value || 'all'}`,
    '',
    'Attachment:',
    `- ${attachmentFilename}`,
    '',
    'Regards,',
    ''
  ].join('\r\n');
}

function exportOutlookEmailDraft() {
  if (!state.raw.length || typeof XLSX === 'undefined') return;

  const workbook = buildWorkbook();
  const workbookBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  const attachmentFilename = `validation_final_report_${timestampStamp()}.xlsx`;
  const base64Attachment = wrapBase64(arrayBufferToBase64(workbookBuffer));
  const stats = computeStats(state.filtered);
  const boundary = `----=_NextPart_${Date.now()}`;
  const subject = normalize(el('emailSubject').value) || defaultEmailSubject();
  const to = normalize(el('emailTo').value);
  const cc = normalize(el('emailCc').value);
  const body = buildEmailBody(stats, attachmentFilename);

  const lines = [
    `To: ${to}`,
    cc ? `Cc: ${cc}` : '',
    `Subject: ${subject}`,
    'MIME-Version: 1.0',
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    '',
    `--${boundary}`,
    'Content-Type: text/plain; charset="utf-8"',
    'Content-Transfer-Encoding: 8bit',
    '',
    body,
    '',
    `--${boundary}`,
    `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name="${attachmentFilename}"`,
    `Content-Disposition: attachment; filename="${attachmentFilename}"`,
    'Content-Transfer-Encoding: base64',
    '',
    base64Attachment,
    '',
    `--${boundary}--`,
    ''
  ].filter(Boolean);

  const emlBlob = new Blob([lines.join('\r\n')], { type: 'message/rfc822' });
  const emlFilename = `validation_report_email_${timestampStamp()}.eml`;
  downloadBlob(emlBlob, emlFilename);
  persistState();
}

function bindEvents() {
  if (typeof XLSX === 'undefined') {
    showBanner('The Excel library did not load. Check your network connection, then refresh the page.');
  }

  dropzone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', event => {
    const file = event.target.files?.[0];
    if (file) handleWorkbook(file);
  });
  dropzone.addEventListener('dragover', event => {
    event.preventDefault();
    dropzone.classList.add('drag');
  });
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('drag'));
  dropzone.addEventListener('drop', event => {
    event.preventDefault();
    dropzone.classList.remove('drag');
    const file = event.dataTransfer.files?.[0];
    if (file) handleWorkbook(file);
  });

  el('applyBtn').addEventListener('click', applyFilters);
  el('clearFiltersBtn').addEventListener('click', clearFilters);
  el('exportWorkbookBtn').addEventListener('click', exportWorkbook);
  el('exportEmailBtn').addEventListener('click', exportOutlookEmailDraft);

  [searchInput, ticketTypeFilter, specificDateFilter, fromDateFilter, toDateFilter, statusFilter].forEach(control => {
    control.addEventListener('input', applyFilters);
    control.addEventListener('change', applyFilters);
  });

  el('emailTo').addEventListener('change', persistState);
  el('emailCc').addEventListener('change', persistState);
  el('emailSubject').addEventListener('change', persistState);

  resultsBody.addEventListener('click', event => {
    const button = event.target.closest('[data-correct-row]');
    if (!button) return;
    markCorrected(button.getAttribute('data-correct-row'));
  });
}

window.addEventListener('error', event => {
  console.error(event.error || event.message);
  showBanner(`Runtime error: ${event.message || 'Unknown error'}`);
});

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', bindEvents);
} else {
  bindEvents();
}
