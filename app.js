const CONTROL_STATUSES = ['на устранении', 'на контроле инспектора оати'];
const RESOLVED_STATUS = 'снят с контроля';

const MOSCOW_DISTRICT_ABBREVIATIONS = Object.freeze({
  'центральный административный округ': 'ЦАО',
  'северный административный округ': 'САО',
  'северо-восточный административный округ': 'СВАО',
  'восточный административный округ': 'ВАО',
  'юго-восточный административный округ': 'ЮВАО',
  'южный административный округ': 'ЮАО',
  'юго-западный административный округ': 'ЮЗАО',
  'западный административный округ': 'ЗАО',
  'северо-западный административный округ': 'СЗАО',
  'зеленоградский административный округ': 'ЗелАО',
  'новомосковский административный округ': 'НАО',
  'троицкий административный округ': 'ТАО',
  'троицкий и новомосковский административный округ': 'ТиНАО',
});

const MOSCOW_DISTRICT_ABBREVIATION_KEYS = new Set(
  Object.values(MOSCOW_DISTRICT_ABBREVIATIONS).map((value) => value.trim().toLowerCase()),
);

const DISTRICT_FALLBACK_KEY = 'без округа';

const violationFieldDefinitions = [
  { key: 'id', label: 'Идентификатор нарушения', candidates: ['идентификатор', 'id', 'uid'] },
  { key: 'status', label: 'Статус нарушения', candidates: ['статус нарушения', 'статус'] },
  { key: 'objectType', label: 'Тип объекта', candidates: ['тип объекта', 'тип объекта контроля'] },
  { key: 'objectName', label: 'Наименование объекта', candidates: ['наименование объекта', 'наименование объекта контроля'] },
  { key: 'inspectionDate', label: 'Дата обследования', candidates: ['дата обследования', 'дата осмотра', 'дата контроля'] },
  { key: 'district', label: 'Округ', candidates: ['округ', 'административный округ', 'округ объекта'] },
];

const objectFieldDefinitions = [
  { key: 'objectType', label: 'Вид объекта', candidates: ['вид объекта', 'тип объекта', 'тип объекта контроля'] },
  { key: 'objectName', label: 'Наименование объекта', candidates: ['наименование объекта', 'наименование объекта контроля'] },
  { key: 'district', label: 'Округ', candidates: ['округ', 'административный округ', 'округ объекта'] },
];

const state = {
  violations: [],
  objects: [],
  violationColumns: [],
  objectColumns: [],
  violationMapping: {},
  objectMapping: {},
  typeMode: 'all',
  customTypes: [],
  availableTypes: [],
  availableDates: [],
};

const elements = {
  violationsInput: document.getElementById('violations-file'),
  objectsInput: document.getElementById('objects-file'),
  mappingSection: document.getElementById('mapping-section'),
  violationsMapping: document.getElementById('violations-mapping'),
  objectsMapping: document.getElementById('objects-mapping'),
  controlsSection: document.getElementById('controls-section'),
  previewSection: document.getElementById('preview-section'),
  currentStart: document.getElementById('current-start'),
  currentEnd: document.getElementById('current-end'),
  previousStart: document.getElementById('previous-start'),
  previousEnd: document.getElementById('previous-end'),
  typeOptions: document.getElementById('type-options'),
  typeMessage: document.getElementById('type-message'),
  previewMessage: document.getElementById('preview-message'),
  reportTable: document.getElementById('report-table'),
  refreshButton: document.getElementById('refresh-report'),
};

const numberFormatter = new Intl.NumberFormat('ru-RU');

// Обработчики файлов
if (elements.violationsInput) {
  elements.violationsInput.addEventListener('change', async (event) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    await loadDataset('violations', file);
  });
}

if (elements.objectsInput) {
  elements.objectsInput.addEventListener('change', async (event) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    await loadDataset('objects', file);
  });
}

if (elements.refreshButton) {
  elements.refreshButton.addEventListener('click', () => {
    updatePreview();
  });
}

const typeModeRadios = Array.from(document.querySelectorAll('input[name="type-mode"]'));
for (const radio of typeModeRadios) {
  radio.addEventListener('change', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }
    state.typeMode = target.value;
    if (state.typeMode === 'all') {
      state.customTypes = [];
    }
    renderTypeOptions();
    updatePreview();
  });
}

for (const select of [elements.currentStart, elements.currentEnd, elements.previousStart, elements.previousEnd]) {
  if (!select) {
    continue;
  }
  select.addEventListener('change', () => {
    updatePreview();
  });
}

async function loadDataset(kind, file) {
  try {
    showPreviewMessage(`Загрузка файла «${file.name}»...`);
    const headerCandidates = kind === 'violations'
      ? violationFieldDefinitions.flatMap((item) => item.candidates)
      : objectFieldDefinitions.flatMap((item) => item.candidates);
    const { records, headers } = await readExcelFile(file, headerCandidates);
    if (!records.length) {
      throw new Error('Не удалось обнаружить данные в выбранном файле.');
    }
    if (kind === 'violations') {
      state.violations = records;
      state.violationColumns = headers;
      state.violationMapping = autoMapColumns(headers, violationFieldDefinitions);
    } else {
      state.objects = records;
      state.objectColumns = headers;
      state.objectMapping = autoMapColumns(headers, objectFieldDefinitions);
    }
    updateVisibility();
    renderMappings();
    updateAvailableFilters();
    updatePreview();
    showPreviewMessage('Файлы загружены, параметры можно настраивать.');
  } catch (error) {
    console.error(error);
    showPreviewMessage(error instanceof Error ? error.message : 'Не удалось обработать файл.');
  }
}

async function readExcelFile(file, headerKeywords) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error('Не удалось найти лист в книге Excel.');
  }
  const sheet = workbook.Sheets[sheetName];
  if (!sheet || !sheet['!ref']) {
    throw new Error('Лист Excel не содержит данных.');
  }
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const headerRowIndex = detectHeaderRowIndex(sheet, range, headerKeywords);
  if (headerRowIndex === null) {
    throw new Error('Не удалось определить строку заголовков.');
  }
  const headers = extractHeaderRow(sheet, headerRowIndex, range);
  const records = extractRecords(sheet, headerRowIndex, headers);
  return { records, headers };
}

function detectHeaderRowIndex(sheet, range, headerKeywords) {
  const normalizedKeywords = headerKeywords.map((item) => normalizeHeaderValue(item));
  let bestRowIndex = null;
  let bestScore = -Infinity;
  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    let matchCount = 0;
    let textCount = 0;
    let valueCount = 0;
    for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      const cell = sheet[address];
      if (!cell) {
        continue;
      }
      const normalized = normalizeHeaderValue(typeof cell.v === 'string' ? cell.v : cell.w ?? cell.v);
      if (!normalized) {
        continue;
      }
      valueCount += 1;
      if (!Number.isNaN(Number(normalized))) {
        continue;
      }
      textCount += 1;
      if (normalizedKeywords.includes(normalized)) {
        matchCount += 1;
      }
    }
    if (!valueCount) {
      continue;
    }
    const score = matchCount * 10 + textCount;
    if (score > bestScore) {
      bestScore = score;
      bestRowIndex = rowIndex;
    }
  }
  return bestRowIndex;
}

function normalizeHeaderValue(value) {
  if (typeof value === 'string') {
    return value.replace(/\s+/g, ' ').trim().toLowerCase();
  }
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim().toLowerCase();
}

function extractHeaderRow(sheet, rowIndex, range) {
  const headers = [];
  for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
    const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
    const cell = sheet[address];
    headers.push(typeof cell?.v === 'string' ? cell.v.trim() : cell?.v ?? `Колонка ${colIndex + 1}`);
  }
  return headers;
}

function extractRecords(sheet, headerRowIndex, headers) {
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const records = [];
  for (let rowIndex = headerRowIndex + 1; rowIndex <= range.e.r; rowIndex += 1) {
    const record = {};
    let hasValue = false;
    for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
      const header = headers[colIndex - range.s.c];
      if (!header || header === true || header === false) {
        continue;
      }
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      const cell = sheet[address];
      const value = extractCellValue(cell);
      if (value !== '' && value !== null && value !== undefined) {
        hasValue = true;
      }
      record[header] = value;
    }
    if (hasValue) {
      records.push(record);
    }
  }
  return records;
}

function extractCellValue(cell) {
  if (!cell) {
    return '';
  }
  if (cell.t === 's' || cell.t === 'str') {
    return typeof cell.v === 'string' ? cell.v.trim() : String(cell.v ?? '');
  }
  if (cell.t === 'n' || cell.t === 'd') {
    return cell.v;
  }
  if (cell.t === 'b') {
    return Boolean(cell.v);
  }
  return cell.v ?? '';
}

function autoMapColumns(columns, definitions) {
  const mapping = {};
  for (const definition of definitions) {
    const normalizedCandidates = definition.candidates.map((candidate) => candidate.toLowerCase());
    let matchedColumn = '';
    for (const column of columns) {
      if (!column) {
        continue;
      }
      const normalized = column.toString().trim().toLowerCase();
      if (normalizedCandidates.includes(normalized)) {
        matchedColumn = column;
        break;
      }
    }
    if (!matchedColumn && columns.length) {
      matchedColumn = columns[0];
    }
    mapping[definition.key] = matchedColumn;
  }
  return mapping;
}

function renderMappings() {
  elements.mappingSection.hidden = !(state.violations.length || state.objects.length);
  elements.violationsMapping.innerHTML = '';
  elements.objectsMapping.innerHTML = '';
  if (state.violations.length) {
    renderMappingList(elements.violationsMapping, state.violationColumns, violationFieldDefinitions, state.violationMapping, (key, value) => {
      state.violationMapping[key] = value;
      updateAvailableFilters();
      updatePreview();
    });
  }
  if (state.objects.length) {
    renderMappingList(elements.objectsMapping, state.objectColumns, objectFieldDefinitions, state.objectMapping, (key, value) => {
      state.objectMapping[key] = value;
      updateAvailableFilters();
      updatePreview();
    });
  }
}

function renderMappingList(container, columns, definitions, mapping, onChange) {
  for (const definition of definitions) {
    const wrapper = document.createElement('div');
    wrapper.className = 'mapping-item';

    const label = document.createElement('label');
    label.textContent = definition.label;

    const select = document.createElement('select');
    for (const column of columns) {
      const option = document.createElement('option');
      option.value = column ?? '';
      option.textContent = column ?? '—';
      if ((column ?? '') === (mapping[definition.key] ?? '')) {
        option.selected = true;
      }
      select.append(option);
    }
    select.addEventListener('change', () => {
      onChange(definition.key, select.value);
    });

    wrapper.append(label, select);
    container.append(wrapper);
  }
}

function updateVisibility() {
  const ready = state.violations.length && state.objects.length;
  elements.mappingSection.hidden = !(state.violations.length || state.objects.length);
  elements.controlsSection.hidden = !ready;
  elements.previewSection.hidden = !ready;
  if (!ready) {
    clearTable();
  }
}

function updateAvailableFilters() {
  if (!(state.violations.length && state.objects.length)) {
    return;
  }
  const dateColumn = state.violationMapping.inspectionDate;
  if (dateColumn) {
    state.availableDates = extractUniqueDates(state.violations, dateColumn);
    renderDateSelect(elements.currentStart, state.availableDates, 0);
    renderDateSelect(elements.currentEnd, state.availableDates, state.availableDates.length - 1);
    const previousEndIndex = state.availableDates.length > 1 ? state.availableDates.length - 2 : state.availableDates.length - 1;
    renderDateSelect(elements.previousStart, state.availableDates, 0);
    renderDateSelect(elements.previousEnd, state.availableDates, previousEndIndex);
  }
  const objectTypeColumn = state.objectMapping.objectType;
  const objectTypes = collectUniqueValues(state.objects, objectTypeColumn).sort((a, b) => a.localeCompare(b, 'ru'));
  state.availableTypes = objectTypes;
  state.customTypes = state.customTypes.filter((value) => state.availableTypes.includes(value));
  renderTypeOptions();
}

function renderDateSelect(select, dates, defaultIndex) {
  if (!select) {
    return;
  }
  select.innerHTML = '';
  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = '— Выберите дату —';
  select.append(placeholder);
  for (const item of dates) {
    const option = document.createElement('option');
    option.value = item.iso;
    option.textContent = formatDateDisplay(item.date);
    select.append(option);
  }
  if (dates.length) {
    const index = Math.min(Math.max(defaultIndex, 0), dates.length - 1);
    if (index >= 0) {
      select.value = dates[index].iso;
    }
  }
}

function renderTypeOptions() {
  elements.typeOptions.innerHTML = '';
  elements.typeMessage.textContent = '';
  if (!state.availableTypes.length) {
    elements.typeOptions.textContent = 'Типы объектов не найдены.';
    return;
  }
  const selectedSet = new Set(state.customTypes);
  for (const type of state.availableTypes) {
    const label = document.createElement('label');
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.value = type;
    input.checked = selectedSet.has(type);
    input.disabled = state.typeMode !== 'custom';
    input.addEventListener('change', () => handleTypeSelection(input));
    const text = document.createElement('span');
    text.textContent = type;
    label.append(input, text);
    elements.typeOptions.append(label);
  }
  if (state.typeMode !== 'custom') {
    const checkboxes = elements.typeOptions.querySelectorAll('input[type="checkbox"]');
    checkboxes.forEach((checkbox) => {
      checkbox.checked = false;
    });
  }
}

function handleTypeSelection(checkbox) {
  if (!(checkbox instanceof HTMLInputElement)) {
    return;
  }
  const type = checkbox.value;
  if (checkbox.checked) {
    if (state.customTypes.length >= 3) {
      checkbox.checked = false;
      elements.typeMessage.textContent = 'Можно выбрать не более трех типов объектов.';
      return;
    }
    state.customTypes.push(type);
  } else {
    state.customTypes = state.customTypes.filter((value) => value !== type);
  }
  elements.typeMessage.textContent = '';
  updatePreview();
}

function updatePreview() {
  if (!(state.violations.length && state.objects.length)) {
    showPreviewMessage('Загрузите оба файла, чтобы получить отчет.');
    clearTable();
    return;
  }
  if (state.typeMode === 'custom' && !state.customTypes.length) {
    showPreviewMessage('Выберите типы объектов для пользовательского режима.');
    clearTable();
    return;
  }
  if (!isMappingComplete(state.violationMapping, violationFieldDefinitions) || !isMappingComplete(state.objectMapping, objectFieldDefinitions)) {
    showPreviewMessage('Проверьте настройки соответствия столбцов.');
    clearTable();
    return;
  }
  const periods = getSelectedPeriods();
  if (!periods) {
    clearTable();
    return;
  }
  const report = buildReport(periods);
  if (!report.rows.length) {
    showPreviewMessage('По заданным условиям данные не найдены. Измените фильтры.');
    clearTable();
    return;
  }
  showPreviewMessage('');
  renderReportTable(report, periods);
}

function isMappingComplete(mapping, definitions) {
  return definitions.every((definition) => Boolean(mapping[definition.key]));
}

function getSelectedPeriods() {
  const currentStartIso = elements.currentStart?.value ?? '';
  const currentEndIso = elements.currentEnd?.value ?? '';
  const previousStartIso = elements.previousStart?.value ?? '';
  const previousEndIso = elements.previousEnd?.value ?? '';
  if (!currentStartIso || !currentEndIso) {
    showPreviewMessage('Выберите даты начала и окончания отчётного периода.');
    return null;
  }
  if (!previousStartIso || !previousEndIso) {
    showPreviewMessage('Выберите даты предыдущего периода.');
    return null;
  }
  if (currentStartIso > currentEndIso) {
    showPreviewMessage('Дата начала отчётного периода не может быть позже даты окончания.');
    return null;
  }
  if (previousStartIso > previousEndIso) {
    showPreviewMessage('Дата начала предыдущего периода не может быть позже даты окончания.');
    return null;
  }
  return {
    current: {
      start: parseIsoDate(currentStartIso),
      end: parseIsoDate(currentEndIso),
    },
    previous: {
      start: parseIsoDate(previousStartIso),
      end: parseIsoDate(previousEndIso),
    },
  };
}

function getDistrictDisplayLabel(label) {
  if (!label) {
    return 'Без округа';
  }
  const normalized = normalizeDistrictKey(label);
  if (!normalized || normalized === 'без округа') {
    return 'Без округа';
  }
  const abbreviation = MOSCOW_DISTRICT_ABBREVIATIONS[normalized];
  if (abbreviation) {
    return abbreviation;
  }
  const trimmed = label.trim();
  // Пример: getDistrictDisplayLabel('Центральный административный округ') вернёт 'ЦАО'.
  return trimmed || 'Без округа';
}

function normalizeDistrictKey(value) {
  const normalized = normalizeKey(value);
  if (!normalized) {
    return '';
  }
  const withoutParentheses = normalized.replace(/\s*\(.*?\)\s*/g, ' ');
  const withoutCityMention = withoutParentheses.replace(/\s+г\.?\s*москв[аеы]?$/u, '');
  return withoutCityMention.replace(/\s+/g, ' ').trim();
}

function buildDistrictLookup(objectRecords, objectDistrictColumn, violationRecords, violationDistrictColumn) {
  const lookup = new Map();

  const registerValue = (rawValue) => {
    const label = getValueAsString(rawValue);
    const normalizedKey = normalizeDistrictKey(label);
    const key = normalizedKey || DISTRICT_FALLBACK_KEY;
    if (!key) {
      return;
    }
    const displayLabel = normalizedKey ? getDistrictDisplayLabel(label) : 'Без округа';
    if (!lookup.has(key)) {
      lookup.set(key, displayLabel);
      return;
    }
    const current = lookup.get(key);
    if (shouldReplaceDistrictLabel(current, displayLabel)) {
      lookup.set(key, displayLabel);
    }
  };

  const registerFromRecords = (records, column) => {
    if (!column) {
      return;
    }
    for (const record of records) {
      registerValue(record[column]);
    }
  };

  registerFromRecords(objectRecords, objectDistrictColumn);
  registerFromRecords(violationRecords, violationDistrictColumn);

  return lookup;
}

function shouldReplaceDistrictLabel(current, candidate) {
  if (!candidate || candidate === current) {
    return false;
  }
  if (!current) {
    return true;
  }
  const currentKey = normalizeDistrictKey(current);
  const candidateKey = normalizeDistrictKey(candidate);
  const currentIsAbbreviation = MOSCOW_DISTRICT_ABBREVIATION_KEYS.has(currentKey);
  const candidateIsAbbreviation = MOSCOW_DISTRICT_ABBREVIATION_KEYS.has(candidateKey);
  if (candidateIsAbbreviation && !currentIsAbbreviation) {
    return true;
  }
  if (!candidateIsAbbreviation && currentIsAbbreviation) {
    return false;
  }
  return candidate.length < current.length;
}

// Пример использования: buildDistrictLookup([{ district: 'Центральный административный округ' }], 'district', [], 'district');

function buildReport(periods) {
  const violationMapping = state.violationMapping;
  const objectMapping = state.objectMapping;
  const typePredicate = createTypePredicate();
  const districtData = new Map();
  const districtLookup = buildDistrictLookup(
    state.objects,
    objectMapping.district,
    state.violations,
    violationMapping.district,
  );
  const ensureEntry = (label) => {
    const normalizedKey = normalizeDistrictKey(label);
    const key = normalizedKey || DISTRICT_FALLBACK_KEY;
    const displayLabel = districtLookup.get(key) ?? (normalizedKey ? getDistrictDisplayLabel(label) : 'Без округа');
    if (!districtData.has(key)) {
      districtData.set(key, {
        label: displayLabel,
        totalObjects: new Set(),
        inspectedObjects: new Set(),
        objectsWithViolations: new Set(),
        currentViolationIds: new Set(),
        previousControlIds: new Set(),
        resolvedIds: new Set(),
        controlIds: new Set(),
      });
    }
    return districtData.get(key);
  };

  for (const record of state.objects) {
    const typeValue = getValueAsString(record[objectMapping.objectType]);
    if (typeValue && !typePredicate(typeValue)) {
      continue;
    }
    if (!typeValue && state.typeMode === 'custom') {
      continue;
    }
    const districtLabel = getValueAsString(record[objectMapping.district]);
    const objectName = getValueAsString(record[objectMapping.objectName]);
    if (!objectName) {
      continue;
    }
    const entry = ensureEntry(districtLabel);
    entry.totalObjects.add(objectName);
  }

  for (const record of state.violations) {
    const typeValue = getValueAsString(record[violationMapping.objectType]);
    if (typeValue && !typePredicate(typeValue)) {
      continue;
    }
    if (!typeValue && state.typeMode === 'custom') {
      continue;
    }
    const districtLabel = getValueAsString(record[violationMapping.district]);
    const objectName = getValueAsString(record[violationMapping.objectName]);
    const violationId = getValueAsString(record[violationMapping.id]);
    const status = getValueAsString(record[violationMapping.status]);
    const dateValue = record[violationMapping.inspectionDate];
    const inspectionDate = parseDateValue(dateValue);
    if (!inspectionDate) {
      continue;
    }
    const entry = ensureEntry(districtLabel);
    if (isWithinPeriod(inspectionDate, periods.current)) {
      if (objectName) {
        entry.inspectedObjects.add(objectName);
      }
      if (violationId) {
        if (objectName) {
          entry.objectsWithViolations.add(objectName);
        }
        entry.currentViolationIds.add(violationId);
      }
      if (status && normalizeText(status) === RESOLVED_STATUS && violationId) {
        entry.resolvedIds.add(violationId);
      }
      if (status && CONTROL_STATUSES.includes(normalizeText(status)) && violationId) {
        entry.controlIds.add(violationId);
      }
    }
    if (isWithinPeriod(inspectionDate, periods.previous)) {
      if (status && CONTROL_STATUSES.includes(normalizeText(status))) {
        if (violationId) {
          entry.previousControlIds.add(violationId);
        }
      }
    }
  }

  const rows = [];
  const totals = {
    totalObjects: 0,
    inspectedObjects: 0,
    objectsWithViolations: 0,
    totalViolations: new Set(),
    currentViolations: new Set(),
    previousControl: new Set(),
    resolved: new Set(),
    onControl: new Set(),
  };

  const sortedEntries = Array.from(districtData.values()).sort((a, b) => a.label.localeCompare(b.label, 'ru'));
  for (const entry of sortedEntries) {
    const totalObjectsCount = entry.totalObjects.size;
    const inspectedCount = entry.inspectedObjects.size;
    const objectWithViolationsCount = entry.objectsWithViolations.size;
    const currentViolationsCount = entry.currentViolationIds.size;
    const previousControlCount = entry.previousControlIds.size;
    const resolvedCount = entry.resolvedIds.size;
    const onControlCount = entry.controlIds.size;
    const totalViolationsCount = new Set([...entry.currentViolationIds, ...entry.previousControlIds]).size;

    rows.push({
      label: entry.label,
      totalObjects: totalObjectsCount,
      inspectedObjects: inspectedCount,
      inspectedPercent: computePercent(inspectedCount, totalObjectsCount),
      violationPercent: computePercent(objectWithViolationsCount, totalObjectsCount),
      totalViolations: totalViolationsCount,
      currentViolations: currentViolationsCount,
      previousControl: previousControlCount,
      resolved: resolvedCount,
      onControl: onControlCount,
    });

    totals.totalObjects += totalObjectsCount;
    totals.inspectedObjects += inspectedCount;
    totals.objectsWithViolations += objectWithViolationsCount;
    for (const id of entry.currentViolationIds) {
      totals.currentViolations.add(id);
    }
    for (const id of entry.previousControlIds) {
      totals.previousControl.add(id);
    }
    for (const id of entry.resolvedIds) {
      totals.resolved.add(id);
    }
    for (const id of entry.controlIds) {
      totals.onControl.add(id);
    }
    for (const id of entry.currentViolationIds) {
      totals.totalViolations.add(id);
    }
    for (const id of entry.previousControlIds) {
      totals.totalViolations.add(id);
    }
  }

  const totalRow = {
    label: 'Итого',
    totalObjects: totals.totalObjects,
    inspectedObjects: totals.inspectedObjects,
    inspectedPercent: computePercent(totals.inspectedObjects, totals.totalObjects),
    violationPercent: computePercent(totals.objectsWithViolations, totals.totalObjects),
    totalViolations: totals.totalViolations.size,
    currentViolations: totals.currentViolations.size,
    previousControl: totals.previousControl.size,
    resolved: totals.resolved.size,
    onControl: totals.onControl.size,
  };

  return { rows, totalRow };
}

function renderReportTable(report, periods) {
  const headers = buildTableHeaders(periods);
  const table = elements.reportTable;
  table.innerHTML = '';

  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');
  for (const headerText of headers) {
    const th = document.createElement('th');
    th.textContent = headerText;
    headerRow.append(th);
  }
  thead.append(headerRow);
  table.append(thead);

  const tbody = document.createElement('tbody');
  for (const row of report.rows) {
    tbody.append(createRowElement(row));
  }
  table.append(tbody);

  const tfoot = document.createElement('tfoot');
  tfoot.append(createRowElement(report.totalRow));
  table.append(tfoot);
}

function createRowElement(row) {
  const tr = document.createElement('tr');
  tr.append(createCell(row.label, true));
  tr.append(createCell(formatInteger(row.totalObjects)));
  tr.append(createCell(formatInteger(row.inspectedObjects)));
  tr.append(createCell(formatPercent(row.inspectedPercent)));
  tr.append(createCell(formatPercent(row.violationPercent)));
  tr.append(createCell(formatInteger(row.totalViolations)));
  tr.append(createCell(formatInteger(row.currentViolations)));
  tr.append(createCell(formatInteger(row.previousControl)));
  tr.append(createCell(formatInteger(row.resolved)));
  tr.append(createCell(formatInteger(row.onControl)));
  return tr;
}

function createCell(value, isHeader = false) {
  const cell = document.createElement(isHeader ? 'th' : 'td');
  cell.textContent = value ?? '';
  return cell;
}

function buildTableHeaders(periods) {
  const totalHeader = buildTotalHeader();
  const rangeHeader = `Проверено объектов с ${formatDateDisplay(periods.current.start)} по ${formatDateDisplay(periods.current.end)}`;
  return [
    'Округ',
    totalHeader,
    rangeHeader,
    '% проверенных объектов от общего количества',
    '% объектов с выявленными нарушениями',
    'Всего выявленных нарушений',
    'Нарушения, выявленные за отчётный период',
    'Нарушения, находящиеся на контроле, с предыдущего периода',
    'Устранено нарушений',
    'Нарушения на контроле',
  ];
}

function buildTotalHeader() {
  if (state.typeMode === 'all' || !state.customTypes.length) {
    return 'Всего объектов';
  }
  if (state.customTypes.length === 1) {
    return `Всего ${state.customTypes[0]}`;
  }
  return `Всего (${state.customTypes.join(', ')})`;
}

function clearTable() {
  elements.reportTable.innerHTML = '';
}

function showPreviewMessage(message) {
  elements.previewMessage.textContent = message;
}

function parseIsoDate(iso) {
  const [year, month, day] = iso.split('-').map((part) => Number.parseInt(part, 10));
  return new Date(year, month - 1, day);
}

function createTypePredicate() {
  if (state.typeMode === 'all' || !state.customTypes.length) {
    return () => true;
  }
  const allowed = new Set(state.customTypes.map((value) => normalizeKey(value)));
  return (value) => allowed.has(normalizeKey(value));
}

function normalizeKey(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return value.toString().trim().toLowerCase();
}

function normalizeText(value) {
  return normalizeKey(value);
}

function getValueAsString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'string') {
    return value.trim();
  }
  if (value instanceof Date) {
    return formatDateDisplay(value);
  }
  return value.toString().trim();
}

function extractUniqueDates(records, column) {
  const dates = [];
  const seen = new Set();
  for (const record of records) {
    const value = record[column];
    const date = parseDateValue(value);
    if (!date) {
      continue;
    }
    const iso = formatIsoDate(date);
    if (!seen.has(iso)) {
      seen.add(iso);
      dates.push({ iso, date });
    }
  }
  dates.sort((a, b) => a.iso.localeCompare(b.iso));
  return dates;
}

function parseDateValue(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  if (value instanceof Date) {
    return new Date(value.getTime());
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    const parsed = XLSX.SSF?.parse_date_code?.(value);
    if (!parsed) {
      return null;
    }
    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return null;
    }
    const dateTimeMatch = trimmed.match(/^(\d{1,2})[.](\d{1,2})[.](\d{4})(?:\s+(\d{1,2}):(\d{2}))?$/);
    if (dateTimeMatch) {
      const [, day, month, year] = dateTimeMatch;
      return new Date(Number(year), Number(month) - 1, Number(day));
    }
    const isoDate = Date.parse(trimmed);
    if (Number.isFinite(isoDate)) {
      return new Date(isoDate);
    }
  }
  return null;
}

function formatIsoDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatDateDisplay(date) {
  if (!(date instanceof Date)) {
    return '';
  }
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
}

function collectUniqueValues(records, column) {
  if (!column) {
    return [];
  }
  const seen = new Map();
  for (const record of records) {
    const value = record[column];
    const text = getValueAsString(value);
    const key = normalizeKey(text);
    if (!key) {
      continue;
    }
    if (!seen.has(key)) {
      seen.set(key, text);
    }
  }
  return Array.from(seen.values());
}

function isWithinPeriod(date, period) {
  if (!period) {
    return false;
  }
  const dateOnly = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  return dateOnly >= period.start && dateOnly <= period.end;
}

function computePercent(part, total) {
  if (!total || !Number.isFinite(part)) {
    return 0;
  }
  return (part / total) * 100;
}

function formatInteger(value) {
  if (!Number.isFinite(value)) {
    return '0';
  }
  return numberFormatter.format(Math.round(value));
}

function formatPercent(value) {
  if (!Number.isFinite(value)) {
    return '0';
  }
  return value.toFixed(1);
}
