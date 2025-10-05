// Ниже я собираю все служебные справочники, чтобы дальше по коду не гоняться за строками руками.
// Такой подход сразу же подсвечивает, какие статусы и сокращения мы ожидаем в входных данных.
const CONTROL_STATUSES = ['на устранении', 'на контроле инспектора оати'];
const CONTROL_STATUS_SET = new Set(CONTROL_STATUSES);
const RESOLVED_STATUS = 'снят с контроля';
const INSPECTION_RESULT_VIOLATION = 'нарушение выявлено';

// Ограничения по работе с файлами: защищают приложение от чрезмерных ресурсов и потенциальных атак.
const MAX_UPLOAD_SIZE_BYTES = 10 * 1024 * 1024; // 10 МиБ.
const MAX_ROWS_PER_SHEET = 20000;
const MAX_COLUMNS_PER_SHEET = 120;
const ALLOWED_FILE_EXTENSIONS = new Set(['xlsx', 'xls', 'xlsm']);

// Универсальный форматтер чисел: используется в интерфейсе и в сообщениях об ошибках.
const numberFormatter = new Intl.NumberFormat('ru-RU');

// Список округов и их официальных сокращений держу в одном месте, чтобы дальше не размазывать магические строки.
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

// Заодно храню список уже нормализованных сокращений, чтобы быстро понимать что в label у нас аббревиатура.
const MOSCOW_DISTRICT_ABBREVIATION_KEYS = new Set(
  Object.values(MOSCOW_DISTRICT_ABBREVIATIONS).map((value) => value.trim().toLowerCase()),
);

// На случай, когда в данных вообще нет округа, завожу понятный ключ.
const DISTRICT_FALLBACK_KEY = 'без округа';

// Здесь расписал какие поля мы ожидаем получить из таблицы нарушений.
const violationFieldDefinitions = [
  { key: 'id', label: 'Идентификатор нарушения', candidates: ['идентификатор', 'id', 'uid'] },
  { key: 'status', label: 'Статус нарушения', candidates: ['статус нарушения', 'статус'] },
  { key: 'violationName', label: 'Наименование нарушения', candidates: ['наименование нарушения', 'нарушение'] },
  {
    key: 'inspectionResult',
    label: 'Результат обследования',
    candidates: ['результат обследования', 'результат осмотра', 'результат проверки'],
  },
  { key: 'objectType', label: 'Тип объекта', candidates: ['тип объекта', 'тип объекта контроля'] },
  { key: 'objectName', label: 'Наименование объекта', candidates: ['наименование объекта', 'наименование объекта контроля'] },
  { key: 'inspectionDate', label: 'Дата обследования', candidates: ['дата обследования', 'дата осмотра', 'дата контроля'] },
  { key: 'district', label: 'Округ', candidates: ['округ', 'административный округ', 'округ объекта'] },
  { key: 'dataSource', label: 'Источник данных', candidates: ['источник данных'], optional: true },
];

// Аналогичный список для справочника объектов.
const objectFieldDefinitions = [
  { key: 'objectType', label: 'Вид объекта', candidates: ['вид объекта', 'тип объекта', 'тип объекта контроля'] },
  { key: 'objectName', label: 'Наименование объекта', candidates: ['наименование объекта', 'наименование объекта контроля'] },
  { key: 'district', label: 'Округ', candidates: ['округ', 'административный округ', 'округ объекта'] },
];

// Централизованное состояние приложения, чтобы не расползались глобальные переменные.
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
  violationMode: 'all',
  customViolations: [],
  availableViolations: [],
  availableDates: [],
  lastReport: null,
  lastPeriods: null,
  dataSourceOption: 'all',
};

// Чтобы лишний раз не перерисовывать таблицу при серии событий, ввёл очередь на requestAnimationFrame.
let scheduledPreviewHandle = null;
let scheduledPreviewKind = null;

// Сразу собираю все элементы DOM, чтобы дальше обращаться по коротким именам.
const elements = {
  violationsInput: document.getElementById('violations-file'),
  objectsInput: document.getElementById('objects-file'),
  mappingSection: document.getElementById('mapping-section'),
  violationsMapping: document.getElementById('violations-mapping'),
  objectsMapping: document.getElementById('objects-mapping'),
  controlsSection: document.getElementById('controls-section'),
  previewSection: document.getElementById('preview-section'),
  violationsLoader: document.getElementById('violations-loader'),
  objectsLoader: document.getElementById('objects-loader'),
  currentStart: document.getElementById('current-start'),
  currentEnd: document.getElementById('current-end'),
  previousStart: document.getElementById('previous-start'),
  previousEnd: document.getElementById('previous-end'),
  typeOptions: document.getElementById('type-options'),
  typeMessage: document.getElementById('type-message'),
  violationOptions: document.getElementById('violation-options'),
  violationMessage: document.getElementById('violation-message'),
  dataSourceSelect: document.getElementById('data-source-filter'),
  dataSourceMessage: document.getElementById('data-source-message'),
  previewMessage: document.getElementById('preview-message'),
  reportTable: document.getElementById('report-table'),
  refreshButton: document.getElementById('refresh-report'),
  downloadButton: document.getElementById('download-report'),
  themeToggle: document.getElementById('theme-toggle'),
  themeToggleLabel: document.getElementById('theme-toggle-text'),
  brandLogo: document.querySelector('.brand-logo'),
  easterEggMessage: document.getElementById('easter-egg-message'),
  creditBadge: document.querySelector('.credit-badge'),
  creditBadgeClose: document.querySelector('.credit-badge__close'),
};

/**
 * Подготавливает имя файла для безопасного отображения в интерфейсе.
 * Убирает управляющие символы и чрезмерно длинные строки, чтобы предотвратить XSS и визуальный шум.
 * @param {string} rawName исходное имя файла из File API
 * @returns {string} отформатированное имя для показа пользователю
 */
function sanitizeDisplayFileName(rawName) {
  if (!rawName) {
    return 'без названия';
  }
  const normalized = rawName.toString().replace(/[\u0000-\u001f<>:"/\\|?*]+/g, '').trim();
  if (!normalized) {
    return 'без названия';
  }
  if (normalized.length > 60) {
    return `${normalized.slice(0, 57)}…`;
  }
  return normalized;
}

/**
 * Возвращает расширение файла без точки.
 * @param {string} fileName имя файла
 * @returns {string} расширение в нижнем регистре
 */
function getFileExtension(fileName) {
  if (!fileName || typeof fileName !== 'string') {
    return '';
  }
  const parts = fileName.split('.');
  return parts.length > 1 ? parts.pop().toLowerCase() : '';
}

/**
 * Преобразует размер файла в человекочитаемый формат (байты → КиБ/МиБ и т.д.).
 * @param {number} bytes исходное количество байтов
 * @returns {string} строка вида «1,5 МиБ»
 */
function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) {
    return '0 байт';
  }
  const units = ['байт', 'КиБ', 'МиБ', 'ГиБ'];
  let value = bytes;
  let unitIndex = 0;
  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }
  if (unitIndex === 0) {
    return `${numberFormatter.format(Math.round(value))} ${units[unitIndex]}`;
  }
  const isInteger = Number.isInteger(value);
  if (isInteger || value >= 10) {
    return `${numberFormatter.format(Math.round(value))} ${units[unitIndex]}`;
  }
  const rounded = Math.round(value * 10) / 10;
  return `${rounded.toString().replace('.', ',')} ${units[unitIndex]}`;
}

/**
 * Проверяет файл перед чтением: размер, расширение и пустоту.
 * Бросает Error с дружелюбным сообщением, если файл не подходит.
 * @param {File} file объект файла из input[type="file"]
 * @param {string} displayName подготовленное имя для сообщений
 */
function assertValidFile(file, displayName) {
  const safeName = displayName || sanitizeDisplayFileName(file?.name);
  if (typeof File !== 'undefined' && !(file instanceof File)) {
    throw new Error('Не удалось прочитать выбранный файл. Попробуйте выбрать его повторно.');
  }
  if (!file || file.size === 0) {
    throw new Error(`Файл «${safeName}» пуст или недоступен. Выберите другой документ.`);
  }
  if (file.size > MAX_UPLOAD_SIZE_BYTES) {
    throw new Error(
      `Файл «${safeName}» превышает допустимый размер ${formatBytes(MAX_UPLOAD_SIZE_BYTES)}. Уменьшите выгрузку или разделите данные.`,
    );
  }
  const extension = getFileExtension(file.name);
  if (extension && !ALLOWED_FILE_EXTENSIONS.has(extension)) {
    throw new Error(`Файл «${safeName}» имеет неподдерживаемое расширение. Используйте Excel в формате XLSX или XLS.`);
  }
}

/**
 * Позволяет отдать управление потоку браузера перед тяжёлой операцией, чтобы интерфейс не зависал.
 * В первую очередь используется requestIdleCallback, при его отсутствии — короткий setTimeout.
 * @returns {Promise<void>} промис, который резолвится после уступки потока
 */
async function yieldToEventLoop() {
  if (typeof window !== 'undefined') {
    if (typeof window.requestIdleCallback === 'function') {
      await new Promise((resolve) => window.requestIdleCallback(() => resolve(), { timeout: 32 }));
      return;
    }
    if (typeof window.requestAnimationFrame === 'function') {
      await new Promise((resolve) => window.requestAnimationFrame(() => resolve()));
      return;
    }
  }
  await new Promise((resolve) => setTimeout(resolve, 16));
}

/**
 * Делает имена столбцов уникальными, чтобы значения из Excel не перезаписывали друг друга.
 * @param {string[]} headers массив исходных заголовков
 * @returns {string[]} массив уникальных заголовков
 */
function makeHeadersUnique(headers) {
  const counters = new Map();
  return headers.map((header, index) => {
    const source =
      typeof header === 'string'
        ? header.trim()
        : header === null || header === undefined
        ? ''
        : header.toString().trim();
    const normalized = source || `Колонка ${index + 1}`;
    const currentCount = counters.get(normalized) ?? 0;
    counters.set(normalized, currentCount + 1);
    if (currentCount === 0) {
      return normalized;
    }
    return `${normalized} (${currentCount + 1})`;
  });
}

const MONTH_NAMES_RU = [
  'Январь',
  'Февраль',
  'Март',
  'Апрель',
  'Май',
  'Июнь',
  'Июль',
  'Август',
  'Сентябрь',
  'Октябрь',
  'Ноябрь',
  'Декабрь',
];

const WEEKDAY_LABELS_RU = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб', 'Вс'];

class AvailabilityDatePicker {
  constructor(container, options = {}) {
    this.container = container;
    this.input = container.querySelector('input[type="hidden"]');
    this.button = container.querySelector('.date-picker__button');
    this.valueNode = container.querySelector('.date-picker__value');
    this.dropdown = container.querySelector('.date-picker__dropdown');
    this.placeholder = options.placeholder ?? this.valueNode?.dataset.placeholder ?? '';
    this.onChange = typeof options.onChange === 'function' ? options.onChange : null;
    this.availableDates = [];
    this.availableSet = new Set();
    this.availableMonths = [];
    this.monthIndex = 0;
    this.isOpen = false;
    this.boundHandlePointerDown = this.handlePointerDown.bind(this);
    this.boundHandleKeyDown = this.handleKeyDown.bind(this);

    if (this.dropdown) {
      this.dropdown.tabIndex = -1;
    }

    if (this.button) {
      this.button.addEventListener('click', () => {
        if (this.isOpen) {
          this.close();
        } else {
          this.open();
        }
      });
    }

    this.updateDisplay();
  }

  getValue() {
    return this.input?.value ?? '';
  }

  setAvailableDates(dates) {
    const normalizedDates = [];
    const seenMonths = new Map();
    if (Array.isArray(dates)) {
      for (const item of dates) {
        if (!item || typeof item.iso !== 'string') {
          continue;
        }
        const iso = item.iso;
        let baseDate = item.date instanceof Date ? new Date(item.date.getTime()) : null;
        if (!(baseDate instanceof Date) || Number.isNaN(baseDate.getTime())) {
          const parsed = parseIsoDate(iso);
          if (Number.isNaN(parsed.getTime())) {
            continue;
          }
          baseDate = parsed;
        }
        normalizedDates.push({ iso, date: baseDate });
        const monthKey = `${baseDate.getFullYear()}-${baseDate.getMonth()}`;
        if (!seenMonths.has(monthKey)) {
          seenMonths.set(monthKey, { year: baseDate.getFullYear(), month: baseDate.getMonth() });
        }
      }
    }
    normalizedDates.sort((a, b) => a.iso.localeCompare(b.iso));
    this.availableDates = normalizedDates;
    this.availableSet = new Set(normalizedDates.map((item) => item.iso));
    this.availableMonths = Array.from(seenMonths.values()).sort((a, b) => {
      if (a.year !== b.year) {
        return a.year - b.year;
      }
      return a.month - b.month;
    });
    if (this.availableMonths.length === 0) {
      this.monthIndex = 0;
    } else if (this.monthIndex >= this.availableMonths.length) {
      this.monthIndex = this.availableMonths.length - 1;
    }
    this.alignMonthWithValue();
    this.renderCalendar();
    this.updateDisplay();
  }

  setValue(value, { emitEvent = false } = {}) {
    const normalized = typeof value === 'string' ? value : '';
    const nextValue = normalized && this.availableSet.size && !this.availableSet.has(normalized) ? '' : normalized;
    const previousValue = this.getValue();
    if (this.input && previousValue !== nextValue) {
      this.input.value = nextValue;
    }
    this.alignMonthWithValue();
    this.renderCalendar();
    this.updateDisplay();
    if (emitEvent && this.onChange && previousValue !== nextValue) {
      this.onChange(nextValue);
    }
  }

  updateDisplay() {
    if (!this.valueNode) {
      return;
    }
    const value = this.getValue();
    if (value) {
      const parsed = parseIsoDate(value);
      if (!Number.isNaN(parsed.getTime())) {
        this.valueNode.textContent = formatDateDisplay(parsed);
      } else {
        this.valueNode.textContent = value;
      }
      this.valueNode.classList.remove('is-placeholder');
    } else {
      this.valueNode.textContent = this.placeholder;
      this.valueNode.classList.add('is-placeholder');
    }
    if (this.button) {
      this.button.setAttribute('aria-expanded', this.isOpen ? 'true' : 'false');
    }
  }

  open() {
    if (!this.dropdown || this.isOpen) {
      return;
    }
    this.isOpen = true;
    this.renderCalendar();
    this.dropdown.hidden = false;
    this.container.classList.add('is-open');
    this.updateDisplay();
    document.addEventListener('pointerdown', this.boundHandlePointerDown, true);
    document.addEventListener('keydown', this.boundHandleKeyDown);
    const focusCallback = () => {
      this.focusInitialDay();
    };
    if (typeof window !== 'undefined' && typeof window.requestAnimationFrame === 'function') {
      window.requestAnimationFrame(focusCallback);
    } else {
      setTimeout(focusCallback, 0);
    }
  }

  close() {
    if (!this.dropdown || !this.isOpen) {
      return;
    }
    this.isOpen = false;
    this.dropdown.hidden = true;
    this.container.classList.remove('is-open');
    this.updateDisplay();
    document.removeEventListener('pointerdown', this.boundHandlePointerDown, true);
    document.removeEventListener('keydown', this.boundHandleKeyDown);
  }

  handlePointerDown(event) {
    if (!this.container.contains(event.target)) {
      this.close();
    }
  }

  handleKeyDown(event) {
    if (event.key === 'Escape') {
      this.close();
      this.button?.focus();
    }
  }

  alignMonthWithValue() {
    const value = this.getValue();
    if (!value) {
      return;
    }
    const parsed = parseIsoDate(value);
    if (Number.isNaN(parsed.getTime())) {
      return;
    }
    const key = `${parsed.getFullYear()}-${parsed.getMonth()}`;
    const index = this.availableMonths.findIndex((item) => `${item.year}-${item.month}` === key);
    if (index >= 0) {
      this.monthIndex = index;
    }
  }

  renderCalendar() {
    if (!this.dropdown) {
      return;
    }
    this.dropdown.innerHTML = '';
    if (!this.availableDates.length) {
      const empty = document.createElement('p');
      empty.className = 'date-picker__empty';
      empty.textContent = 'Нет доступных дат';
      this.dropdown.append(empty);
      return;
    }
    const monthInfo = this.availableMonths[this.monthIndex] ?? this.availableMonths[0];
    if (!monthInfo) {
      const fallback = document.createElement('p');
      fallback.className = 'date-picker__empty';
      fallback.textContent = 'Нет доступных дат';
      this.dropdown.append(fallback);
      return;
    }
    const calendar = document.createElement('div');
    calendar.className = 'date-picker__calendar';

    const header = document.createElement('div');
    header.className = 'date-picker__header';
    const prevButton = document.createElement('button');
    prevButton.type = 'button';
    prevButton.className = 'date-picker__nav-button';
    prevButton.textContent = '‹';
    prevButton.setAttribute('aria-label', 'Предыдущий месяц');
    prevButton.disabled = this.monthIndex <= 0;
    prevButton.addEventListener('click', () => {
      if (this.monthIndex > 0) {
        this.monthIndex -= 1;
        this.renderCalendar();
      }
    });

    const nextButton = document.createElement('button');
    nextButton.type = 'button';
    nextButton.className = 'date-picker__nav-button';
    nextButton.textContent = '›';
    nextButton.setAttribute('aria-label', 'Следующий месяц');
    nextButton.disabled = this.monthIndex >= this.availableMonths.length - 1;
    nextButton.addEventListener('click', () => {
      if (this.monthIndex < this.availableMonths.length - 1) {
        this.monthIndex += 1;
        this.renderCalendar();
      }
    });

    const monthLabel = document.createElement('span');
    monthLabel.className = 'date-picker__month-label';
    monthLabel.textContent = `${MONTH_NAMES_RU[monthInfo.month] ?? ''} ${monthInfo.year}`.trim();

    header.append(prevButton, monthLabel, nextButton);
    calendar.append(header);

    const weekdaysRow = document.createElement('div');
    weekdaysRow.className = 'date-picker__weekdays';
    for (const label of WEEKDAY_LABELS_RU) {
      const weekday = document.createElement('span');
      weekday.className = 'date-picker__weekday';
      weekday.textContent = label;
      weekdaysRow.append(weekday);
    }
    calendar.append(weekdaysRow);

    const grid = document.createElement('div');
    grid.className = 'date-picker__grid';
    const firstDay = new Date(monthInfo.year, monthInfo.month, 1);
    const offset = (firstDay.getDay() + 6) % 7;
    for (let i = 0; i < offset; i += 1) {
      const filler = document.createElement('span');
      filler.className = 'date-picker__day is-disabled';
      filler.setAttribute('aria-hidden', 'true');
      grid.append(filler);
    }
    const daysInMonth = new Date(monthInfo.year, monthInfo.month + 1, 0).getDate();
    const currentValue = this.getValue();
    for (let day = 1; day <= daysInMonth; day += 1) {
      const dateObject = new Date(monthInfo.year, monthInfo.month, day);
      const iso = formatIsoDate(dateObject);
      if (this.availableSet.has(iso)) {
        const dayButton = document.createElement('button');
        dayButton.type = 'button';
        dayButton.className = 'date-picker__day is-available';
        dayButton.textContent = String(day);
        dayButton.dataset.value = iso;
        dayButton.setAttribute('aria-label', formatDateDisplay(dateObject));
        if (iso === currentValue) {
          dayButton.classList.add('is-selected');
        }
        dayButton.addEventListener('click', () => {
          this.setValue(iso, { emitEvent: true });
          this.close();
          this.button?.focus();
        });
        grid.append(dayButton);
      } else {
        const dayCell = document.createElement('span');
        dayCell.className = 'date-picker__day is-disabled';
        dayCell.textContent = String(day);
        grid.append(dayCell);
      }
    }

    calendar.append(grid);
    this.dropdown.append(calendar);
  }

  focusInitialDay() {
    if (!this.dropdown) {
      return;
    }
    const target = this.dropdown.querySelector('.date-picker__day.is-selected')
      || this.dropdown.querySelector('.date-picker__day.is-available');
    if (target instanceof HTMLElement) {
      target.focus();
    } else {
      this.dropdown.focus();
    }
  }
}

const datePickers = {};

const datePickerConfigs = [
  { key: 'currentStart', selector: '[data-picker="current-start"]' },
  { key: 'currentEnd', selector: '[data-picker="current-end"]' },
  { key: 'previousStart', selector: '[data-picker="previous-start"]' },
  { key: 'previousEnd', selector: '[data-picker="previous-end"]' },
];

for (const config of datePickerConfigs) {
  const container = document.querySelector(config.selector);
  if (!container) {
    continue;
  }
  datePickers[config.key] = new AvailabilityDatePicker(container, {
    placeholder: '— Выберите дату —',
    onChange: () => {
      schedulePreviewUpdate();
    },
  });
}

const THEME_STORAGE_KEY = 'oati-theme-preference';
const themeMediaQuery = typeof window !== 'undefined' && window.matchMedia
  ? window.matchMedia('(prefers-color-scheme: dark)')
  : null;

initTheme();

function readStoredTheme() {
  try {
    const storedValue = localStorage.getItem(THEME_STORAGE_KEY);
    if (storedValue === 'light' || storedValue === 'dark') {
      return storedValue;
    }
  } catch (error) {
    console.warn('Не удалось получить сохраненную тему интерфейса:', error);
  }
  return null;
}

function persistTheme(theme) {
  try {
    localStorage.setItem(THEME_STORAGE_KEY, theme);
  } catch (error) {
    console.warn('Не удалось сохранить выбранную тему интерфейса:', error);
  }
}

function resolveInitialTheme() {
  const stored = readStoredTheme();
  if (stored) {
    return stored;
  }
  if (themeMediaQuery?.matches) {
    return 'dark';
  }
  return 'light';
}

function applyTheme(theme, { persist = false } = {}) {
  const normalized = theme === 'dark' ? 'dark' : 'light';
  document.body.classList.toggle('theme-dark', normalized === 'dark');
  document.body.classList.toggle('theme-light', normalized === 'light');
  document.body.dataset.theme = normalized;

  if (elements.themeToggle) {
    elements.themeToggle.setAttribute('aria-pressed', normalized === 'dark' ? 'true' : 'false');
  }

  if (elements.themeToggleLabel) {
    elements.themeToggleLabel.textContent = normalized === 'dark' ? 'Светлый режим' : 'Темный режим';
  }

  if (persist) {
    persistTheme(normalized);
  }
}

function initTheme() {
  const initialTheme = resolveInitialTheme();
  applyTheme(initialTheme);

  if (elements.themeToggle) {
    elements.themeToggle.addEventListener('click', () => {
      const nextTheme = document.body.classList.contains('theme-dark') ? 'light' : 'dark';
      applyTheme(nextTheme, { persist: true });
    });
  }

  if (themeMediaQuery) {
    const handleMediaChange = (event) => {
      if (readStoredTheme()) {
        return;
      }
      applyTheme(event.matches ? 'dark' : 'light');
    };

    if (typeof themeMediaQuery.addEventListener === 'function') {
      themeMediaQuery.addEventListener('change', handleMediaChange);
    } else if (typeof themeMediaQuery.addListener === 'function') {
      themeMediaQuery.addListener(handleMediaChange);
    }
  }
}

if (elements.brandLogo && elements.easterEggMessage) {
  let hideMessageTimer = null;
  let concealMessageTimer = null;

  const scheduleHide = () => {
    if (hideMessageTimer) {
      window.clearTimeout(hideMessageTimer);
    }
    hideMessageTimer = window.setTimeout(() => {
      elements.easterEggMessage.classList.remove('is-visible');
      if (concealMessageTimer) {
        window.clearTimeout(concealMessageTimer);
      }
      concealMessageTimer = window.setTimeout(() => {
        elements.easterEggMessage.hidden = true;
        concealMessageTimer = null;
      }, 320);
      hideMessageTimer = null;
    }, 4000);
  };

  elements.brandLogo.addEventListener('dblclick', () => {
    if (hideMessageTimer) {
      window.clearTimeout(hideMessageTimer);
      hideMessageTimer = null;
    }
    if (concealMessageTimer) {
      window.clearTimeout(concealMessageTimer);
      concealMessageTimer = null;
    }
    elements.easterEggMessage.hidden = false;
    elements.easterEggMessage.classList.add('is-visible');
    scheduleHide();
  });
}

if (elements.creditBadge && elements.creditBadgeClose) {
  elements.creditBadgeClose.addEventListener('click', () => {
    elements.creditBadge.hidden = true;
    elements.creditBadge.setAttribute('aria-hidden', 'true');
  });
}

// Обработчики файлов
if (elements.violationsInput) {
  // Когда пользователь подкидывает таблицу нарушений — читаем файл и обновляем состояние.
  elements.violationsInput.addEventListener('change', async (event) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    const displayName = sanitizeDisplayFileName(file.name);
    setLoadingIndicator('violations', true, displayName);
    try {
      await loadDataset('violations', file, { displayName });
    } finally {
      setLoadingIndicator('violations', false);
    }
  });
}

if (elements.objectsInput) {
  // Аналогичный сценарий для перечня объектов.
  elements.objectsInput.addEventListener('change', async (event) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    const displayName = sanitizeDisplayFileName(file.name);
    setLoadingIndicator('objects', true, displayName);
    try {
      await loadDataset('objects', file, { displayName });
    } finally {
      setLoadingIndicator('objects', false);
    }
  });
}

if (elements.refreshButton) {
  // Принудительный пересчет отчёта — без ожидания кадра, чтобы пользователь увидел мгновенную реакцию.
  elements.refreshButton.addEventListener('click', () => {
    schedulePreviewUpdate({ immediate: true });
  });
}

if (elements.downloadButton) {
  elements.downloadButton.addEventListener('click', () => {
    if (state.lastReport && state.lastPeriods) {
      exportReportToExcel(state.lastReport, state.lastPeriods);
    }
  });
}

if (elements.dataSourceSelect) {
  elements.dataSourceSelect.addEventListener('change', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement)) {
      return;
    }
    state.dataSourceOption = target.value;
    schedulePreviewUpdate();
  });
}

const typeModeRadios = Array.from(document.querySelectorAll('input[name="type-mode"]'));
for (const radio of typeModeRadios) {
  // При переключении режима выборки типов пересобираем фильтры и обновляем таблицу.
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
    schedulePreviewUpdate();
  });
}

const violationModeRadios = Array.from(document.querySelectorAll('input[name="violation-mode"]'));
for (const radio of violationModeRadios) {
  radio.addEventListener('change', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }
    state.violationMode = target.value;
    if (state.violationMode === 'all') {
      state.customViolations = [];
    }
    renderViolationOptions();
    schedulePreviewUpdate();
  });
}

/**
 * Загружает Excel-файл, валидирует его и обновляет состояние приложения.
 * Выполняет тяжёлые операции асинхронно, чтобы не блокировать интерфейс на слабых ПК.
 * @param {('violations'|'objects')} kind тип загружаемого датасета
 * @param {File} file файл Excel из input[type="file"]
 * @param {{ displayName?: string }} [options] дополнительные параметры отображения
 */
async function loadDataset(kind, file, options = {}) {
  const displayName = sanitizeDisplayFileName(options.displayName ?? file?.name);
  try {
    assertValidFile(file, displayName);
    // Сообщаю пользователю, что файл читается.
    showPreviewMessage(`Загрузка файла «${displayName}»...`);
    // Даю браузеру кадр на отрисовку, чтобы индикатор загрузки отобразился мгновенно.
    await yieldToEventLoop();
    // В зависимости от типа файла берём подходящий набор ключевых слов для поиска заголовков.
    const headerCandidates = kind === 'violations'
      ? violationFieldDefinitions.flatMap((item) => item.candidates)
      : objectFieldDefinitions.flatMap((item) => item.candidates);
    // Считываю таблицу и получаю как записи, так и список колонок.
    const { records, headers } = await readExcelFile(file, headerCandidates);
    if (!records.length) {
      throw new Error('Не удалось обнаружить данные в выбранном файле.');
    }
    if (kind === 'violations') {
      // Для таблицы нарушений сохраняю данные и автоматически строю соответствия колонок.
      state.violations = records;
      state.violationColumns = headers;
      state.violationMapping = autoMapColumns(headers, violationFieldDefinitions);
    } else {
      // Для справочника объектов делаю ту же процедуру.
      state.objects = records;
      state.objectColumns = headers;
      state.objectMapping = autoMapColumns(headers, objectFieldDefinitions);
    }
    // После обновления данных пересобираю интерфейс и планирую обновление отчёта.
    updateVisibility();
    renderMappings();
    updateAvailableFilters();
    schedulePreviewUpdate({ immediate: true });
    showPreviewMessage('Файлы загружены, параметры можно настраивать.');
  } catch (error) {
    console.error(error);
    showPreviewMessage(error instanceof Error ? error.message : 'Не удалось обработать файл.');
  }
}

/**
 * Управляет отображением прелоадера при чтении Excel-файла.
 * Одновременно выставляет aria-атрибуты, чтобы пользователи со скринридерами получали корректный статус.
 * @param {('violations'|'objects')} kind тип загружаемого файла
 * @param {boolean} isLoading статус чтения файла
 * @param {string} [fileName] имя файла для отображения в индикаторе
 */
function setLoadingIndicator(kind, isLoading, fileName) {
  const loaderMap = {
    violations: elements.violationsLoader,
    objects: elements.objectsLoader,
  };
  const loader = loaderMap[kind];
  if (!loader) {
    return;
  }
  const safeLabel = fileName ? sanitizeDisplayFileName(fileName) : '';
  if (isLoading) {
    const textNode = loader.querySelector('.file-loader__text');
    if (textNode) {
      textNode.textContent = safeLabel ? `Обрабатываем «${safeLabel}»` : 'Обрабатываем файл';
    }
    loader.hidden = false;
  } else {
    loader.hidden = true;
  }
  loader.setAttribute('aria-hidden', String(!isLoading));
  loader.setAttribute('aria-busy', String(isLoading));
}

/**
 * Считывает первый лист Excel-файла и возвращает набор записей вместе с заголовками.
 * @param {File} file Excel-файл, уже прошедший базовую валидацию
 * @param {string[]} headerKeywords список подсказок для поиска строки заголовков
 * @returns {Promise<{records: Record<string, unknown>[], headers: string[]}>}
 */
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
  if (headers.length > MAX_COLUMNS_PER_SHEET) {
    throw new Error(
      `Файл содержит ${numberFormatter.format(headers.length)} столбцов. Допустимо не более ${numberFormatter.format(MAX_COLUMNS_PER_SHEET)}.`,
    );
  }
  await yieldToEventLoop();
  const records = extractRecords(sheet, headerRowIndex, headers, { maxRows: MAX_ROWS_PER_SHEET });
  return { records, headers };
}

/**
 * Находит индекс строки, которая с наибольшей вероятностью содержит заголовки таблицы.
 * @param {XLSX.WorkSheet} sheet лист Excel
 * @param {{ s: { r: number, c: number }, e: { r: number, c: number } }} range рабочий диапазон листа
 * @param {string[]} headerKeywords подсказки для поиска заголовка
 * @returns {number|null} индекс строки с заголовками либо null, если ничего не нашли
 */
function detectHeaderRowIndex(sheet, range, headerKeywords) {
  const normalizedKeywords = headerKeywords.map((item) => normalizeHeaderValue(item));
  let bestRowIndex = null;
  let bestScore = -Infinity;
  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    // Для каждой строки считаю, сколько в ней текстовых ячеек и совпадений с ожидаемыми названиями.
    let matchCount = 0;
    let textCount = 0;
    let valueCount = 0;
    for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      const cell = sheet[address];
      if (!cell) {
        continue;
      }
      // Excel может хранить значение как v или форматированную строку w, поэтому подстраховываюсь.
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
    // Чуть сильнее весим совпадения по ключевым словам, но при этом учитываем общее количество текста.
    const score = matchCount * 10 + textCount;
    if (score > bestScore) {
      bestScore = score;
      bestRowIndex = rowIndex;
    }
  }
  return bestRowIndex;
}

/**
 * Приводит значение заголовка к нижнему регистру без лишних пробелов.
 * @param {unknown} value произвольное значение из ячейки
 * @returns {string} нормализованная строка
 */
function normalizeHeaderValue(value) {
  if (typeof value === 'string') {
    return value.replace(/\s+/g, ' ').trim().toLowerCase();
  }
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim().toLowerCase();
}

/**
 * Собирает строку заголовков и делает названия столбцов уникальными.
 * @param {XLSX.WorkSheet} sheet лист Excel
 * @param {number} rowIndex индекс строки с заголовками
 * @param {{ s: { r: number, c: number }, e: { r: number, c: number } }} range рабочий диапазон листа
 * @returns {string[]} массив уникальных заголовков
 */
function extractHeaderRow(sheet, rowIndex, range) {
  const headers = [];
  for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
    const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
    const cell = sheet[address];
    headers.push(typeof cell?.v === 'string' ? cell.v.trim() : cell?.v ?? `Колонка ${colIndex + 1}`);
  }
  return makeHeadersUnique(headers);
}

/**
 * Формирует массив объектов данных на основе заголовков, пропуская полностью пустые строки.
 * Дополнительно ограничивает число обрабатываемых строк для защиты производительности.
 * @param {XLSX.WorkSheet} sheet лист Excel
 * @param {number} headerRowIndex индекс строки заголовков
 * @param {string[]} headers список заголовков
 * @param {{ maxRows?: number }} [options] ограничения на количество строк
 * @returns {Record<string, unknown>[]} массив строк данных
 */
function extractRecords(sheet, headerRowIndex, headers, options = {}) {
  const { maxRows = Infinity } = options;
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
      if (records.length > maxRows) {
        throw new Error(
          `Файл содержит более ${numberFormatter.format(maxRows)} строк с данными. Сократите выгрузку или используйте несколько файлов.`,
        );
      }
    }
  }
  return records;
}

/**
 * Аккуратно извлекает значение из ячейки Excel без потери типа данных.
 * @param {XLSX.CellObject | undefined} cell ячейка листа
 * @returns {unknown} значение из ячейки
 */
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

// Строю авто-сопоставление «ключ поля → колонка» по заранее заданным синонимам.
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
    // Если не нашли подходящую колонку — оставляю пустую ячейку, чтобы пользователь явно подтвердил выбор.
    mapping[definition.key] = matchedColumn;
  }
  return mapping;
}

// На основе текущих данных рисую выпадающие списки с сопоставлением колонок.
function renderMappings() {
  elements.mappingSection.hidden = !(state.violations.length || state.objects.length);
  elements.violationsMapping.innerHTML = '';
  elements.objectsMapping.innerHTML = '';
  if (state.violations.length) {
    renderMappingList(elements.violationsMapping, state.violationColumns, violationFieldDefinitions, state.violationMapping, (key, value) => {
      state.violationMapping[key] = value;
      updateAvailableFilters();
      schedulePreviewUpdate();
    });
  }
  if (state.objects.length) {
    renderMappingList(elements.objectsMapping, state.objectColumns, objectFieldDefinitions, state.objectMapping, (key, value) => {
      state.objectMapping[key] = value;
      updateAvailableFilters();
      schedulePreviewUpdate();
    });
  }
}

// Вспомогательная функция, которая собирает список select'ов для конкретного набора полей.
function renderMappingList(container, columns, definitions, mapping, onChange) {
  const fragment = document.createDocumentFragment();
  for (const definition of definitions) {
    const wrapper = document.createElement('div');
    wrapper.className = 'mapping-item';

    const label = document.createElement('label');
    label.textContent = definition.label;

    const select = document.createElement('select');
    const placeholderOption = document.createElement('option');
    placeholderOption.value = '';
    placeholderOption.textContent = '— Не выбрано —';
    if (!mapping[definition.key]) {
      placeholderOption.selected = true;
    }
    select.append(placeholderOption);
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
    fragment.append(wrapper);
  }
  container.append(fragment);
}

// Управляю видимостью секций, чтобы интерфейс не расползался до загрузки обоих файлов.
function updateVisibility() {
  const ready = state.violations.length && state.objects.length;
  elements.mappingSection.hidden = !(state.violations.length || state.objects.length);
  elements.controlsSection.hidden = !ready;
  elements.previewSection.hidden = !ready;
  if (!ready) {
    clearTable();
  }
}

// Пересобираю списки доступных дат и типов объектов исходя из свежих данных и выбранных колонок.
function updateAvailableFilters() {
  if (!(state.violations.length && state.objects.length)) {
    return;
  }
  const dateColumn = state.violationMapping.inspectionDate;
  if (dateColumn) {
    state.availableDates = extractUniqueDates(state.violations, dateColumn);
    updateDatePicker(datePickers.currentStart, state.availableDates, 0);
    updateDatePicker(datePickers.currentEnd, state.availableDates, state.availableDates.length - 1);
    const previousEndIndex = state.availableDates.length > 1 ? state.availableDates.length - 2 : state.availableDates.length - 1;
    updateDatePicker(datePickers.previousStart, state.availableDates, 0);
    updateDatePicker(datePickers.previousEnd, state.availableDates, previousEndIndex);
  } else {
    state.availableDates = [];
    updateDatePicker(datePickers.currentStart, state.availableDates, null);
    updateDatePicker(datePickers.currentEnd, state.availableDates, null);
    updateDatePicker(datePickers.previousStart, state.availableDates, null);
    updateDatePicker(datePickers.previousEnd, state.availableDates, null);
  }
  const objectTypeColumn = state.objectMapping.objectType;
  const objectTypes = collectUniqueValues(state.objects, objectTypeColumn).sort((a, b) => a.localeCompare(b, 'ru'));
  state.availableTypes = objectTypes;
  // Очищаю выбор от устаревших значений, если пользователь менял сопоставление колонок.
  state.customTypes = state.customTypes.filter((value) => state.availableTypes.includes(value));
  renderTypeOptions();
  const violationNameColumn = state.violationMapping.violationName;
  if (violationNameColumn) {
    const violationNames = collectUniqueValues(state.violations, violationNameColumn).sort((a, b) => a.localeCompare(b, 'ru'));
    state.availableViolations = violationNames;
    state.customViolations = state.customViolations.filter((value) => state.availableViolations.includes(value));
  } else {
    state.availableViolations = [];
    state.customViolations = [];
  }
  renderViolationOptions();
  updateDataSourceControl();
}

function updateDatePicker(picker, dates, defaultIndex) {
  if (!picker) {
    return;
  }
  const availableValues = new Set(dates.map((item) => item.iso));
  const previousValue = picker.getValue();
  picker.setAvailableDates(dates);
  let nextValue = '';
  if (previousValue && availableValues.has(previousValue)) {
    nextValue = previousValue;
  } else if (typeof defaultIndex === 'number' && dates.length) {
    const index = Math.min(Math.max(defaultIndex, 0), dates.length - 1);
    nextValue = dates[index]?.iso ?? '';
  }
  picker.setValue(nextValue);
}

function renderSearchableMultiselect(options) {
  const {
    container,
    values,
    selectedValues,
    disabled,
    placeholder,
    emptyLabel,
    limit,
    limitMessage,
    messageElement,
    onAdd,
    onRemove,
  } = options;
  if (!container) {
    return;
  }
  if (messageElement) {
    messageElement.textContent = '';
  }
  container.innerHTML = '';
  container.classList.remove('multi-select-empty');
  if (!values.length) {
    container.textContent = emptyLabel;
    container.classList.add('multi-select-empty');
    return;
  }
  const wrapper = document.createElement('div');
  wrapper.className = 'multi-select';
  if (disabled) {
    wrapper.classList.add('is-disabled');
  }
  const field = document.createElement('div');
  field.className = 'multi-select-field';
  const tags = document.createElement('div');
  tags.className = 'multi-select-tags';
  for (const value of selectedValues) {
    const tag = document.createElement('span');
    tag.className = 'multi-select-tag';
    tag.textContent = value;
    const removeButton = document.createElement('button');
    removeButton.type = 'button';
    removeButton.className = 'multi-select-tag-remove';
    removeButton.setAttribute('aria-label', `Удалить «${value}»`);
    removeButton.textContent = '×';
    removeButton.disabled = disabled;
    removeButton.addEventListener('click', () => {
      onRemove?.(value);
    });
    tag.append(removeButton);
    tags.append(tag);
  }
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'multi-select-input';
  input.placeholder = disabled ? 'Активируйте пользовательский режим' : placeholder;
  input.autocomplete = 'off';
  input.disabled = disabled;
  field.append(tags, input);
  wrapper.append(field);
  const dropdown = document.createElement('div');
  dropdown.className = 'multi-select-dropdown';
  wrapper.append(dropdown);
  container.append(wrapper);
  const selectedSet = new Set(selectedValues);

  const renderDropdown = () => {
    dropdown.innerHTML = '';
    if (disabled) {
      const note = document.createElement('p');
      note.className = 'multi-select-note';
      note.textContent = 'Переключите режим «Выбрать», чтобы воспользоваться поиском и выбором.';
      dropdown.append(note);
      return;
    }
    const query = input.value.trim().toLowerCase();
    const filtered = values.filter((value) => !selectedSet.has(value) && value.toLowerCase().includes(query));
    if (!filtered.length) {
      const empty = document.createElement('p');
      empty.className = 'multi-select-note';
      empty.textContent = query ? 'Совпадений не найдено.' : 'Все значения уже добавлены.';
      dropdown.append(empty);
      return;
    }
    const canAddMore = selectedValues.length < limit;
    if (!canAddMore) {
      if (messageElement) {
        messageElement.textContent = limitMessage;
      }
    } else if (messageElement && messageElement.textContent === limitMessage) {
      messageElement.textContent = '';
    }
    for (const value of filtered) {
      const optionButton = document.createElement('button');
      optionButton.type = 'button';
      optionButton.className = 'multi-select-option';
      optionButton.textContent = value;
      optionButton.disabled = !canAddMore;
      optionButton.addEventListener('mousedown', (event) => {
        event.preventDefault();
      });
      optionButton.addEventListener('click', () => {
        if (!canAddMore) {
          if (messageElement) {
            messageElement.textContent = limitMessage;
          }
          return;
        }
        onAdd?.(value);
      });
      dropdown.append(optionButton);
    }
  };

  renderDropdown();

  if (!disabled) {
    const openDropdown = () => {
      wrapper.classList.add('is-open');
      renderDropdown();
    };
    const closeDropdown = () => {
      wrapper.classList.remove('is-open');
      input.value = '';
      renderDropdown();
    };
    wrapper.addEventListener('mousedown', (event) => {
      if (event.target === wrapper || event.target === field || event.target === tags) {
        event.preventDefault();
        input.focus();
      }
    });
    input.addEventListener('focus', openDropdown);
    input.addEventListener('input', openDropdown);
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        const query = input.value.trim().toLowerCase();
        const candidates = values.filter((value) => !selectedSet.has(value) && value.toLowerCase().includes(query));
        if (!candidates.length) {
          return;
        }
        if (selectedValues.length >= limit) {
          if (messageElement) {
            messageElement.textContent = limitMessage;
          }
          return;
        }
        onAdd?.(candidates[0]);
      } else if (event.key === 'Escape') {
        event.preventDefault();
        closeDropdown();
      }
    });
    input.addEventListener('blur', () => {
      setTimeout(() => {
        closeDropdown();
      }, 120);
    });
  }
}

function renderTypeOptions() {
  renderSearchableMultiselect({
    container: elements.typeOptions,
    values: state.availableTypes,
    selectedValues: state.customTypes,
    disabled: state.typeMode !== 'custom',
    placeholder: 'Начните вводить тип объекта…',
    emptyLabel: 'Типы объектов не найдены.',
    limit: 3,
    limitMessage: 'Можно выбрать не более трех типов объектов.',
    messageElement: elements.typeMessage,
    onAdd: (value) => {
      if (state.customTypes.includes(value) || state.customTypes.length >= 3) {
        if (elements.typeMessage) {
          elements.typeMessage.textContent = 'Можно выбрать не более трех типов объектов.';
        }
        return;
      }
      state.customTypes = [...state.customTypes, value];
      if (elements.typeMessage) {
        elements.typeMessage.textContent = '';
      }
      schedulePreviewUpdate();
      renderTypeOptions();
    },
    onRemove: (value) => {
      state.customTypes = state.customTypes.filter((item) => item !== value);
      if (elements.typeMessage) {
        elements.typeMessage.textContent = '';
      }
      schedulePreviewUpdate();
      renderTypeOptions();
    },
  });
}

function renderViolationOptions() {
  renderSearchableMultiselect({
    container: elements.violationOptions,
    values: state.availableViolations,
    selectedValues: state.customViolations,
    disabled: state.violationMode !== 'custom',
    placeholder: 'Поиск наименования нарушения…',
    emptyLabel: 'Наименования нарушений не найдены.',
    limit: 5,
    limitMessage: 'Можно выбрать не более пяти наименований нарушений.',
    messageElement: elements.violationMessage,
    onAdd: (value) => {
      if (state.customViolations.includes(value) || state.customViolations.length >= 5) {
        if (elements.violationMessage) {
          elements.violationMessage.textContent = 'Можно выбрать не более пяти наименований нарушений.';
        }
        return;
      }
      state.customViolations = [...state.customViolations, value];
      if (elements.violationMessage) {
        elements.violationMessage.textContent = '';
      }
      schedulePreviewUpdate();
      renderViolationOptions();
    },
    onRemove: (value) => {
      state.customViolations = state.customViolations.filter((item) => item !== value);
      if (elements.violationMessage) {
        elements.violationMessage.textContent = '';
      }
      schedulePreviewUpdate();
      renderViolationOptions();
    },
  });
}

// Настройки источника данных требуют отдельной обработки.
function updateDataSourceControl() {
  if (!elements.dataSourceSelect || !elements.dataSourceMessage) {
    return;
  }
  const previousOption = state.dataSourceOption;
  const column = state.violationMapping.dataSource;
  if (!column) {
    state.dataSourceOption = 'all';
    elements.dataSourceSelect.value = 'all';
    elements.dataSourceSelect.disabled = true;
    elements.dataSourceMessage.textContent = 'Сопоставьте колонку «Источник данных», чтобы выбрать источник отчёта.';
    if (previousOption !== state.dataSourceOption) {
      schedulePreviewUpdate();
    }
    return;
  }
  elements.dataSourceSelect.disabled = false;
  elements.dataSourceSelect.value = state.dataSourceOption;
  const summary = summarizeDataSourceCategories(column);
  if (summary.unknown) {
    elements.dataSourceMessage.textContent = 'Неизвестные значения источника учитываются только при выборе «ОАТИ и ЦАФАП».';
  } else if (!summary.oati && !summary.cafap && !summary.both) {
    elements.dataSourceMessage.textContent = 'В выбранной колонке отсутствуют значения источника данных.';
  } else {
    elements.dataSourceMessage.textContent = '';
  }
}

function summarizeDataSourceCategories(column) {
  const summary = { oati: false, cafap: false, both: false, unknown: false };
  for (const record of state.violations) {
    const rawValue = record[column];
    const category = categorizeDataSource(rawValue);
    if (!category) {
      if (getValueAsString(rawValue)) {
        summary.unknown = true;
      }
      continue;
    }
    if (category === 'oati') {
      summary.oati = true;
    } else if (category === 'cafap') {
      summary.cafap = true;
    } else if (category === 'both') {
      summary.both = true;
    } else {
      summary.unknown = true;
    }
  }
  return summary;
}

// Управляю состоянием кнопки выгрузки, чтобы не давать скачивать устаревшие данные.
function setExportAvailability(isEnabled) {
  if (!elements.downloadButton) {
    return;
  }
  elements.downloadButton.disabled = !isEnabled;
  elements.downloadButton.setAttribute('aria-disabled', String(!isEnabled));
}

// Каждое новое пересчитывание отчёта сбрасывает кеш и блокирует выгрузку до готовности.
function resetExportState() {
  state.lastReport = null;
  state.lastPeriods = null;
  setExportAvailability(false);
}

// Если нужно пересчитать отчёт прямо сейчас, то отменяю ранее запланированное обновление.
function cancelScheduledPreviewUpdate() {
  if (scheduledPreviewHandle === null) {
    return;
  }
  if (scheduledPreviewKind === 'raf' && typeof cancelAnimationFrame === 'function') {
    cancelAnimationFrame(scheduledPreviewHandle);
  }
  if (scheduledPreviewKind === 'timeout') {
    clearTimeout(scheduledPreviewHandle);
  }
  scheduledPreviewHandle = null;
  scheduledPreviewKind = null;
}

/**
 * Планирует пересчёт отчёта. Несколько событий объединяются в одно обновление, что снижает нагрузку.
 * @param {{ immediate?: boolean }} [options] флаг немедленного пересчёта без ожидания кадра
 */
function schedulePreviewUpdate(options = {}) {
  const { immediate = false } = options;
  if (immediate) {
    // Для принудительного обновления (кнопка «Пересчитать») отменяем очередь и выполняем пересчёт синхронно.
    cancelScheduledPreviewUpdate();
    runPreviewUpdate();
    return;
  }
  if (scheduledPreviewHandle !== null) {
    return;
  }
  const runner = () => {
    scheduledPreviewHandle = null;
    scheduledPreviewKind = null;
    runPreviewUpdate();
  };
  if (typeof requestAnimationFrame === 'function') {
    scheduledPreviewHandle = requestAnimationFrame(runner);
    scheduledPreviewKind = 'raf';
    return;
  }
  scheduledPreviewHandle = setTimeout(runner, 50);
  scheduledPreviewKind = 'timeout';
}

/**
 * Основная точка входа для пересчёта отчёта.
 * Проверяет корректность настроек, пересобирает таблицу и управляет состоянием экспорта.
 */
function runPreviewUpdate() {
  resetExportState();
  if (!(state.violations.length && state.objects.length)) {
    // Без двух файлов смысла продолжать нет — показываю пользователю подсказку и очищаю таблицу.
    showPreviewMessage('Загрузите оба файла, чтобы получить отчет.');
    clearTable();
    return;
  }
  if (state.typeMode === 'custom' && !state.customTypes.length) {
    showPreviewMessage('Выберите типы объектов для пользовательского режима.');
    clearTable();
    return;
  }
  if (state.violationMode === 'custom' && !state.customViolations.length) {
    showPreviewMessage('Выберите наименования нарушений для пользовательского режима.');
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
    // Если после фильтрации ничего не осталось, честно об этом предупреждаю.
    showPreviewMessage('По заданным условиям данные не найдены. Измените фильтры.');
    clearTable();
    return;
  }
  showPreviewMessage('');
  state.lastReport = report;
  state.lastPeriods = periods;
  renderReportTable(report, periods);
  setExportAvailability(true);
}

// Проверяю, что пользователь проставил все необходимые соответствия.
function isMappingComplete(mapping, definitions) {
  return definitions.every((definition) => definition.optional || Boolean(mapping[definition.key]));
}

// Собираю выбранные диапазоны дат и параллельно валидирую ввод.
function getSelectedPeriods() {
  const currentStartIso = elements.currentStart?.value ?? '';
  const currentEndIso = elements.currentEnd?.value ?? '';
  const previousStartIso = elements.previousStart?.value ?? '';
  const previousEndIso = elements.previousEnd?.value ?? '';
  if (!currentStartIso || !currentEndIso) {
    // Без отчётного периода отчёт не построить.
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

// Формирую человекочитаемое отображение округа: сначала стараюсь подобрать сокращение.
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

// Привожу обозначение округа к ключу, пригодному для сравнения и агрегации.
function normalizeDistrictKey(value) {
  const normalized = normalizeKey(value);
  if (!normalized) {
    return '';
  }
  const withoutParentheses = normalized.replace(/\s*\(.*?\)\s*/g, ' ');
  const withoutCityMention = withoutParentheses.replace(/\s+г\.?\s*москв[аеы]?$/u, '');
  return withoutCityMention.replace(/\s+/g, ' ').trim();
}

// Собираю словарь округов, чтобы грамотно сводить надписи вида «г. Москва, ЦАО» к единому виду.
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

// Решаю, стоит ли заменить уже сохранённое название округа на новое (например, на более короткое сокращение).
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

/**
 * Агрегирует показатели по округам и формирует итоговую строку отчёта.
 * @param {{ current: { start: Date, end: Date }, previous: { start: Date, end: Date } }} periods выбранные периоды
 * @returns {{ rows: Array<Record<string, any>>, totalRow: Record<string, any> }} строки детального отчёта и суммарная строка
 */
function buildReport(periods) {
  const violationMapping = state.violationMapping;
  const objectMapping = state.objectMapping;
  const typePredicate = createTypePredicate();
  const violationPredicate = createViolationPredicate();
  const dataSourcePredicate = createDataSourcePredicate();
  // Здесь буду копить агрегированные данные по округам.
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
    const aggregatedKey = normalizeKey(displayLabel) || DISTRICT_FALLBACK_KEY;
    if (!districtData.has(aggregatedKey)) {
      districtData.set(aggregatedKey, {
        label: displayLabel,
        totalObjects: new Set(),
        inspectedObjects: new Set(),
        objectsWithDetectedViolations: new Set(),
        currentViolationIds: new Set(),
        previousControlIds: new Set(),
        resolvedIds: new Set(),
        controlIds: new Set(),
      });
    } else {
      const entry = districtData.get(aggregatedKey);
      if (entry && shouldReplaceDistrictLabel(entry.label, displayLabel)) {
        entry.label = displayLabel;
      }
    }
    return districtData.get(aggregatedKey);
  };

  const allowedViolationObjects = new Set();
  if (state.violationMode === 'custom' && state.customViolations.length) {
    const violationNameColumn = violationMapping.violationName;
    const violationObjectNameColumn = violationMapping.objectName;
    if (violationNameColumn && violationObjectNameColumn) {
      for (const record of state.violations) {
        if (!dataSourcePredicate(record)) {
          continue;
        }
        const violationName = getValueAsString(record[violationNameColumn]);
        if (!violationName || !violationPredicate(violationName)) {
          continue;
        }
        const relatedObjectName = getValueAsString(record[violationObjectNameColumn]);
        if (!relatedObjectName) {
          continue;
        }
        allowedViolationObjects.add(normalizeKey(relatedObjectName));
      }
    }
  }

  // Сначала прохожусь по объектам и считаю общее количество по округам.
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
    if (state.violationMode === 'custom') {
      if (!allowedViolationObjects.size) {
        continue;
      }
      if (!allowedViolationObjects.has(normalizeKey(objectName))) {
        continue;
      }
    }
    const entry = ensureEntry(districtLabel);
    entry.totalObjects.add(objectName);
  }

  // Затем обрабатываю таблицу нарушений, чтобы посчитать динамику и статусы.
  const inspectionResultColumn = violationMapping.inspectionResult;

  for (const record of state.violations) {
    if (!dataSourcePredicate(record)) {
      continue;
    }
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
    const violationName = getValueAsString(record[violationMapping.violationName]);
    if (violationName && !violationPredicate(violationName)) {
      continue;
    }
    if (!violationName && state.violationMode === 'custom') {
      continue;
    }
    const normalizedStatus = normalizeText(status);
    const dateValue = record[violationMapping.inspectionDate];
    const inspectionDate = parseDateValue(dateValue);
    if (!inspectionDate) {
      continue;
    }
    const entry = ensureEntry(districtLabel);
    const inspectionResultValue = inspectionResultColumn
      ? getValueAsString(record[inspectionResultColumn])
      : '';
    const normalizedInspectionResult = normalizeText(inspectionResultValue);

    if (isWithinPeriod(inspectionDate, periods.current)) {
      if (objectName) {
        entry.inspectedObjects.add(objectName);
      }
      if (
        inspectionResultColumn &&
        objectName &&
        normalizedInspectionResult === INSPECTION_RESULT_VIOLATION
      ) {
        entry.objectsWithDetectedViolations.add(objectName);
      }
      if (violationId) {
        entry.currentViolationIds.add(violationId);
      }
      if (normalizedStatus === RESOLVED_STATUS && violationId) {
        entry.resolvedIds.add(violationId);
      }
      if (violationId && CONTROL_STATUS_SET.has(normalizedStatus)) {
        entry.controlIds.add(violationId);
      }
    }
    if (isWithinPeriod(inspectionDate, periods.previous)) {
      if (CONTROL_STATUS_SET.has(normalizedStatus)) {
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
    objectsWithDetectedViolations: 0,
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
    const objectsWithDetectedViolationsCount = entry.objectsWithDetectedViolations.size;
    const currentViolationsCount = entry.currentViolationIds.size;
    const previousControlCount = entry.previousControlIds.size;
    const resolvedCount = entry.resolvedIds.size;
    const onControlCount = entry.controlIds.size;
    // Нарушения «всего» считаю как объединение текущих и переходящих на контроль.
    const totalViolationsCount = new Set([...entry.currentViolationIds, ...entry.previousControlIds]).size;

    rows.push({
      label: entry.label,
      totalObjects: totalObjectsCount,
      inspectedObjects: inspectedCount,
      inspectedPercent: computePercent(inspectedCount, totalObjectsCount),
      violationPercent: computePercent(objectsWithDetectedViolationsCount, inspectedCount),
      totalViolations: totalViolationsCount,
      currentViolations: currentViolationsCount,
      previousControl: previousControlCount,
      resolved: resolvedCount,
      onControl: onControlCount,
    });

    totals.totalObjects += totalObjectsCount;
    totals.inspectedObjects += inspectedCount;
    totals.objectsWithDetectedViolations += objectsWithDetectedViolationsCount;
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
    label: 'ИТОГО',
    totalObjects: totals.totalObjects,
    inspectedObjects: totals.inspectedObjects,
    inspectedPercent: computePercent(totals.inspectedObjects, totals.totalObjects),
    violationPercent: computePercent(totals.objectsWithDetectedViolations, totals.inspectedObjects),
    totalViolations: totals.totalViolations.size,
    currentViolations: totals.currentViolations.size,
    previousControl: totals.previousControl.size,
    resolved: totals.resolved.size,
    onControl: totals.onControl.size,
  };

  return { rows, totalRow };
}

/**
 * Строит HTML-таблицу отчёта на основе агрегированных данных.
 * @param {{ rows: Array<Record<string, any>>, totalRow: Record<string, any> }} report агрегированные строки
 * @param {{ current: { start: Date, end: Date }, previous: { start: Date, end: Date } }} periods выбранные периоды
 */
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

// Создаю строку таблицы с нужными форматами чисел и процентов.
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

// Универсальный помощник для создания ячеек: пригодится и для th, и для td.
function createCell(value, isHeader = false) {
  const cell = document.createElement(isHeader ? 'th' : 'td');
  cell.textContent = value ?? '';
  return cell;
}

// Формирую заголовки таблицы, включая динамический текст по выбранному периоду.
function buildTableHeaders(periods) {
  const totalHeader = buildTotalHeader();
  const rangeHeader = `Проверено ОДХ с ${formatDateDisplay(periods.current.start)} по ${formatDateDisplay(periods.current.end)}`;
  return [
    'Округ',
    totalHeader,
    rangeHeader,
    '% проверенных объектов от общего количества ОДХ',
    '% объектов с нарушениями',
    'Всего нарушений',
    'Нарушения, выявленные за отчётный период',
    'Нарушения, находящиеся на контроле с предыдущей проверки',
    'Устранено нарушений',
    'Нарушения на контроле',
  ];
}

// Подбираю надпись для столбца «Всего», учитывая выбранные типы объектов.
function buildTotalHeader() {
  if (state.typeMode === 'all' || !state.customTypes.length) {
    return 'Всего ОДХ';
  }
  if (state.customTypes.length === 1) {
    return `Всего ${state.customTypes[0]}`;
  }
  return `Всего (${state.customTypes.join(', ')})`;
}

function buildExcelHeaderNumbers(columnCount) {
  const template = [1, 2, 3, 4, '4.1', '4.2', 5, 6, '6.1', '6.2'];
  if (columnCount <= template.length) {
    return template.slice(0, columnCount);
  }
  const extended = template.slice();
  for (let index = template.length; index < columnCount; index += 1) {
    extended.push(String(index + 1));
  }
  return extended;
}

function buildExcelTitle(periods) {
  const periodRange = formatTitlePeriod(periods?.current);
  const parts = ['Нарушения на ОДХ'];
  if (periodRange) {
    parts[0] += ` (отчёт за ${periodRange})`;
  }
  const dataSource = describeDataSourceForTitle();
  if (dataSource) {
    parts.push(dataSource);
  }
  return parts.join(', ');
}

function describeDataSourceForTitle() {
  switch (state.dataSourceOption) {
    case 'oati':
      return 'накопленные только ОАТИ';
    case 'cafap':
      return 'накопленные только ЦАФАП';
    default:
      return 'выявленные ОАТИ и ЦАФАП';
  }
}

// Очищаю таблицу — пригодится в сценариях с ошибками и сменой фильтров.
function clearTable() {
  elements.reportTable.innerHTML = '';
}

// Сообщения пользователю отображаю в одном месте, чтобы не размазывать текст по коду.
function showPreviewMessage(message) {
  elements.previewMessage.textContent = message;
}

// Превращаю ISO-строку вида «2024-03-15» в объект Date без временной части.
function parseIsoDate(iso) {
  const [year, month, day] = iso.split('-').map((part) => Number.parseInt(part, 10));
  return new Date(year, month - 1, day);
}

// Готовлю предикат для фильтрации по типам объектов с учётом выбранного режима.
function createTypePredicate() {
  if (state.typeMode === 'all' || !state.customTypes.length) {
    return () => true;
  }
  const allowed = new Set(state.customTypes.map((value) => normalizeKey(value)));
  return (value) => allowed.has(normalizeKey(value));
}

function createViolationPredicate() {
  if (state.violationMode === 'all' || !state.customViolations.length) {
    return () => true;
  }
  const allowed = new Set(state.customViolations.map((value) => normalizeKey(value)));
  return (value) => allowed.has(normalizeKey(value));
}

function createDataSourcePredicate() {
  const column = state.violationMapping.dataSource;
  if (!column) {
    return () => true;
  }
  const mode = state.dataSourceOption;
  if (mode === 'all') {
    return () => true;
  }
  return (record) => {
    const category = categorizeDataSource(record[column]);
    if (!category) {
      return false;
    }
    if (mode === 'oati') {
      return category === 'oati';
    }
    if (mode === 'cafap') {
      return category === 'cafap';
    }
    return true;
  };
}

// Универсальный способ привести строку к нижнему регистру и убрать лишние пробелы.
function normalizeKey(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return value.toString().trim().toLowerCase();
}

// Просто обёртка над normalizeKey — оставил для лучшей читаемости кода выше.
function normalizeText(value) {
  return normalizeKey(value);
}

// Перевожу значение в строку так, чтобы его можно было спокойно показывать в интерфейсе.
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

function categorizeDataSource(value) {
  const text = getValueAsString(value);
  const normalized = normalizeText(text);
  if (!normalized) {
    return '';
  }
  const hasOati = normalized.includes('оати');
  const hasCafap = normalized.includes('цафап');
  if (hasOati && hasCafap) {
    return 'both';
  }
  if (hasOati) {
    return 'oati';
  }
  if (hasCafap) {
    return 'cafap';
  }
  return '';
}

// Выдираю уникальные даты из нужной колонки и сортирую их.
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

// Привожу значение из Excel к Date, учитывая возможные форматы (число, строка, объект Date).
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
    // Поддерживаю формат «дд.мм.гггг» с необязательным временем.
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

// Простой форматтер для ISO-строки — пригодился в select'ах и названиях файлов.
function formatIsoDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// Отображаю даты в привычном формате «дд.мм.гггг».
function formatDateDisplay(date) {
  if (!(date instanceof Date)) {
    return '';
  }
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
}

function formatDateDisplayShort(date) {
  if (!(date instanceof Date)) {
    return '';
  }
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  return `${day}.${month}`;
}

// Собираю уникальные строки из указанной колонки, сохраняя оригинальное написание.
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

// Проверяю, попадает ли дата в выбранный диапазон (с учётом только календарной части).
function isWithinPeriod(date, period) {
  if (!period) {
    return false;
  }
  const dateOnly = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  return dateOnly >= period.start && dateOnly <= period.end;
}

// Расчёт процента с защитой от деления на ноль.
function computePercent(part, total) {
  if (!total || !Number.isFinite(part)) {
    return 0;
  }
  return (part / total) * 100;
}

// Форматирую целые значения с пробелами-разделителями.
function formatInteger(value) {
  if (!Number.isFinite(value)) {
    return '0';
  }
  return numberFormatter.format(Math.round(value));
}

// Проценты показываю с точностью до одного десятичного знака.
function formatPercent(value) {
  if (!Number.isFinite(value)) {
    return '0';
  }
  return value.toFixed(1);
}

/**
 * Формирует Excel-файл с отчётом, добавляя форматирование и стили, близкие к утверждённому шаблону.
 * @param {{ rows: Array<Record<string, any>>, totalRow: Record<string, any> }} report агрегированные данные
 * @param {{ current: { start: Date, end: Date }, previous: { start: Date, end: Date } }} periods выбранные периоды
 */
function exportReportToExcel(report, periods) {
  try {
    const headers = buildTableHeaders(periods);
    const columnCount = headers.length;
    const headerNumbers = buildExcelHeaderNumbers(columnCount);
    const title = buildExcelTitle(periods);

    const createEmptyRow = () => Array.from({ length: columnCount }, () => null);
    const aoa = [];

    aoa.push(createEmptyRow());
    const titleRowIndex = aoa.length;
    aoa.push([title, ...Array(Math.max(0, columnCount - 1)).fill(null)]);
    aoa.push(createEmptyRow());
    const headerRowIndex = aoa.length;
    aoa.push(headers);
    const numberRowIndex = aoa.length;
    aoa.push(headerNumbers);

    const dataStartRow = aoa.length;
    aoa.push(buildExportDataRow(report.totalRow));
    const detailStartRow = aoa.length;
    for (const row of report.rows) {
      aoa.push(buildExportDataRow(row));
    }
    const lastDataRowIndex = aoa.length - 1;

    const sheet = XLSX.utils.aoa_to_sheet(aoa);
    sheet['!merges'] = [
      { s: { r: titleRowIndex, c: 0 }, e: { r: titleRowIndex, c: columnCount - 1 } },
    ];

    const baseColumnWidths = [26, 16, 22, 28, 22, 20, 34, 34, 24, 24];
    sheet['!cols'] = Array.from({ length: columnCount }, (_, index) => ({
      wch: baseColumnWidths[index] ?? 20,
    }));

    sheet['!rows'] = sheet['!rows'] ?? [];
    sheet['!rows'][titleRowIndex] = { hpt: 26 };
    sheet['!rows'][headerRowIndex] = { hpt: 48 };
    sheet['!rows'][numberRowIndex] = { hpt: 22 };
    sheet['!rows'][dataStartRow] = { hpt: 24 };

    const baseFont = { name: 'Times New Roman', color: { rgb: 'FF1F2937' } };
    const borderColor = 'FF94A3B8';
    const titleStyle = {
      font: { ...baseFont, bold: true, sz: 14 },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      fill: { fgColor: { rgb: 'FFE2E8F0' } },
      border: buildBorderStyle(borderColor),
    };
    const headerStyle = {
      font: { ...baseFont, bold: true },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      fill: { fgColor: { rgb: 'FFCBD5F5' } },
      border: buildBorderStyle(borderColor),
    };
    const numberingStyle = {
      font: { ...baseFont, bold: true },
      alignment: { horizontal: 'center', vertical: 'center' },
      fill: { fgColor: { rgb: 'FFE0E7FF' } },
      border: buildBorderStyle(borderColor),
    };
    const baseDataStyle = {
      font: { ...baseFont },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: buildBorderStyle(borderColor),
    };
    const zebraDataStyle = {
      fill: { fgColor: { rgb: 'FFF8FAFC' } },
    };
    const totalRowStyle = {
      font: { ...baseFont, bold: true, sz: 12 },
      alignment: { horizontal: 'center', vertical: 'center' },
      fill: { fgColor: { rgb: 'FFFDEBD3' } },
      border: buildBorderStyle(borderColor),
    };
    const firstColumnStyle = {
      alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    };

    for (let colIndex = 0; colIndex < columnCount; colIndex += 1) {
      setCellStyle(sheet, titleRowIndex, colIndex, titleStyle);
      setCellStyle(sheet, headerRowIndex, colIndex, headerStyle);
      setCellStyle(sheet, numberRowIndex, colIndex, numberingStyle);
    }

    for (let colIndex = 0; colIndex < columnCount; colIndex += 1) {
      const styles = [totalRowStyle];
      if (colIndex === 0) {
        styles.push(firstColumnStyle);
      }
      setCellStyle(sheet, dataStartRow, colIndex, ...styles);
    }

    if (detailStartRow <= lastDataRowIndex) {
      for (let rowIndex = detailStartRow; rowIndex <= lastDataRowIndex; rowIndex += 1) {
        const isZebraRow = (rowIndex - detailStartRow) % 2 === 1;
        for (let colIndex = 0; colIndex < columnCount; colIndex += 1) {
          const styles = [baseDataStyle];
          if (isZebraRow) {
            styles.push(zebraDataStyle);
          }
          if (colIndex === 0) {
            styles.push(firstColumnStyle);
          }
          setCellStyle(sheet, rowIndex, colIndex, ...styles);
        }
        sheet['!rows'][rowIndex] = { ...(sheet['!rows'][rowIndex] ?? {}), hpt: 20 };
      }
    }

    const integerColumns = [1, 2, 5, 6, 7, 8, 9];
    const percentColumns = [3, 4];
    for (let rowIndex = dataStartRow; rowIndex <= lastDataRowIndex; rowIndex += 1) {
      for (const column of integerColumns) {
        setNumberFormat(sheet, rowIndex, column, '#,##0');
      }
      for (const column of percentColumns) {
        setNumberFormat(sheet, rowIndex, column, '0.0%');
      }
    }

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, 'Отчёт');
    const fileName = buildExportFileName(periods);
    XLSX.writeFile(workbook, fileName);
  } catch (error) {
    console.error('Не удалось выгрузить таблицу.', error);
    showPreviewMessage('Не удалось выгрузить таблицу. Повторите попытку.');
  }
}

// Готовлю массив значений для записи строки в Excel (проценты перевожу в доли).
function buildExportDataRow(row) {
  return [
    row.label ?? '',
    safeNumber(row.totalObjects),
    safeNumber(row.inspectedObjects),
    safeNumber(row.inspectedPercent) / 100,
    safeNumber(row.violationPercent) / 100,
    safeNumber(row.totalViolations),
    safeNumber(row.currentViolations),
    safeNumber(row.previousControl),
    safeNumber(row.resolved),
    safeNumber(row.onControl),
  ];
}

// Страхуюсь от NaN: Excel не любит нечисловые значения в числовых столбцах.
function safeNumber(value) {
  if (!Number.isFinite(value)) {
    return 0;
  }
  return Number(value);
}

// Накладываю на ячейку набор стилей, аккуратно объединяя их.
function setCellStyle(sheet, rowIndex, colIndex, ...styles) {
  const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const cell = sheet[address];
  if (!cell) {
    return;
  }
  cell.s = mergeStyles(cell.s ?? {}, ...styles);
}

// Назначаю числовой формат ячейке — пригодится для процентов и целых.
function setNumberFormat(sheet, rowIndex, colIndex, format) {
  const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const cell = sheet[address];
  if (!cell) {
    return;
  }
  cell.z = format;
}

// Аккуратно сливаю несколько объектов стилей в один.
function mergeStyles(...styles) {
  const result = {};
  for (const style of styles) {
    if (style) {
      deepMerge(result, style);
    }
  }
  return result;
}

function describeSelectedDataSource() {
  switch (state.dataSourceOption) {
    case 'oati':
      return 'Источник данных: ОАТИ';
    case 'cafap':
      return 'Источник данных: ЦАФАП';
    default:
      return 'Источник данных: ОАТИ и ЦАФАП';
  }
}

// Рекурсивное слияние объектов — минимальная реализация без внешних зависимостей.
function deepMerge(target, source) {
  for (const [key, value] of Object.entries(source)) {
    if (value && typeof value === 'object' && !Array.isArray(value)) {
      if (!target[key] || typeof target[key] !== 'object') {
        target[key] = {};
      }
      deepMerge(target[key], value);
    } else {
      target[key] = value;
    }
  }
  return target;
}

// Унифицированный стиль рамки, чтобы не дублировать одну и ту же структуру.
function buildBorderStyle(color) {
  return {
    top: { style: 'thin', color: { rgb: color } },
    bottom: { style: 'thin', color: { rgb: color } },
    left: { style: 'thin', color: { rgb: color } },
    right: { style: 'thin', color: { rgb: color } },
  };
}

// Красиво отображаю диапазон дат в описательной части Excel-отчёта.
function formatPeriodRange(period) {
  if (!period) {
    return '';
  }
  const start = formatDateDisplay(period.start);
  const end = formatDateDisplay(period.end);
  if (!start || !end) {
    return '';
  }
  return `${start} — ${end}`;
}

function formatTitlePeriod(period) {
  if (!period) {
    return '';
  }
  const sameYear =
    period.start instanceof Date &&
    period.end instanceof Date &&
    period.start.getFullYear() === period.end.getFullYear();
  const start = sameYear ? formatDateDisplayShort(period.start) : formatDateDisplay(period.start);
  const end = sameYear ? formatDateDisplayShort(period.end) : formatDateDisplay(period.end);
  if (!start || !end) {
    return '';
  }
  return `${start} — ${end}`;
}

// Подготавливаю дату для безопасного использования в имени файла.
function formatDateForFileName(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// Формирую итоговое имя файла и вычищаю недопустимые символы.
function buildExportFileName(periods) {
  const start = formatDateForFileName(periods.current.start);
  const end = formatDateForFileName(periods.current.end);
  const sanitized = `Отчёт ОАТИ ${start}_${end}.xlsx`.replace(/[\\/:*?"<>|]+/g, '_');
  return sanitized;
}
