// Ниже я собираю все служебные справочники, чтобы дальше по коду не гоняться за строками руками.
// Такой подход сразу же подсвечивает, какие статусы и сокращения мы ожидаем в входных данных.
const CONTROL_STATUSES = ['на устранении', 'на контроле инспектора оати'];
const CONTROL_STATUS_SET = new Set(CONTROL_STATUSES);
const CURRENT_CONTROL_STATUSES = [
  'на устранении',
  'на контроле инспектора оати',
  'новое',
  'ошибка передачи данных в цафап',
];
const CURRENT_CONTROL_STATUS_SET = new Set(CURRENT_CONTROL_STATUSES);
const RESOLVED_STATUS = 'снят с контроля';
const DRAFT_STATUS = 'черновик';
const INSPECTION_RESULT_VIOLATION = 'нарушение выявлено';

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
  'троицкий административный округ': 'ТиНАО',
  'троицкий и новомосковский административный округ': 'ТиНАО',
});

// Заодно храню список уже нормализованных сокращений, чтобы быстро понимать что в label у нас аббревиатура.
const MOSCOW_DISTRICT_ABBREVIATION_KEYS = new Set(
  Object.values(MOSCOW_DISTRICT_ABBREVIATIONS).map((value) => value.trim().toLowerCase()),
);

// На случай, когда в данных вообще нет округа, завожу понятный ключ.
const DISTRICT_FALLBACK_KEY = 'без округа';

const TINAO_DISPLAY_LABEL = 'ТиНАО';
const TINAO_AGGREGATED_KEY = 'тинао';
const TINAO_COMBINATION_KEYS = new Set([TINAO_AGGREGATED_KEY, 'нао', 'тао']);

const MAX_BULK_EXPORT_RULES = 20;
const BULK_EXPORT_STORAGE_KEY = 'oati:bulk-export-rules';
const BULK_EXPORT_PERIODS_STORAGE_KEY = 'oati:bulk-export-periods';
const BULK_EXPORT_TEMPLATES_STORAGE_KEY = 'oati:bulk-export-templates';
const SYSTEM_BULK_EXPORT_TEMPLATES = Object.freeze(createSystemBulkExportTemplates());

const DISTRICT_SORT_ORDER = new Map([
  ['итого', 0],
  ['цао', 1],
  ['сао', 2],
  ['свао', 3],
  ['вао', 4],
  ['ювао', 5],
  ['юао', 6],
  ['юзао', 7],
  ['зао', 8],
  ['сзао', 9],
  ['зелао', 10],
  [TINAO_AGGREGATED_KEY, 11],
  ['без округа', 12],
]);

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
  {
    key: 'objectId',
    label: 'ID объекта',
    candidates: [
      'id объекта',
      'id объекта контроля',
      'id объекта (uuid)',
      'ид объекта',
      'uuid объекта',
    ],
    optional: true,
  },
  { key: 'inspectionDate', label: 'Дата обследования', candidates: ['дата обследования', 'дата осмотра', 'дата контроля'] },
  { key: 'district', label: 'Округ', candidates: ['округ', 'административный округ', 'округ объекта'] },
  { key: 'dataSource', label: 'Источник данных', candidates: ['источник данных'], optional: true },
];

// Аналогичный список для справочника объектов.
const objectFieldDefinitions = [
  {
    key: 'objectId',
    label: 'ID объекта',
    candidates: ['id объекта', 'id объекта контроля', 'id', 'ид объекта', 'идентификатор объекта'],
  },
  {
    key: 'externalObjectId',
    label: 'Внешний идентификатор объекта',
    candidates: [
      'внешний идентификатор объекта',
      'внешний id объекта',
      'внешний идентификатор',
      'внешний id',
      'external object id',
      'external id',
    ],
  },
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
  combineTiNaoDistricts: false,
  bulkExportRules: [],
  bulkExportPeriods: {
    currentStart: '',
    currentEnd: '',
    previousStart: '',
    previousEnd: '',
  },
  bulkExportSearch: {},
  bulkExportTemplates: {
    system: SYSTEM_BULK_EXPORT_TEMPLATES,
    custom: [],
    appliedTemplateId: null,
  },
};

function resolveReportContext(overrides = {}) {
  return {
    typeMode: overrides.typeMode ?? state.typeMode,
    customTypes: Array.isArray(overrides.customTypes)
      ? [...overrides.customTypes]
      : [...state.customTypes],
    violationMode: overrides.violationMode ?? state.violationMode,
    customViolations: Array.isArray(overrides.customViolations)
      ? [...overrides.customViolations]
      : [...state.customViolations],
    dataSourceOption: overrides.dataSourceOption ?? state.dataSourceOption,
    combineTiNaoDistricts:
      typeof overrides.combineTiNaoDistricts === 'boolean'
        ? overrides.combineTiNaoDistricts
        : state.combineTiNaoDistricts,
  };
}

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
  combineTiNaoToggle: document.getElementById('combine-tinao'),
  previewMessage: document.getElementById('preview-message'),
  reportTable: document.getElementById('report-table'),
  refreshButton: document.getElementById('refresh-report'),
  downloadButton: document.getElementById('download-report'),
  openBulkExportButton: document.getElementById('open-bulk-export'),
  themeToggle: document.getElementById('theme-toggle'),
  themeToggleLabel: document.getElementById('theme-toggle-text'),
  brandLogo: document.querySelector('.brand-logo'),
  easterEggMessage: document.getElementById('easter-egg-message'),
  creditBadge: document.querySelector('.credit-badge'),
  creditBadgeClose: document.querySelector('.credit-badge__close'),
  bulkExportOverlay: document.getElementById('bulk-export-overlay'),
  bulkExportPanel: document.getElementById('bulk-export-panel'),
  closeBulkExportButton: document.getElementById('close-bulk-export'),
  bulkExportTemplatesContainer: document.getElementById('bulk-export-templates'),
  bulkExportPeriodsContainer: document.getElementById('bulk-export-periods'),
  bulkExportPeriodsMessage: document.getElementById('bulk-export-periods-message'),
  bulkExportMessage: document.getElementById('bulk-export-message'),
  bulkExportRulesContainer: document.getElementById('bulk-export-rules'),
  addBulkExportRuleButton: document.getElementById('add-bulk-export-rule'),
  saveBulkExportTemplateButton: document.getElementById('save-bulk-export-template'),
  runBulkExportButton: document.getElementById('run-bulk-export'),
};

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

/**
 * Датапикер, работающий только с доступными датами из исходных файлов.
 * Компонент отвечает за отображение календаря, навигацию по месяцам
 * и выбор значения с последующим уведомлением слушателей.
 */
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

  /**
   * Обновляет список доступных дат и перестраивает календарь.
   * @param {Array<{iso: string, date: Date}>} dates — массив дат в формате ISO и объектов Date.
   */
  setAvailableDates(dates) {
    const normalizedDates = [];
    const seenMonths = new Map();
    if (Array.isArray(dates)) {
      for (const item of dates) {
        if (!item || typeof item.iso !== 'string') {
          continue;
        }
        const iso = item.iso.trim();
        if (!iso) {
          continue;
        }
        let baseDate = item.date instanceof Date ? new Date(item.date.getTime()) : null;
        if (!isValidDate(baseDate)) {
          const parsed = parseIsoDate(iso);
          if (!parsed) {
            continue;
          }
          baseDate = parsed;
        }
        const normalizedIso = formatIsoDate(baseDate);
        if (!normalizedIso) {
          continue;
        }
        normalizedDates.push({ iso: normalizedIso, date: baseDate });
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

  /**
   * Устанавливает выбранную дату, валидируя её наличие в списке доступных значений.
   * @param {string} value — ISO-строка даты (YYYY-MM-DD).
   * @param {{emitEvent?: boolean}} [options] — настройки уведомлений об изменении значения.
   */
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

  /**
   * Перерисовывает значение на кнопке выбора даты и обновляет атрибуты доступности.
   */
  updateDisplay() {
    if (!this.valueNode) {
      return;
    }
    const value = this.getValue();
    if (value) {
      const parsed = parseIsoDate(value);
      if (parsed) {
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

  /**
   * Открывает календарь, создаёт слушатели кликов и клавиш для закрытия.
   */
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

  /**
   * Закрывает календарь и удаляет вспомогательные обработчики.
   */
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
    if (!parsed) {
      return;
    }
    const key = `${parsed.getFullYear()}-${parsed.getMonth()}`;
    const index = this.availableMonths.findIndex((item) => `${item.year}-${item.month}` === key);
    if (index >= 0) {
      this.monthIndex = index;
    }
  }

  /**
   * Строит DOM календаря за выбранный месяц и помечает доступные дни.
   */
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
      if (!iso) {
        const dayCell = document.createElement('span');
        dayCell.className = 'date-picker__day is-disabled';
        dayCell.textContent = String(day);
        grid.append(dayCell);
        continue;
      }
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

  /**
   * Устанавливает фокус на выбранную дату или первую доступную ячейку календаря.
   */
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
  return 'dark';
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

// Форматер чисел, чтобы везде были привычные для отчётов пробелы.
const numberFormatter = new Intl.NumberFormat('ru-RU');

// Обработчики файлов
if (elements.violationsInput) {
  // Когда пользователь подкидывает таблицу нарушений — читаем файл и обновляем состояние.
  elements.violationsInput.addEventListener('change', async (event) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    setLoadingIndicator('violations', true, file.name);
    try {
      await loadDataset('violations', file);
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
    setLoadingIndicator('objects', true, file.name);
    try {
      await loadDataset('objects', file);
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

if (elements.openBulkExportButton) {
  elements.openBulkExportButton.addEventListener('click', () => {
    openBulkExportOverlay();
  });
}

if (elements.closeBulkExportButton) {
  elements.closeBulkExportButton.addEventListener('click', (event) => {
    event.preventDefault();
    closeBulkExportOverlay();
  });
}

if (elements.bulkExportOverlay) {
  elements.bulkExportOverlay.addEventListener('click', (event) => {
    if (event.target === elements.bulkExportOverlay) {
      closeBulkExportOverlay();
    }
  });
}

if (elements.addBulkExportRuleButton) {
  elements.addBulkExportRuleButton.addEventListener('click', () => {
    addBulkExportRule();
  });
}

if (elements.saveBulkExportTemplateButton) {
  elements.saveBulkExportTemplateButton.addEventListener('click', () => {
    handleSaveBulkExportTemplate();
  });
}

if (elements.runBulkExportButton) {
  elements.runBulkExportButton.addEventListener('click', () => {
    handleBulkExportDownload();
  });
}

document.addEventListener('keydown', handleBulkOverlayKeydown);

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

if (elements.combineTiNaoToggle) {
  elements.combineTiNaoToggle.addEventListener('change', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }
    state.combineTiNaoDistricts = target.checked;
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
 * Загружает Excel-файл, извлекает записи и обновляет состояние приложения.
 * После считывания автоматически настраивает соответствия колонок, доступные фильтры
 * и планирует пересчёт отчёта.
 *
 * @param {'violations' | 'objects'} kind Тип загружаемого набора данных.
 * @param {File} file Объект файла, выбранный пользователем.
 * @returns {Promise<void>} Промис завершается после обновления интерфейса.
 */
async function loadDataset(kind, file) {
  try {
    // Сообщаю пользователю, что файл читается.
    showPreviewMessage(`Загрузка файла «${file.name}»...`);
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

function setLoadingIndicator(kind, isLoading, fileName) {
  const loaderMap = {
    violations: elements.violationsLoader,
    objects: elements.objectsLoader,
  };
  const loader = loaderMap[kind];
  if (!loader) {
    return;
  }
  if (isLoading) {
    const textNode = loader.querySelector('.file-loader__text');
    if (textNode) {
      textNode.textContent = fileName ? `Обрабатываем «${fileName}»` : 'Обрабатываем файл';
    }
    loader.hidden = false;
  } else {
    loader.hidden = true;
  }
}

/**
 * Читает Excel-файл и подбирает строку заголовков для последующей агрегации данных.
 * Используется для обоих наборов данных — нарушений и объектов.
 *
 * @param {File} file Файл Excel, загруженный пользователем.
 * @param {string[]} headerKeywords Перечень ключевых слов, помогающих определить строку заголовков.
 * @returns {Promise<{records: object[], headers: string[]}>} Набор записей и найденные заголовки.
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
  const records = extractRecords(sheet, headerRowIndex, headers);
  return { records, headers };
}

/**
 * Определяет индекс строки, наиболее похожей на строку заголовков.
 * Строка выбирается по количеству текстовых значений и совпадениям с ожидаемыми именами столбцов.
 *
 * @param {XLSX.WorkSheet} sheet Активный лист книги Excel.
 * @param {XLSX.Range} range Рабочий диапазон листа.
 * @param {string[]} headerKeywords Нормализованные ключевые слова для заголовков.
 * @returns {number | null} Индекс строки заголовков либо null, если строка не найдена.
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
 * Унифицирует заголовок для сравнения: убирает лишние пробелы и приводит к нижнему регистру.
 *
 * @param {unknown} value Значение ячейки заголовка.
 * @returns {string} Нормализованная строка.
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
 * Извлекает заголовки столбцов из выбранной строки.
 *
 * @param {XLSX.WorkSheet} sheet Активный лист книги Excel.
 * @param {number} rowIndex Индекс строки заголовков.
 * @param {XLSX.Range} range Рабочий диапазон листа.
 * @returns {string[]} Массив подписей столбцов.
 */
function extractHeaderRow(sheet, rowIndex, range) {
  const headers = [];
  for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex += 1) {
    const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
    const cell = sheet[address];
    headers.push(typeof cell?.v === 'string' ? cell.v.trim() : cell?.v ?? `Колонка ${colIndex + 1}`);
  }
  return headers;
}

/**
 * Собирает записи на основе найденных заголовков, пропуская пустые строки.
 *
 * @param {XLSX.WorkSheet} sheet Активный лист книги Excel.
 * @param {number} headerRowIndex Индекс строки заголовков.
 * @param {string[]} headers Список заголовков.
 * @returns {object[]} Массив объектов со значениями строк.
 */
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

/**
 * Преобразует значение ячейки в читаемый формат, сохраняя числовые и булевы типы.
 *
 * @param {XLSX.CellObject | undefined} cell Ячейка таблицы.
 * @returns {string | number | boolean | Date | ''} Подготовленное значение.
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

/**
 * Автоматически сопоставляет ключевые поля с колонками исходной таблицы.
 *
 * @param {string[]} columns Названия колонок, найденные в файле.
 * @param {Array<{key: string, candidates: string[]}>} definitions Описание полей и допустимых названий.
 * @returns {Record<string, string>} Карта соответствий ключей и названий колонок.
 */
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

/**
 * Строит секцию с настройками сопоставления колонок для загруженных таблиц.
 */
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

/**
 * Создаёт DOM-элементы select для выбора колонок под каждое требуемое поле.
 *
 * @param {HTMLElement} container Контейнер для вывода списка сопоставлений.
 * @param {string[]} columns Массив названий колонок.
 * @param {Array<{key: string, label: string}>} definitions Схема ожидаемых полей.
 * @param {Record<string, string>} mapping Текущее сопоставление ключей и колонок.
 * @param {(key: string, value: string) => void} onChange Колбэк при изменении выбранной колонки.
 */
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

/**
 * Управляет видимостью основных секций интерфейса на основе состояния данных.
 */
function updateVisibility() {
  const ready = state.violations.length && state.objects.length;
  elements.mappingSection.hidden = !(state.violations.length || state.objects.length);
  elements.controlsSection.hidden = !ready;
  elements.previewSection.hidden = !ready;
  if (!ready) {
    clearTable();
  }
}

/**
 * Обновляет доступные фильтры (даты, типы, нарушения) в соответствии с загруженными данными.
 */
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
  syncBulkExportRulesWithAvailableOptions();
}

/**
 * Обновляет конкретный датапикер новым набором доступных дат.
 *
 * @param {AvailabilityDatePicker | undefined} picker Экземпляр датапикера.
 * @param {Array<{iso: string, date: Date}>} dates Список доступных дат.
 * @param {number | null} defaultIndex Индекс даты по умолчанию.
 */
function updateDatePicker(picker, dates, defaultIndex) {
  if (!picker) {
    return;
  }
  const availableValues = new Set(dates.map((item) => item.iso).filter(Boolean));
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

/**
 * Отрисовывает компонент с тегами, поиском и выпадающим списком значений.
 * Используется для выбора типов объектов и наименований нарушений.
 *
 * @param {Object} options Настройки компонента мультiselect.
 * @param {HTMLElement} options.container Контейнер для вывода элемента.
 * @param {string[]} options.values Доступные значения.
 * @param {string[]} options.selectedValues Уже выбранные значения.
 * @param {boolean} options.disabled Флаг блокировки компонента.
 * @param {string} options.placeholder Подсказка внутри поля поиска.
 * @param {string} options.emptyLabel Сообщение при отсутствии значений.
 * @param {number} options.limit Максимальное количество значений.
 * @param {string} options.limitMessage Сообщение при превышении лимита.
 * @param {HTMLElement} [options.messageElement] Узел для вывода подсказок.
 * @param {(value: string) => void} [options.onAdd] Колбэк при добавлении значения.
 * @param {(value: string) => void} [options.onRemove] Колбэк при удалении значения.
 */
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

/**
 * Отрисовывает интерфейс выбора типов объектов контроля.
 */
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

/**
 * Отрисовывает интерфейс выбора наименований нарушений.
 */
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

/**
 * Управляет состоянием фильтра по источнику данных и вспомогательным сообщением.
 */
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

/**
 * Анализирует колонку источника данных и определяет, какие категории присутствуют.
 *
 * @param {string} column Название колонки с источником данных.
 * @returns {{oati: boolean, cafap: boolean, both: boolean, unknown: boolean}} Сводка по категориям источников.
 */
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

function generateBulkRuleId() {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return `rule-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
}

function generateBulkTemplateId() {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return `template-${Date.now()}-${Math.random().toString(16).slice(2, 8)}`;
}

function createSystemBulkExportTemplates() {
  const roadType = 'ВК_Объекты дорожного хозяйства';
  const yardType = 'ВК_Дворовые территории';
  return [
    {
      id: 'system-vk-road',
      name: roadType,
      description:
        'Набор для массовых выгрузок по объектам дорожного хозяйства: включает общее покрытие и акцент на ямах, бордюрном камне и уборке.',
      rules: [
        {
          name: 'Правило 1. Все нарушения — ОАТИ и ЦАФАП',
          typeMode: 'custom',
          customTypes: [roadType],
          violationMode: 'all',
          customViolations: [],
          dataSourceOption: 'all',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 2. Все нарушения — только ОАТИ',
          typeMode: 'custom',
          customTypes: [roadType],
          violationMode: 'all',
          customViolations: [],
          dataSourceOption: 'oati',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 3. Ямы и выбоины — ОАТИ и ЦАФАП',
          typeMode: 'custom',
          customTypes: [roadType],
          violationMode: 'custom',
          customViolations: ['Ямы и выбоины в асфальтобетонном покрытии'],
          dataSourceOption: 'all',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 4. Ямы и выбоины — только ОАТИ',
          typeMode: 'custom',
          customTypes: [roadType],
          violationMode: 'custom',
          customViolations: ['Ямы и выбоины в асфальтобетонном покрытии'],
          dataSourceOption: 'oati',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 5. Бортовой камень — ОАТИ и ЦАФАП',
          typeMode: 'custom',
          customTypes: [roadType],
          violationMode: 'custom',
          customViolations: ['Отсутствует/повреждён бортовой камень'],
          dataSourceOption: 'all',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 6. Бортовой камень — только ОАТИ',
          typeMode: 'custom',
          customTypes: [roadType],
          violationMode: 'custom',
          customViolations: ['Отсутствует/повреждён бортовой камень'],
          dataSourceOption: 'oati',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 7. Неубранные объекты — ОАТИ и ЦАФАП',
          typeMode: 'custom',
          customTypes: [roadType],
          violationMode: 'custom',
          customViolations: [
            'Не убран объект дорожного хозяйства',
            'Не убран пешеходный переход',
            'Не убрана дворовая территория',
          ],
          dataSourceOption: 'all',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 8. Неубранные объекты — только ОАТИ',
          typeMode: 'custom',
          customTypes: [roadType],
          violationMode: 'custom',
          customViolations: [
            'Не убран объект дорожного хозяйства',
            'Не убран пешеходный переход',
            'Не убрана дворовая территория',
          ],
          dataSourceOption: 'oati',
          combineTiNaoDistricts: true,
        },
      ],
    },
    {
      id: 'system-vk-yard',
      name: yardType,
      description:
        'Стандарт для дворовых территорий: сочетает общие показатели, работы по бортовому камню, ямочному ремонту и состоянию МАФов.',
      rules: [
        {
          name: 'Правило 1. Все нарушения — ОАТИ и ЦАФАП',
          typeMode: 'custom',
          customTypes: [yardType],
          violationMode: 'all',
          customViolations: [],
          dataSourceOption: 'all',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 2. Все нарушения — только ОАТИ',
          typeMode: 'custom',
          customTypes: [yardType],
          violationMode: 'all',
          customViolations: [],
          dataSourceOption: 'oati',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 3. Бортовой камень — ОАТИ и ЦАФАП',
          typeMode: 'custom',
          customTypes: [yardType],
          violationMode: 'custom',
          customViolations: ['Отсутствует/повреждён бортовой камень'],
          dataSourceOption: 'all',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 4. Бортовой камень — только ОАТИ',
          typeMode: 'custom',
          customTypes: [yardType],
          violationMode: 'custom',
          customViolations: ['Отсутствует/повреждён бортовой камень'],
          dataSourceOption: 'oati',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 5. Ямы и выбоины — ОАТИ и ЦАФАП',
          typeMode: 'custom',
          customTypes: [yardType],
          violationMode: 'custom',
          customViolations: ['Ямы и выбоины в асфальтобетонном покрытии'],
          dataSourceOption: 'all',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 6. Ямы и выбоины — только ОАТИ',
          typeMode: 'custom',
          customTypes: [yardType],
          violationMode: 'custom',
          customViolations: ['Ямы и выбоины в асфальтобетонном покрытии'],
          dataSourceOption: 'oati',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 7. Неубранные объекты и МАФы — ОАТИ и ЦАФАП',
          typeMode: 'custom',
          customTypes: [yardType],
          violationMode: 'custom',
          customViolations: [
            'Не убран объект дорожного хозяйства',
            'Не восстановлен демонтированный МАФ',
            'Не убрана дворовая территория',
          ],
          dataSourceOption: 'all',
          combineTiNaoDistricts: true,
        },
        {
          name: 'Правило 8. Состояние МАФов — только ОАТИ',
          typeMode: 'custom',
          customTypes: [yardType],
          violationMode: 'custom',
          customViolations: [
            'Грязный МАФ',
            'Не убран пешеходный переход',
            'Не окрашен МАФ',
            'Сломан МАФ',
            'Сломан МАФ (нарушение норм безопасности)',
          ],
          dataSourceOption: 'oati',
          combineTiNaoDistricts: true,
        },
      ],
    },
  ];
}

function normalizeBulkExportPeriods(periods) {
  return {
    currentStart: typeof periods?.currentStart === 'string' ? periods.currentStart.trim() : '',
    currentEnd: typeof periods?.currentEnd === 'string' ? periods.currentEnd.trim() : '',
    previousStart: typeof periods?.previousStart === 'string' ? periods.previousStart.trim() : '',
    previousEnd: typeof periods?.previousEnd === 'string' ? periods.previousEnd.trim() : '',
  };
}

function normalizeBulkTemplateRule(rule, index = 0) {
  const typeMode = rule?.typeMode === 'custom' ? 'custom' : 'all';
  const violationMode = rule?.violationMode === 'custom' ? 'custom' : 'all';
  const normalized = {
    name: typeof rule?.name === 'string' ? rule.name.trim() : '',
    typeMode,
    customTypes: typeMode === 'custom' ? normalizeStringArray(rule?.customTypes ?? [], 3) : [],
    violationMode,
    customViolations:
      violationMode === 'custom' ? normalizeStringArray(rule?.customViolations ?? [], 5) : [],
    dataSourceOption: ['oati', 'cafap', 'all'].includes(rule?.dataSourceOption)
      ? rule.dataSourceOption
      : 'all',
    combineTiNaoDistricts: Boolean(rule?.combineTiNaoDistricts),
  };
  if (!normalized.name) {
    normalized.name = `Правило ${index + 1}`;
  }
  return normalized;
}

function normalizeBulkExportTemplate(template) {
  const id = typeof template?.id === 'string' && template.id.trim() ? template.id.trim() : generateBulkTemplateId();
  const name = typeof template?.name === 'string' ? template.name.trim() : '';
  const description = typeof template?.description === 'string' ? template.description.trim() : '';
  const rulesSource = Array.isArray(template?.rules) ? template.rules : [];
  const rules = rulesSource.slice(0, MAX_BULK_EXPORT_RULES).map((rule, index) => ({
    ...normalizeBulkTemplateRule(rule, index),
  }));
  return { id, name, description, rules };
}

function loadBulkExportRulesFromStorage() {
  if (typeof window === 'undefined' || !window.localStorage) {
    return [];
  }
  try {
    const raw = window.localStorage.getItem(BULK_EXPORT_STORAGE_KEY);
    if (!raw) {
      return [];
    }
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }
    return parsed.slice(0, MAX_BULK_EXPORT_RULES).map((item) => normalizeBulkExportRule(item));
  } catch (error) {
    console.warn('Не удалось прочитать сохранённые правила массовой выгрузки.', error);
    return [];
  }
}

function loadCustomBulkExportTemplatesFromStorage() {
  if (typeof window === 'undefined' || !window.localStorage) {
    return [];
  }
  try {
    const raw = window.localStorage.getItem(BULK_EXPORT_TEMPLATES_STORAGE_KEY);
    if (!raw) {
      return [];
    }
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }
    return parsed.map((item, index) => normalizeBulkExportTemplate({ ...item, id: item?.id ?? `custom-${index}` }));
  } catch (error) {
    console.warn('Не удалось прочитать сохранённые шаблоны массовой выгрузки.', error);
    return [];
  }
}

function loadBulkExportPeriodsFromStorage() {
  if (typeof window === 'undefined' || !window.localStorage) {
    return normalizeBulkExportPeriods(getCurrentPeriodInputs());
  }
  try {
    const raw = window.localStorage.getItem(BULK_EXPORT_PERIODS_STORAGE_KEY);
    if (!raw) {
      return normalizeBulkExportPeriods(getCurrentPeriodInputs());
    }
    const parsed = JSON.parse(raw);
    return normalizeBulkExportPeriods(parsed);
  } catch (error) {
    console.warn('Не удалось прочитать сохранённые периоды массовой выгрузки.', error);
    return normalizeBulkExportPeriods(getCurrentPeriodInputs());
  }
}

function saveBulkExportRulesToStorage() {
  if (typeof window === 'undefined' || !window.localStorage) {
    return;
  }
  try {
    const payload = JSON.stringify(state.bulkExportRules.map((rule) => normalizeBulkExportRule(rule)));
    window.localStorage.setItem(BULK_EXPORT_STORAGE_KEY, payload);
  } catch (error) {
    console.warn('Не удалось сохранить правила массовой выгрузки.', error);
  }
}

function saveBulkExportPeriodsToStorage() {
  if (typeof window === 'undefined' || !window.localStorage) {
    return;
  }
  try {
    const payload = JSON.stringify(normalizeBulkExportPeriods(state.bulkExportPeriods));
    window.localStorage.setItem(BULK_EXPORT_PERIODS_STORAGE_KEY, payload);
  } catch (error) {
    console.warn('Не удалось сохранить периоды массовой выгрузки.', error);
  }
}

function saveCustomBulkExportTemplatesToStorage() {
  if (typeof window === 'undefined' || !window.localStorage) {
    return;
  }
  try {
    const payload = JSON.stringify(state.bulkExportTemplates.custom.map((template) => normalizeBulkExportTemplate(template)));
    window.localStorage.setItem(BULK_EXPORT_TEMPLATES_STORAGE_KEY, payload);
  } catch (error) {
    console.warn('Не удалось сохранить шаблоны массовой выгрузки.', error);
  }
}

function normalizeStringArray(values, limit) {
  if (!Array.isArray(values)) {
    return [];
  }
  const result = [];
  const seen = new Set();
  for (const value of values) {
    if (result.length >= limit) {
      break;
    }
    const text = typeof value === 'string' ? value.trim() : '';
    if (!text) {
      continue;
    }
    const key = normalizeKey(text);
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    result.push(text);
  }
  return result;
}

function normalizeBulkExportRule(rule) {
  const periods = rule?.periods ?? {};
  const currentStart = typeof periods.currentStart === 'string' ? periods.currentStart.trim() : '';
  const currentEnd = typeof periods.currentEnd === 'string' ? periods.currentEnd.trim() : '';
  const previousStart = typeof periods.previousStart === 'string' ? periods.previousStart.trim() : '';
  const previousEnd = typeof periods.previousEnd === 'string' ? periods.previousEnd.trim() : '';
  const typeMode = rule?.typeMode === 'custom' ? 'custom' : 'all';
  const violationMode = rule?.violationMode === 'custom' ? 'custom' : 'all';
  const dataSourceOption = ['oati', 'cafap', 'all'].includes(rule?.dataSourceOption)
    ? rule.dataSourceOption
    : 'all';
  return {
    id: typeof rule?.id === 'string' && rule.id ? rule.id : generateBulkRuleId(),
    name: typeof rule?.name === 'string' ? rule.name.trim() : '',
    typeMode,
    customTypes: typeMode === 'custom' ? normalizeStringArray(rule?.customTypes ?? [], 3) : [],
    violationMode,
    customViolations:
      violationMode === 'custom' ? normalizeStringArray(rule?.customViolations ?? [], 5) : [],
    dataSourceOption,
    combineTiNaoDistricts: Boolean(rule?.combineTiNaoDistricts),
    periods: {
      currentStart,
      currentEnd,
      previousStart,
      previousEnd,
    },
  };
}

function serializeRuleToTemplate(rule, index) {
  const normalized = normalizeBulkExportRule(rule);
  return normalizeBulkTemplateRule(
    {
      name: normalized.name,
      typeMode: normalized.typeMode,
      customTypes: normalized.customTypes,
      violationMode: normalized.violationMode,
      customViolations: normalized.customViolations,
      dataSourceOption: normalized.dataSourceOption,
      combineTiNaoDistricts: normalized.combineTiNaoDistricts,
    },
    index,
  );
}

function getCurrentPeriodInputs() {
  return {
    currentStart: elements.currentStart?.value ?? '',
    currentEnd: elements.currentEnd?.value ?? '',
    previousStart: elements.previousStart?.value ?? '',
    previousEnd: elements.previousEnd?.value ?? '',
  };
}

function createDefaultBulkExportRule() {
  const context = resolveReportContext();
  const periods = normalizeBulkExportPeriods(state.bulkExportPeriods);
  return normalizeBulkExportRule({
    id: generateBulkRuleId(),
    name: `Правило ${state.bulkExportRules.length + 1}`,
    typeMode: context.typeMode,
    customTypes: context.typeMode === 'custom' ? context.customTypes : [],
    violationMode: context.violationMode,
    customViolations: context.violationMode === 'custom' ? context.customViolations : [],
    dataSourceOption: context.dataSourceOption,
    combineTiNaoDistricts: context.combineTiNaoDistricts,
    periods,
  });
}

function addBulkExportRule() {
  if (state.bulkExportRules.length >= MAX_BULK_EXPORT_RULES) {
    setBulkExportMessage(`Можно создать не более ${MAX_BULK_EXPORT_RULES} правил.`);
    return;
  }
  clearAppliedBulkTemplateTracking();
  const rule = createDefaultBulkExportRule();
  ensureBulkExportSearchState(rule.id);
  state.bulkExportRules = [...state.bulkExportRules, rule];
  saveBulkExportRulesToStorage();
  renderBulkExportRules();
  setBulkExportMessage('Правило добавлено. Настройте параметры перед выгрузкой.');
}

function deleteBulkExportRule(ruleId) {
  const next = state.bulkExportRules.filter((rule) => rule.id !== ruleId);
  if (next.length === state.bulkExportRules.length) {
    return;
  }
  clearAppliedBulkTemplateTracking();
  state.bulkExportRules = next;
  delete state.bulkExportSearch[ruleId];
  saveBulkExportRulesToStorage();
  renderBulkExportRules();
  setBulkExportMessage('Правило удалено.');
}

function updateBulkExportRule(ruleId, updater, options = {}) {
  const { render = true } = options;
  const index = state.bulkExportRules.findIndex((rule) => rule.id === ruleId);
  if (index === -1) {
    return;
  }
  const current = state.bulkExportRules[index];
  const nextValue = typeof updater === 'function' ? updater({ ...current }) : { ...current, ...updater };
  if (!nextValue) {
    return;
  }
  const normalized = normalizeBulkExportRule({ ...current, ...nextValue, id: current.id });
  const hasChanged = JSON.stringify(current) !== JSON.stringify(normalized);
  if (!hasChanged) {
    return;
  }
  clearAppliedBulkTemplateTracking();
  state.bulkExportRules = [
    ...state.bulkExportRules.slice(0, index),
    normalized,
    ...state.bulkExportRules.slice(index + 1),
  ];
  saveBulkExportRulesToStorage();
  if (render) {
    renderBulkExportRules();
  }
}

function setBulkExportMessage(message) {
  if (elements.bulkExportMessage) {
    elements.bulkExportMessage.textContent = message ?? '';
  }
}

function renderBulkExportRules() {
  const container = elements.bulkExportRulesContainer;
  if (!container) {
    return;
  }
  container.innerHTML = '';
  if (!state.bulkExportRules.length) {
    const empty = document.createElement('div');
    empty.className = 'bulk-export-empty';
    empty.textContent = 'Добавьте правило, чтобы настроить массовую выгрузку.';
    container.append(empty);
  } else {
    state.bulkExportRules.forEach((rule, index) => {
      ensureBulkExportSearchState(rule.id);
      container.append(createBulkExportRuleElement(rule, index));
    });
  }
  if (elements.addBulkExportRuleButton) {
    elements.addBulkExportRuleButton.disabled = state.bulkExportRules.length >= MAX_BULK_EXPORT_RULES;
  }
  if (elements.runBulkExportButton) {
    elements.runBulkExportButton.disabled = state.bulkExportRules.length === 0;
  }
  updateBulkExportTemplateControls();
}

function renderBulkExportTemplates() {
  const container = elements.bulkExportTemplatesContainer;
  if (!container) {
    return;
  }
  container.innerHTML = '';
  const fieldset = document.createElement('fieldset');
  fieldset.className = 'bulk-export-fieldset bulk-export-fieldset--templates';
  const legend = document.createElement('legend');
  legend.textContent = 'Шаблоны массовой выгрузки';
  fieldset.append(legend);

  const description = document.createElement('p');
  description.className = 'bulk-export-note';
  description.textContent =
    'Выберите готовый набор фильтров или сохраните собственный шаблон, чтобы делиться настройками с коллегами.';
  fieldset.append(description);

  const groups = [
    {
      title: 'Системные шаблоны',
      templates: state.bulkExportTemplates.system,
      emptyText: 'Системные шаблоны временно недоступны.',
      kind: 'system',
    },
    {
      title: 'Мои шаблоны',
      templates: state.bulkExportTemplates.custom,
      emptyText: 'Сохраните собственный шаблон, чтобы использовать его повторно.',
      kind: 'custom',
    },
  ];

  groups.forEach((group) => {
    const groupWrapper = document.createElement('section');
    groupWrapper.className = 'bulk-export-template-group';

    const title = document.createElement('h3');
    title.className = 'bulk-export-template-group__title';
    title.textContent = group.title;
    groupWrapper.append(title);

    if (!group.templates.length) {
      const empty = document.createElement('p');
      empty.className = 'bulk-export-template-empty';
      empty.textContent = group.emptyText;
      groupWrapper.append(empty);
    } else {
      const list = document.createElement('div');
      list.className = 'bulk-export-template-list';
      group.templates.forEach((template) => {
        list.append(createBulkExportTemplateCard(template, group.kind));
      });
      groupWrapper.append(list);
    }

    fieldset.append(groupWrapper);
  });

  container.append(fieldset);
}

function createBulkExportTemplateCard(template, kind) {
  const card = document.createElement('article');
  card.className = 'bulk-export-template-card';
  if (state.bulkExportTemplates.appliedTemplateId === template.id) {
    card.classList.add('bulk-export-template-card--active');
  }

  const header = document.createElement('div');
  header.className = 'bulk-export-template-card__header';

  const title = document.createElement('h4');
  title.className = 'bulk-export-template-card__title';
  title.textContent = template.name || 'Без названия';
  header.append(title);

  if (kind === 'system') {
    const badge = document.createElement('span');
    badge.className = 'bulk-export-template-card__badge';
    badge.textContent = 'Системный шаблон';
    header.append(badge);
  }

  card.append(header);

  const summary = document.createElement('p');
  summary.className = 'bulk-export-template-card__summary';
  const ruleCount = template.rules.length;
  summary.textContent =
    template.description && template.description.trim().length
      ? template.description
      : `Содержит ${ruleCount} ${formatRuPlural(ruleCount, 'правило', 'правила', 'правил')}.`;
  card.append(summary);

  if (template.rules.length) {
    const rulesList = document.createElement('ul');
    rulesList.className = 'bulk-export-template-card__rules';
    template.rules.forEach((rule, index) => {
      const item = document.createElement('li');
      item.className = 'bulk-export-template-card__rule';
      const name = document.createElement('span');
      name.className = 'bulk-export-template-card__rule-name';
      name.textContent = rule.name || `Правило ${index + 1}`;
      const details = document.createElement('p');
      details.className = 'bulk-export-template-card__rule-summary';
      details.textContent = formatBulkTemplateRuleSummary(rule);
      item.append(name, details);
      rulesList.append(item);
    });
    card.append(rulesList);
  }

  const actions = document.createElement('div');
  actions.className = 'bulk-export-template-card__actions';

  const applyButton = document.createElement('button');
  applyButton.type = 'button';
  applyButton.className = 'bulk-export-template-card__apply';
  const isActive = state.bulkExportTemplates.appliedTemplateId === template.id;
  applyButton.textContent = isActive ? 'Применён' : 'Применить';
  applyButton.disabled = isActive;
  applyButton.addEventListener('click', () => {
    applyBulkExportTemplate(template.id);
  });
  actions.append(applyButton);

  if (kind === 'custom') {
    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'bulk-export-template-card__delete';
    deleteButton.textContent = 'Удалить';
    deleteButton.addEventListener('click', () => {
      deleteCustomBulkExportTemplate(template.id);
    });
    actions.append(deleteButton);
  }

  card.append(actions);

  return card;
}

function formatBulkTemplateRuleSummary(rule) {
  const typePart =
    rule.typeMode === 'custom'
      ? `Типы: ${rule.customTypes.join(', ')}`
      : 'Типы: все';
  const violationPart =
    rule.violationMode === 'custom'
      ? `Нарушения: ${rule.customViolations.join(', ')}`
      : 'Нарушения: все';
  let sourcePart = 'Источник: ОАТИ и ЦАФАП';
  if (rule.dataSourceOption === 'oati') {
    sourcePart = 'Источник: только ОАТИ';
  } else if (rule.dataSourceOption === 'cafap') {
    sourcePart = 'Источник: только ЦАФАП';
  }
  const tinaoPart = rule.combineTiNaoDistricts ? 'ТиНАО объединён' : 'ТиНАО отдельно';
  return `${typePart} • ${violationPart} • ${sourcePart} • ${tinaoPart}`;
}

function updateBulkExportTemplateControls() {
  if (elements.saveBulkExportTemplateButton) {
    elements.saveBulkExportTemplateButton.disabled = state.bulkExportRules.length === 0;
  }
}

function markAppliedBulkTemplate(templateId) {
  state.bulkExportTemplates.appliedTemplateId = templateId ?? null;
  renderBulkExportTemplates();
}

function clearAppliedBulkTemplateTracking() {
  if (!state.bulkExportTemplates.appliedTemplateId) {
    return;
  }
  state.bulkExportTemplates.appliedTemplateId = null;
  const container = elements.bulkExportTemplatesContainer;
  if (!container) {
    return;
  }
  const activeCard = container.querySelector('.bulk-export-template-card--active');
  if (activeCard) {
    activeCard.classList.remove('bulk-export-template-card--active');
  }
  const activeButton = container.querySelector('.bulk-export-template-card__apply:disabled');
  if (activeButton instanceof HTMLButtonElement) {
    activeButton.disabled = false;
    activeButton.textContent = 'Применить';
  }
}

function applyBulkExportTemplate(templateId, options = {}) {
  const { silent = false, force = false } = options;
  const templates = [...state.bulkExportTemplates.system, ...state.bulkExportTemplates.custom];
  const template = templates.find((item) => item.id === templateId);
  if (!template) {
    if (!silent) {
      setBulkExportMessage('Не удалось применить шаблон: он не найден.');
    }
    return;
  }
  if (!force && state.bulkExportRules.length && !silent) {
    const confirmation = window.confirm(
      `Текущие правила массовой выгрузки будут заменены шаблоном «${template.name}». Продолжить?`,
    );
    if (!confirmation) {
      return;
    }
  }
  const periods = normalizeBulkExportPeriods(state.bulkExportPeriods);
  const preparedRules = template.rules.map((rule, index) =>
    normalizeBulkExportRule({
      ...rule,
      id: generateBulkRuleId(),
      periods,
    }),
  );
  state.bulkExportSearch = {};
  state.bulkExportRules = preparedRules;
  preparedRules.forEach((rule) => ensureBulkExportSearchState(rule.id));
  saveBulkExportRulesToStorage();
  markAppliedBulkTemplate(template.id);
  renderBulkExportRules();
  if (!silent) {
    setBulkExportMessage(`Шаблон «${template.name}» применён. Проверьте параметры перед выгрузкой.`);
  }
}

function handleSaveBulkExportTemplate() {
  if (!state.bulkExportRules.length) {
    setBulkExportMessage('Добавьте хотя бы одно правило, прежде чем сохранять шаблон.');
    return;
  }
  const invalidRules = [];
  state.bulkExportRules.forEach((rule, index) => {
    const { errors } = validateBulkRule(rule);
    if (errors.length) {
      invalidRules.push(`Правило ${index + 1}: ${errors.join(', ')}`);
    }
  });
  if (invalidRules.length) {
    setBulkExportMessage(
      `Не удалось сохранить шаблон: скорректируйте настройки. ${invalidRules.join('; ')}.`,
    );
    return;
  }
  const name = window.prompt('Введите название шаблона массовой выгрузки');
  if (name === null) {
    setBulkExportMessage('Сохранение шаблона отменено.');
    return;
  }
  const trimmedName = name.trim();
  if (!trimmedName) {
    setBulkExportMessage('Укажите осмысленное название для шаблона.');
    return;
  }
  const normalizedRules = state.bulkExportRules.map((rule, index) => serializeRuleToTemplate(rule, index));
  const existingIndex = state.bulkExportTemplates.custom.findIndex(
    (template) => normalizeKey(template.name) === normalizeKey(trimmedName),
  );
  if (existingIndex !== -1) {
    const shouldReplace = window.confirm(
      `Шаблон «${state.bulkExportTemplates.custom[existingIndex].name}» уже существует. Обновить его?`,
    );
    if (!shouldReplace) {
      setBulkExportMessage('Сохранение шаблона отменено.');
      return;
    }
    const existing = state.bulkExportTemplates.custom[existingIndex];
    const updated = normalizeBulkExportTemplate({
      ...existing,
      name: trimmedName,
      rules: normalizedRules,
    });
    state.bulkExportTemplates.custom = [
      ...state.bulkExportTemplates.custom.slice(0, existingIndex),
      { ...updated, id: existing.id },
      ...state.bulkExportTemplates.custom.slice(existingIndex + 1),
    ];
    saveCustomBulkExportTemplatesToStorage();
    markAppliedBulkTemplate(existing.id);
    setBulkExportMessage(`Шаблон «${trimmedName}» обновлён и доступен всем пользователям этого устройства.`);
    return;
  }
  const template = normalizeBulkExportTemplate({
    id: generateBulkTemplateId(),
    name: trimmedName,
    rules: normalizedRules,
  });
  state.bulkExportTemplates.custom = [...state.bulkExportTemplates.custom, template];
  saveCustomBulkExportTemplatesToStorage();
  markAppliedBulkTemplate(template.id);
  setBulkExportMessage(`Шаблон «${trimmedName}» сохранён и добавлен к списку доступных.`);
}

function deleteCustomBulkExportTemplate(templateId) {
  const index = state.bulkExportTemplates.custom.findIndex((template) => template.id === templateId);
  if (index === -1) {
    setBulkExportMessage('Не удалось найти шаблон для удаления.');
    return;
  }
  const template = state.bulkExportTemplates.custom[index];
  const confirmed = window.confirm(`Удалить шаблон «${template.name}»?`);
  if (!confirmed) {
    return;
  }
  state.bulkExportTemplates.custom = [
    ...state.bulkExportTemplates.custom.slice(0, index),
    ...state.bulkExportTemplates.custom.slice(index + 1),
  ];
  saveCustomBulkExportTemplatesToStorage();
  if (state.bulkExportTemplates.appliedTemplateId === templateId) {
    state.bulkExportTemplates.appliedTemplateId = null;
  }
  renderBulkExportTemplates();
  setBulkExportMessage(`Шаблон «${template.name}» удалён.`);
}

function createBulkExportRuleElement(rule, index) {
  const wrapper = document.createElement('article');
  wrapper.className = 'bulk-export-rule';
  wrapper.dataset.ruleId = rule.id;

  const header = document.createElement('div');
  header.className = 'bulk-export-rule__header';

  const titleInput = document.createElement('input');
  titleInput.type = 'text';
  titleInput.className = 'bulk-export-rule__title-input';
  titleInput.value = rule.name;
  titleInput.placeholder = `Правило ${index + 1}`;
  titleInput.addEventListener('input', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }
    updateBulkExportRule(rule.id, { name: target.value }, { render: false });
  });
  titleInput.addEventListener('blur', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }
    updateBulkExportRule(rule.id, { name: target.value.trim() });
  });

  const removeButton = document.createElement('button');
  removeButton.type = 'button';
  removeButton.className = 'bulk-export-remove';
  removeButton.textContent = 'Удалить';
  removeButton.addEventListener('click', () => deleteBulkExportRule(rule.id));

  header.append(titleInput, removeButton);
  wrapper.append(header);

  const grid = document.createElement('div');
  grid.className = 'bulk-export-rule__grid';

  grid.append(createTypeFieldset(rule));
  grid.append(createViolationFieldset(rule));
  grid.append(createAdditionalFieldset(rule));

  wrapper.append(grid);

  const issues = collectBulkRuleIssues(rule);
  if (issues.length) {
    const warning = document.createElement('div');
    warning.className = 'bulk-export-warning';
    warning.textContent = issues.join(' ');
    wrapper.append(warning);
  }

  return wrapper;
}

function renderBulkExportSharedPeriods() {
  const container = elements.bulkExportPeriodsContainer;
  if (!container) {
    return;
  }
  container.innerHTML = '';

  const fieldset = document.createElement('fieldset');
  fieldset.className = 'bulk-export-fieldset';
  const legend = document.createElement('legend');
  legend.textContent = 'Периоды для массовой выгрузки';
  fieldset.append(legend);

  fieldset.append(createSharedPeriodGroup('current', 'Отчётный период'));
  fieldset.append(createSharedPeriodGroup('previous', 'Предыдущий период'));

  container.append(fieldset);

  if (elements.bulkExportPeriodsMessage) {
    const issues = collectBulkPeriodsIssues(state.bulkExportPeriods);
    elements.bulkExportPeriodsMessage.textContent = issues.join(' ');
  }
}

function createSharedPeriodGroup(prefix, title) {
  const wrapper = document.createElement('div');
  wrapper.className = 'bulk-export-period-group';

  const groupTitle = document.createElement('span');
  groupTitle.className = 'bulk-export-period-group__title';
  groupTitle.textContent = title;
  wrapper.append(groupTitle);

  const inputs = document.createElement('div');
  inputs.className = 'bulk-export-date-inputs';
  inputs.append(
    createSharedDateInput(`${prefix}Start`, 'Начало'),
    createSharedDateInput(`${prefix}End`, 'Окончание'),
  );
  wrapper.append(inputs);

  return wrapper;
}

function createSharedDateInput(key, label) {
  const wrapper = document.createElement('label');
  wrapper.className = 'bulk-export-date-input';

  const text = document.createElement('span');
  text.textContent = label;
  wrapper.append(text);

  const input = document.createElement('input');
  input.type = 'date';
  input.value = state.bulkExportPeriods[key] ?? '';
  const availableRange = deriveAvailableDateRange();
  if (availableRange) {
    input.min = availableRange.min;
    input.max = availableRange.max;
  }
  input.addEventListener('change', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }
    updateBulkExportPeriods({ [key]: target.value });
  });
  wrapper.append(input);
  return wrapper;
}

function updateBulkExportPeriods(changes) {
  const next = normalizeBulkExportPeriods({
    ...state.bulkExportPeriods,
    ...changes,
  });
  if (JSON.stringify(next) === JSON.stringify(state.bulkExportPeriods)) {
    return;
  }
  state.bulkExportPeriods = next;
  saveBulkExportPeriodsToStorage();
  if (elements.bulkExportPeriodsMessage) {
    const issues = collectBulkPeriodsIssues(state.bulkExportPeriods);
    elements.bulkExportPeriodsMessage.textContent = issues.join(' ');
  }
}

function collectBulkPeriodsIssues(periods) {
  const issues = [];
  if (!state.availableDates.length) {
    issues.push('Загрузите данные и сопоставьте колонку «Дата обследования», чтобы выбрать периоды для массовой выгрузки.');
    return issues;
  }
  const normalized = normalizeBulkExportPeriods(periods);
  const currentStart = parseIsoDate(normalized.currentStart);
  const currentEnd = parseIsoDate(normalized.currentEnd);
  if (!currentStart || !currentEnd) {
    issues.push('Укажите даты начала и окончания отчётного периода.');
  } else if (currentStart > currentEnd) {
    issues.push('Дата начала отчётного периода не может быть позже даты окончания.');
  }
  const previousStart = parseIsoDate(normalized.previousStart);
  const previousEnd = parseIsoDate(normalized.previousEnd);
  if (!previousStart || !previousEnd) {
    issues.push('Заполните предыдущий период.');
  } else if (previousStart > previousEnd) {
    issues.push('Дата начала предыдущего периода не может быть позже даты окончания.');
  }
  return issues;
}

function validateBulkPeriods(periods) {
  const errors = [];
  const normalized = normalizeBulkExportPeriods(periods);
  const currentStart = parseIsoDate(normalized.currentStart);
  const currentEnd = parseIsoDate(normalized.currentEnd);
  if (!currentStart || !currentEnd) {
    errors.push('не заполнен отчётный период');
  } else if (currentStart > currentEnd) {
    errors.push('отчётный период указан некорректно');
  }
  const previousStart = parseIsoDate(normalized.previousStart);
  const previousEnd = parseIsoDate(normalized.previousEnd);
  if (!previousStart || !previousEnd) {
    errors.push('не заполнен предыдущий период');
  } else if (previousStart > previousEnd) {
    errors.push('предыдущий период указан некорректно');
  }
  const result =
    !errors.length && currentStart && currentEnd && previousStart && previousEnd
      ? {
          current: { start: currentStart, end: currentEnd },
          previous: { start: previousStart, end: previousEnd },
        }
      : null;
  return { periods: result, errors };
}

function syncBulkExportPeriodsWithAvailability() {
  const normalized = normalizeBulkExportPeriods(state.bulkExportPeriods);
  let next = { ...normalized };
  let changed = false;
  if (!state.availableDates.length) {
    if (Object.values(next).some(Boolean)) {
      next = {
        currentStart: '',
        currentEnd: '',
        previousStart: '',
        previousEnd: '',
      };
      changed = true;
    }
  } else {
    const range = deriveAvailableDateRange();
    if (range) {
      for (const key of ['currentStart', 'currentEnd', 'previousStart', 'previousEnd']) {
        const value = next[key];
        if (value && (value < range.min || value > range.max)) {
          next[key] = '';
          changed = true;
        }
      }
    }
  }
  if (changed) {
    state.bulkExportPeriods = next;
    saveBulkExportPeriodsToStorage();
  } else {
    state.bulkExportPeriods = normalized;
  }
  renderBulkExportSharedPeriods();
}

function deriveAvailableDateRange() {
  if (!state.availableDates.length) {
    return null;
  }
  const first = state.availableDates[0]?.iso;
  const last = state.availableDates[state.availableDates.length - 1]?.iso;
  if (!first || !last) {
    return null;
  }
  return { min: first, max: last };
}

function formatRuPlural(count, form1, form2, form5) {
  const mod10 = count % 10;
  const mod100 = count % 100;
  if (mod10 === 1 && mod100 !== 11) {
    return form1;
  }
  if (mod10 >= 2 && mod10 <= 4 && (mod100 < 10 || mod100 >= 20)) {
    return form2;
  }
  return form5;
}

function normalizeSearchValue(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase();
}

function ensureBulkExportSearchState(ruleId) {
  if (!state.bulkExportSearch[ruleId]) {
    state.bulkExportSearch[ruleId] = { types: '', violations: '' };
  } else {
    const current = state.bulkExportSearch[ruleId];
    if (typeof current.types !== 'string') {
      current.types = '';
    }
    if (typeof current.violations !== 'string') {
      current.violations = '';
    }
  }
  return state.bulkExportSearch[ruleId];
}

function createBulkExportSelectionSummary(values, emptyText, forms) {
  const wrapper = document.createElement('div');
  wrapper.className = 'bulk-export-selection';
  if (!values.length) {
    wrapper.classList.add('bulk-export-selection--empty');
    const caption = document.createElement('span');
    caption.className = 'bulk-export-selection__caption';
    caption.textContent = emptyText;
    wrapper.append(caption);
    return wrapper;
  }
  const caption = document.createElement('span');
  caption.className = 'bulk-export-selection__caption';
  const plural = formatRuPlural(values.length, forms[0], forms[1], forms[2]);
  caption.textContent = `Выбрано ${values.length} ${plural}`;
  wrapper.append(caption);

  const chips = document.createElement('div');
  chips.className = 'bulk-export-chips';
  values.forEach((value) => {
    const chip = document.createElement('span');
    chip.className = 'bulk-export-chip';
    chip.textContent = value;
    chips.append(chip);
  });
  wrapper.append(chips);
  return wrapper;
}

function createBulkExportSearchControl({ ruleId, key, optionsWrapper, placeholder, emptyMessage }) {
  const searchState = ensureBulkExportSearchState(ruleId);
  const wrapper = document.createElement('div');
  wrapper.className = 'bulk-export-search';

  const icon = document.createElement('span');
  icon.className = 'bulk-export-search__icon';
  icon.setAttribute('aria-hidden', 'true');
  icon.textContent = '🔍';

  const input = document.createElement('input');
  input.type = 'search';
  input.className = 'bulk-export-search__input';
  input.placeholder = placeholder;
  input.setAttribute('aria-label', placeholder);
  input.value = typeof searchState[key] === 'string' ? searchState[key] : '';

  const clearButton = document.createElement('button');
  clearButton.type = 'button';
  clearButton.className = 'bulk-export-search__clear';
  clearButton.textContent = 'Очистить';
  clearButton.setAttribute('aria-label', 'Очистить поиск');

  const emptyIndicator = document.createElement('p');
  emptyIndicator.className = 'bulk-export-search__empty';
  emptyIndicator.textContent = emptyMessage;
  emptyIndicator.hidden = true;

  const applyFilter = () => {
    const term = normalizeSearchValue(input.value);
    searchState[key] = input.value;
    let matchCount = 0;
    const options = optionsWrapper.querySelectorAll('.bulk-export-option');
    options.forEach((option) => {
      const value = option.dataset.searchValue ?? '';
      const isMatch = !term || value.includes(term);
      option.hidden = !isMatch;
      if (isMatch) {
        matchCount += 1;
      }
    });
    emptyIndicator.hidden = matchCount > 0;
    clearButton.hidden = term.length === 0;
  };

  input.addEventListener('input', applyFilter);
  input.addEventListener('search', applyFilter);
  clearButton.addEventListener('click', () => {
    if (!input.value) {
      return;
    }
    input.value = '';
    searchState[key] = '';
    applyFilter();
    input.focus();
  });

  wrapper.append(icon, input, clearButton);
  applyFilter();

  return { wrapper, emptyIndicator, applyFilter };
}

function createTypeFieldset(rule) {
  const fieldset = document.createElement('fieldset');
  fieldset.className = 'bulk-export-fieldset';
  const legend = document.createElement('legend');
  legend.textContent = 'Типы объектов контроля';
  fieldset.append(legend);

  const select = document.createElement('select');
  const optionAll = document.createElement('option');
  optionAll.value = 'all';
  optionAll.textContent = 'Все типы объектов';
  const optionCustom = document.createElement('option');
  optionCustom.value = 'custom';
  optionCustom.textContent = 'Выбрать вручную';
  select.append(optionAll, optionCustom);
  select.value = rule.typeMode;
  select.addEventListener('change', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement)) {
      return;
    }
    const nextMode = target.value === 'custom' ? 'custom' : 'all';
    updateBulkExportRule(rule.id, {
      typeMode: nextMode,
      customTypes: nextMode === 'custom' ? rule.customTypes : [],
    });
  });
  fieldset.append(select);

  if (rule.typeMode !== 'custom') {
    const note = document.createElement('p');
    note.className = 'bulk-export-note';
    note.textContent = 'Используются все типы объектов из загруженных данных.';
    fieldset.append(note);
    return fieldset;
  }

  if (!state.availableTypes.length) {
    const note = document.createElement('p');
    note.className = 'bulk-export-note';
    note.textContent = 'Типы объектов появятся после загрузки файлов и сопоставления колонок.';
    fieldset.append(note);
    return fieldset;
  }

  const selected = new Set(rule.customTypes);
  fieldset.append(
    createBulkExportSelectionSummary(
      rule.customTypes,
      'Вы пока не выбрали типы объектов.',
      ['тип', 'типа', 'типов'],
    ),
  );

  const optionsWrapper = document.createElement('div');
  optionsWrapper.className = 'bulk-export-options';
  const limitReached = rule.customTypes.length >= 3;
  for (const value of state.availableTypes) {
    const optionLabel = document.createElement('label');
    optionLabel.className = 'bulk-export-option';
    optionLabel.dataset.searchValue = normalizeSearchValue(value);
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.value = value;
    input.checked = selected.has(value);
    if (!selected.has(value) && limitReached) {
      input.disabled = true;
    }
    input.addEventListener('change', (event) => {
      const target = event.target;
      if (!(target instanceof HTMLInputElement)) {
        return;
      }
      if (target.checked) {
        if (rule.customTypes.length >= 3) {
          target.checked = false;
          setBulkExportMessage('Можно выбрать не более трех типов объектов.');
          return;
        }
        updateBulkExportRule(rule.id, {
          customTypes: [...rule.customTypes, value],
        });
      } else {
        updateBulkExportRule(rule.id, {
          customTypes: rule.customTypes.filter((item) => item !== value),
        });
      }
    });
    const text = document.createElement('span');
    text.textContent = value;
    optionLabel.append(input, text);
    optionsWrapper.append(optionLabel);
  }

  const searchControls = createBulkExportSearchControl({
    ruleId: rule.id,
    key: 'types',
    optionsWrapper,
    placeholder: 'Поиск по типам объектов',
    emptyMessage: 'Совпадений не найдено. Попробуйте изменить запрос.',
  });
  fieldset.append(searchControls.wrapper);
  fieldset.append(optionsWrapper);
  fieldset.append(searchControls.emptyIndicator);

  const note = document.createElement('p');
  note.className = 'bulk-export-note';
  note.textContent = `Можно выбрать не более трех типов объектов. Выбрано: ${rule.customTypes.length}.`;
  fieldset.append(note);

  return fieldset;
}

function createViolationFieldset(rule) {
  const fieldset = document.createElement('fieldset');
  fieldset.className = 'bulk-export-fieldset';
  const legend = document.createElement('legend');
  legend.textContent = 'Наименования нарушений';
  fieldset.append(legend);

  const select = document.createElement('select');
  const optionAll = document.createElement('option');
  optionAll.value = 'all';
  optionAll.textContent = 'Все нарушения';
  const optionCustom = document.createElement('option');
  optionCustom.value = 'custom';
  optionCustom.textContent = 'Выбрать вручную';
  select.append(optionAll, optionCustom);
  select.value = rule.violationMode;
  select.addEventListener('change', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement)) {
      return;
    }
    const nextMode = target.value === 'custom' ? 'custom' : 'all';
    updateBulkExportRule(rule.id, {
      violationMode: nextMode,
      customViolations: nextMode === 'custom' ? rule.customViolations : [],
    });
  });
  fieldset.append(select);

  if (rule.violationMode !== 'custom') {
    const note = document.createElement('p');
    note.className = 'bulk-export-note';
    note.textContent = 'Фильтр по наименованиям нарушений не применяется.';
    fieldset.append(note);
    return fieldset;
  }

  if (!state.availableViolations.length) {
    const note = document.createElement('p');
    note.className = 'bulk-export-note';
    note.textContent = 'Загрузите таблицу нарушений и сопоставьте колонку, чтобы выбрать наименования.';
    fieldset.append(note);
    return fieldset;
  }

  const selected = new Set(rule.customViolations);
  fieldset.append(
    createBulkExportSelectionSummary(
      rule.customViolations,
      'Вы пока не выбрали наименования нарушений.',
      ['наименование', 'наименования', 'наименований'],
    ),
  );

  const optionsWrapper = document.createElement('div');
  optionsWrapper.className = 'bulk-export-options';
  const limitReached = rule.customViolations.length >= 5;
  for (const value of state.availableViolations) {
    const optionLabel = document.createElement('label');
    optionLabel.className = 'bulk-export-option';
    optionLabel.dataset.searchValue = normalizeSearchValue(value);
    const input = document.createElement('input');
    input.type = 'checkbox';
    input.value = value;
    input.checked = selected.has(value);
    if (!selected.has(value) && limitReached) {
      input.disabled = true;
    }
    input.addEventListener('change', (event) => {
      const target = event.target;
      if (!(target instanceof HTMLInputElement)) {
        return;
      }
      if (target.checked) {
        if (rule.customViolations.length >= 5) {
          target.checked = false;
          setBulkExportMessage('Можно выбрать не более пяти наименований нарушений.');
          return;
        }
        updateBulkExportRule(rule.id, {
          customViolations: [...rule.customViolations, value],
        });
      } else {
        updateBulkExportRule(rule.id, {
          customViolations: rule.customViolations.filter((item) => item !== value),
        });
      }
    });
    const text = document.createElement('span');
    text.textContent = value;
    optionLabel.append(input, text);
    optionsWrapper.append(optionLabel);
  }

  const searchControls = createBulkExportSearchControl({
    ruleId: rule.id,
    key: 'violations',
    optionsWrapper,
    placeholder: 'Поиск по наименованиям нарушений',
    emptyMessage: 'Совпадений не найдено. Скорректируйте поисковый запрос.',
  });
  fieldset.append(searchControls.wrapper);
  fieldset.append(optionsWrapper);
  fieldset.append(searchControls.emptyIndicator);

  const note = document.createElement('p');
  note.className = 'bulk-export-note';
  note.textContent = `Можно выбрать не более пяти наименований нарушений. Выбрано: ${rule.customViolations.length}.`;
  fieldset.append(note);

  return fieldset;
}

function createAdditionalFieldset(rule) {
  const fieldset = document.createElement('fieldset');
  fieldset.className = 'bulk-export-fieldset';
  const legend = document.createElement('legend');
  legend.textContent = 'Дополнительные параметры';
  fieldset.append(legend);

  const sourceLabel = document.createElement('label');
  const sourceText = document.createElement('span');
  sourceText.textContent = 'Источник данных';
  sourceLabel.append(sourceText);
  const select = document.createElement('select');
  const optionAll = document.createElement('option');
  optionAll.value = 'all';
  optionAll.textContent = 'ОАТИ и ЦАФАП';
  const optionOati = document.createElement('option');
  optionOati.value = 'oati';
  optionOati.textContent = 'Только ОАТИ';
  const optionCafap = document.createElement('option');
  optionCafap.value = 'cafap';
  optionCafap.textContent = 'Только ЦАФАП';
  select.append(optionAll, optionOati, optionCafap);
  select.value = rule.dataSourceOption;
  const hasDataSourceColumn = Boolean(state.violationMapping.dataSource);
  select.disabled = !hasDataSourceColumn;
  select.addEventListener('change', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLSelectElement)) {
      return;
    }
    const value = target.value;
    updateBulkExportRule(rule.id, {
      dataSourceOption: ['oati', 'cafap'].includes(value) ? value : 'all',
    });
  });
  sourceLabel.append(select);
  fieldset.append(sourceLabel);

  if (!hasDataSourceColumn && rule.dataSourceOption !== 'all') {
    const note = document.createElement('p');
    note.className = 'bulk-export-note';
    note.textContent = 'Фильтр по источнику станет доступен после сопоставления соответствующей колонки.';
    fieldset.append(note);
  }

  const tinaoLabel = document.createElement('label');
  tinaoLabel.className = 'bulk-export-option';
  const checkbox = document.createElement('input');
  checkbox.type = 'checkbox';
  checkbox.checked = rule.combineTiNaoDistricts;
  checkbox.addEventListener('change', (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) {
      return;
    }
    updateBulkExportRule(rule.id, { combineTiNaoDistricts: target.checked }, { render: false });
  });
  const checkboxText = document.createElement('span');
  checkboxText.textContent = 'Объединять округа НАО и ТАО в «ТиНАО»';
  tinaoLabel.append(checkbox, checkboxText);
  fieldset.append(tinaoLabel);

  return fieldset;
}

function collectBulkRuleIssues(rule) {
  const issues = [];
  if (rule.typeMode === 'custom' && !rule.customTypes.length) {
    issues.push('Выберите типы объектов или переключитесь на использование всех типов.');
  }
  if (rule.violationMode === 'custom' && !rule.customViolations.length) {
    issues.push('Добавьте хотя бы одно наименование нарушения или отключите фильтр.');
  }
  if (!state.violationMapping.dataSource && rule.dataSourceOption !== 'all') {
    issues.push('Фильтр по источнику будет применён только после сопоставления колонки «Источник данных».');
  }
  return issues;
}

function validateBulkRule(rule) {
  const errors = [];
  if (rule.typeMode === 'custom' && !rule.customTypes.length) {
    errors.push('не выбраны типы объектов для фильтра');
  }
  if (rule.violationMode === 'custom' && !rule.customViolations.length) {
    errors.push('не выбраны наименования нарушений для фильтра');
  }
  if (!state.violationMapping.dataSource && rule.dataSourceOption !== 'all') {
    errors.push('фильтр по источнику данных станет доступен после сопоставления колонки «Источник данных»');
  }
  return { errors };
}

function syncBulkExportRulesWithAvailableOptions() {
  if (!state.bulkExportRules.length) {
    renderBulkExportRules();
    syncBulkExportPeriodsWithAvailability();
    return;
  }
  const typeSet = new Set(state.availableTypes);
  const violationSet = new Set(state.availableViolations);
  let changed = false;
  const next = state.bulkExportRules.map((rule) => {
    const filteredTypes = rule.customTypes.filter((value) => typeSet.has(value));
    const filteredViolations = rule.customViolations.filter((value) => violationSet.has(value));
    if (filteredTypes.length !== rule.customTypes.length || filteredViolations.length !== rule.customViolations.length) {
      changed = true;
    }
    return normalizeBulkExportRule({
      ...rule,
      customTypes: filteredTypes,
      customViolations: filteredViolations,
    });
  });
  if (changed) {
    state.bulkExportRules = next;
    saveBulkExportRulesToStorage();
  }
  renderBulkExportRules();
  syncBulkExportPeriodsWithAvailability();
}

function openBulkExportOverlay() {
  if (!elements.bulkExportOverlay) {
    return;
  }
  elements.bulkExportOverlay.hidden = false;
  document.body.classList.add('no-scroll');
  setBulkExportMessage('');
  renderBulkExportSharedPeriods();
  renderBulkExportRules();
  if (elements.bulkExportPanel) {
    elements.bulkExportPanel.focus();
  }
}

function closeBulkExportOverlay() {
  if (!elements.bulkExportOverlay) {
    return;
  }
  elements.bulkExportOverlay.hidden = true;
  document.body.classList.remove('no-scroll');
  const trigger = elements.openBulkExportButton;
  if (trigger instanceof HTMLElement) {
    const isHidden = trigger.hasAttribute('hidden') || Boolean(trigger.closest('[hidden]'));
    if (!isHidden) {
      try {
        trigger.focus();
      } catch (error) {
        console.warn('Не удалось вернуть фокус к кнопке открытия массовой выгрузки.', error);
      }
    }
  }
}

function isBulkOverlayOpen() {
  return Boolean(elements.bulkExportOverlay && !elements.bulkExportOverlay.hidden);
}

function handleBulkOverlayKeydown(event) {
  if (event.key === 'Escape' && isBulkOverlayOpen()) {
    closeBulkExportOverlay();
  }
}

function handleBulkExportDownload() {
  if (!state.bulkExportRules.length) {
    setBulkExportMessage('Добавьте хотя бы одно правило массовой выгрузки.');
    return;
  }
  if (!(state.violations.length && state.objects.length)) {
    setBulkExportMessage('Загрузите таблицы нарушений и объектов, чтобы сформировать массовую выгрузку.');
    return;
  }
  if (
    !isMappingComplete(state.violationMapping, violationFieldDefinitions) ||
    !isMappingComplete(state.objectMapping, objectFieldDefinitions)
  ) {
    setBulkExportMessage('Проверьте настройки соответствия столбцов перед массовой выгрузкой.');
    return;
  }
  const { periods, errors: periodErrors } = validateBulkPeriods(state.bulkExportPeriods);
  if (!periods || periodErrors.length) {
    setBulkExportMessage(`Не удалось подготовить массовую выгрузку: ${periodErrors.join(', ')}.`);
    renderBulkExportSharedPeriods();
    return;
  }
  const problems = [];
  const prepared = [];
  state.bulkExportRules.forEach((rule, index) => {
    const { errors } = validateBulkRule(rule);
    if (errors.length) {
      const ruleName = rule.name || `Правило ${index + 1}`;
      problems.push(`${ruleName}: ${errors.join(', ')}`);
      return;
    }
    const context = {
      typeMode: rule.typeMode,
      customTypes: rule.typeMode === 'custom' ? rule.customTypes : [],
      violationMode: rule.violationMode,
      customViolations: rule.violationMode === 'custom' ? rule.customViolations : [],
      dataSourceOption: state.violationMapping.dataSource ? rule.dataSourceOption : 'all',
      combineTiNaoDistricts: rule.combineTiNaoDistricts,
    };
    prepared.push({ rule, periods, context, index });
  });
  if (problems.length) {
    setBulkExportMessage(`Не удалось подготовить массовую выгрузку: ${problems.join('; ')}.`);
    return;
  }
  try {
    const workbook = XLSX.utils.book_new();
    prepared.forEach(({ rule, periods, context, index }) => {
      const report = buildReport(periods, context);
      const sheetLabel = rule.name ? `${index + 1}. ${rule.name}` : `Правило ${index + 1}`;
      const { sheet, sheetName } = createWorksheetForReport(report, periods, {
        context,
        sheetName: sheetLabel,
        titlePrefix: rule.name || `Правило ${index + 1}`,
      });
      XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
    });
    const fileName = buildBulkExportFileName();
    XLSX.writeFile(workbook, fileName);
    setBulkExportMessage(`Скачан файл «${fileName}».`);
  } catch (error) {
    console.error('Не удалось сформировать массовую выгрузку.', error);
    setBulkExportMessage('Произошла ошибка при формировании файла. Повторите попытку.');
  }
}

function buildBulkExportFileName() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  const hours = String(now.getHours()).padStart(2, '0');
  const minutes = String(now.getMinutes()).padStart(2, '0');
  const raw = `Массовая выгрузка ОАТИ ${year}-${month}-${day}_${hours}-${minutes}.xlsx`;
  return raw.replace(/[\\/:*?"<>|]+/g, '_');
}

function initializeBulkExportFeature() {
  state.bulkExportSearch = {};
  state.bulkExportTemplates = {
    system: SYSTEM_BULK_EXPORT_TEMPLATES,
    custom: loadCustomBulkExportTemplatesFromStorage(),
    appliedTemplateId: null,
  };
  const storedRules = loadBulkExportRulesFromStorage();
  state.bulkExportRules = storedRules;
  state.bulkExportPeriods = loadBulkExportPeriodsFromStorage();
  if (state.bulkExportRules.length) {
    state.bulkExportRules.forEach((rule) => ensureBulkExportSearchState(rule.id));
    renderBulkExportRules();
    renderBulkExportTemplates();
  } else if (state.bulkExportTemplates.system.length) {
    applyBulkExportTemplate(state.bulkExportTemplates.system[0].id, { silent: true, force: true });
  } else {
    renderBulkExportRules();
    renderBulkExportTemplates();
  }
  syncBulkExportPeriodsWithAvailability();
  if (elements.bulkExportOverlay) {
    elements.bulkExportOverlay.hidden = true;
  }
  if (document.body) {
    document.body.classList.remove('no-scroll');
  }
}

/**
 * Переключает доступность кнопки экспорта в Excel.
 *
 * @param {boolean} isEnabled Признак доступности выгрузки.
 */
function setExportAvailability(isEnabled) {
  if (!elements.downloadButton) {
    return;
  }
  elements.downloadButton.disabled = !isEnabled;
  elements.downloadButton.setAttribute('aria-disabled', String(!isEnabled));
}

/**
 * Сбрасывает кеш отчёта и блокирует кнопку выгрузки.
 */
function resetExportState() {
  state.lastReport = null;
  state.lastPeriods = null;
  setExportAvailability(false);
}

/**
 * Отменяет запланированное обновление предпросмотра.
 */
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
 * Планирует пересчёт предпросмотра. Для групп событий используется requestAnimationFrame,
 * а при принудительном обновлении пересчёт выполняется немедленно.
 *
 * @param {{immediate?: boolean}} [options] Настройки поведения функции.
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
 * Выполняет пересчёт отчёта с учётом всех выбранных фильтров и параметров.
 * Результат отображается в таблице предпросмотра и готовится к экспорту.
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

/**
 * Проверяет, что пользователь указал колонки для всех обязательных полей.
 *
 * @param {Record<string, string>} mapping Карта соответствий «ключ поля → колонка».
 * @param {Array<{key: string, optional?: boolean}>} definitions Список ожидаемых полей.
 * @returns {boolean} true, если сопоставление заполнено.
 */
function isMappingComplete(mapping, definitions) {
  return definitions.every((definition) => definition.optional || Boolean(mapping[definition.key]));
}

/**
 * Считывает выбранные пользователем периоды отчёта и предыдущего периода.
 * Одновременно выполняет базовую валидацию.
 *
 * @returns {{current: {start: Date, end: Date}, previous: {start: Date, end: Date}} | null}
 *   Возвращает объект с выбранными периодами или null при ошибке.
 */
function getSelectedPeriods() {
  const current = parsePeriodSelection(
    elements.currentStart?.value ?? '',
    elements.currentEnd?.value ?? '',
    {
      missing: 'Выберите даты начала и окончания отчётного периода.',
      invalid: 'Выберите корректные даты отчётного периода.',
      order: 'Дата начала отчётного периода не может быть позже даты окончания.',
    },
  );
  if (!current) {
    return null;
  }

  const previous = parsePeriodSelection(
    elements.previousStart?.value ?? '',
    elements.previousEnd?.value ?? '',
    {
      missing: 'Выберите даты предыдущего периода.',
      invalid: 'Выберите корректные даты предыдущего периода.',
      order: 'Дата начала предыдущего периода не может быть позже даты окончания.',
    },
  );
  if (!previous) {
    return null;
  }

  return { current, previous };
}

/**
 * Валидирует и преобразует выбранный пользователем диапазон дат.
 *
 * @param {string} startIso Дата начала периода в формате ISO.
 * @param {string} endIso Дата окончания периода в формате ISO.
 * @param {{missing: string, invalid: string, order: string}} messages Сообщения об ошибках.
 * @returns {{start: Date, end: Date} | null} Диапазон дат или null, если проверка не пройдена.
 */
function parsePeriodSelection(startIso, endIso, messages) {
  const startValue = typeof startIso === 'string' ? startIso.trim() : '';
  const endValue = typeof endIso === 'string' ? endIso.trim() : '';
  if (!startValue || !endValue) {
    showPreviewMessage(messages.missing);
    return null;
  }
  const startDate = parseIsoDate(startValue);
  const endDate = parseIsoDate(endValue);
  if (!startDate || !endDate) {
    showPreviewMessage(messages.invalid);
    return null;
  }
  if (startDate > endDate) {
    showPreviewMessage(messages.order);
    return null;
  }
  return { start: startDate, end: endDate };
}

/**
 * Возвращает человекочитаемое название округа, предпочитая официальные сокращения.
 *
 * @param {string} label Исходное значение из данных.
 * @returns {string} Округ для отображения.
 */
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

/**
 * Нормализует название округа для внутренних ключей (в нижнем регистре без лишних символов).
 *
 * @param {string} value Значение округа.
 * @returns {string} Нормализованный ключ.
 */
function normalizeDistrictKey(value) {
  const normalized = normalizeKey(value);
  if (!normalized) {
    return '';
  }
  const withoutParentheses = normalized.replace(/\s*\(.*?\)\s*/g, ' ');
  const withoutCityMention = withoutParentheses.replace(/\s+г\.?\s*москв[аеы]?$/u, '');
  return withoutCityMention.replace(/\s+/g, ' ').trim();
}

/**
 * Формирует таблицу соответствий между оригинальными подписями округов и отображаемыми значениями.
 *
 * @param {object[]} objectRecords Записи из справочника объектов.
 * @param {string} objectDistrictColumn Название колонки с округом в справочнике.
 * @param {object[]} violationRecords Записи из таблицы нарушений.
 * @param {string} violationDistrictColumn Название колонки с округом в таблице нарушений.
 * @returns {Map<string, string>} Словарь «нормализованный ключ → отображаемое значение».
 */
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

/**
 * Определяет, нужно ли заменить существующее отображаемое имя округа новым вариантом.
 *
 * @param {string} current Текущее значение в словаре.
 * @param {string} candidate Кандидат для замены.
 * @returns {boolean} true, если стоит заменить значение.
 */
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

function buildObjectIdentifierCandidates(idValue, nameValue, externalIdValue = '') {
  const candidates = [];
  const normalizedId = normalizeKey(getValueAsString(idValue));
  const normalizedExternalId = normalizeKey(getValueAsString(externalIdValue));
  const normalizedName = normalizeKey(getValueAsString(nameValue));
  if (normalizedId) {
    candidates.push(`id:${normalizedId}`);
  }
  if (normalizedExternalId) {
    candidates.push(`external:${normalizedExternalId}`);
  }
  if (normalizedName) {
    candidates.push(`name:${normalizedName}`);
  }
  if (candidates.length <= 1) {
    return candidates;
  }
  const unique = new Set();
  for (const candidate of candidates) {
    if (!unique.has(candidate)) {
      unique.add(candidate);
    }
  }
  return Array.from(unique);
}

function getPrimaryObjectIdentifier(candidates) {
  return candidates.length ? candidates[0] : '';
}

/**
 * Строит отчёт по округам на основании выбранных фильтров и периодов.
 *
 * @param {{current: {start: Date, end: Date}, previous: {start: Date, end: Date}}} periods
 *   Выбранные пользователем периоды.
 * @returns {{rows: Array<object>, totalRow: object}} Строки отчёта и итоговая строка.
 */
function buildReport(periods, overrides = {}) {
  const violationMapping = state.violationMapping;
  const objectMapping = state.objectMapping;
  const context = resolveReportContext(overrides);
  const typePredicate = createTypePredicate(context);
  const violationPredicate = createViolationPredicate(context);
  const dataSourcePredicate = createDataSourcePredicate(context);
  // Здесь буду копить агрегированные данные по округам.
  const districtData = new Map();
  const combineTiNao = context.combineTiNaoDistricts;
  const districtLookup = buildDistrictLookup(
    state.objects,
    objectMapping.district,
    state.violations,
    violationMapping.district,
  );
  const ensureEntry = (label) => {
    const normalizedKey = normalizeDistrictKey(label);
    const key = normalizedKey || DISTRICT_FALLBACK_KEY;
    let displayLabel =
      districtLookup.get(key) ?? (normalizedKey ? getDistrictDisplayLabel(label) : 'Без округа');
    let aggregatedKey = normalizeKey(displayLabel) || DISTRICT_FALLBACK_KEY;

    if (combineTiNao && shouldMergeWithTiNao(aggregatedKey)) {
      aggregatedKey = TINAO_AGGREGATED_KEY;
      displayLabel = TINAO_DISPLAY_LABEL;
    }
    if (!districtData.has(aggregatedKey)) {
      districtData.set(aggregatedKey, {
        label: displayLabel,
        totalObjects: new Set(),
        inspectedObjects: new Set(),
        objectsWithDetectedViolations: new Set(),
        currentViolationsCount: 0,
        previousControlCount: 0,
        resolvedCount: 0,
        currentControlStatusesCount: 0,
      });
    } else {
      const entry = districtData.get(aggregatedKey);
      if (entry && shouldReplaceDistrictLabel(entry.label, displayLabel)) {
        entry.label = displayLabel;
      }
    }
    return districtData.get(aggregatedKey);
  };

  const violationObjectIdColumn = violationMapping.objectId;
  const violationNameColumn = violationMapping.violationName;
  const objectIdColumn = objectMapping.objectId;
  const externalObjectIdColumn = objectMapping.externalObjectId;

  // Сначала прохожусь по объектам и считаю общее количество по округам.
  for (const record of state.objects) {
    const typeValue = getValueAsString(record[objectMapping.objectType]);
    if (typeValue && !typePredicate(typeValue)) {
      continue;
    }
    if (!typeValue && context.typeMode === 'custom') {
      continue;
    }
    const districtLabel = getValueAsString(record[objectMapping.district]);
    const externalIdentifierValue = externalObjectIdColumn
      ? getValueAsString(record[externalObjectIdColumn])
      : '';
    const normalizedExternalIdentifier = normalizeKey(externalIdentifierValue);
    const identifierCandidates = buildObjectIdentifierCandidates(
      objectIdColumn ? record[objectIdColumn] : '',
      objectMapping.objectName ? record[objectMapping.objectName] : '',
      externalObjectIdColumn ? record[externalObjectIdColumn] : '',
    );
    const primaryObjectIdentifier = getPrimaryObjectIdentifier(identifierCandidates);
    const totalObjectIdentifier = normalizedExternalIdentifier
      ? `external:${normalizedExternalIdentifier}`
      : primaryObjectIdentifier;
    if (!totalObjectIdentifier) {
      continue;
    }
    const entry = ensureEntry(districtLabel);
    entry.totalObjects.add(totalObjectIdentifier);
  }

  // Затем обрабатываю таблицу нарушений, чтобы посчитать динамику и статусы.
  const inspectionResultColumn = violationMapping.inspectionResult;

  const shouldApplyViolationFilter = context.violationMode === 'custom' && context.customViolations.length;

  for (const record of state.violations) {
    if (!dataSourcePredicate(record)) {
      continue;
    }
    const typeValue = getValueAsString(record[violationMapping.objectType]);
    if (typeValue && !typePredicate(typeValue)) {
      continue;
    }
    if (!typeValue && context.typeMode === 'custom') {
      continue;
    }
    const districtLabel = getValueAsString(record[violationMapping.district]);
    const objectName = getValueAsString(record[violationMapping.objectName]);
    const status = getValueAsString(record[violationMapping.status]);
    const violationName = violationNameColumn
      ? getValueAsString(record[violationNameColumn])
      : '';
    const normalizedViolationName = normalizeKey(violationName);
    const hasViolationName = Boolean(normalizedViolationName);
    const matchesViolationFilter =
      !shouldApplyViolationFilter || (hasViolationName && violationPredicate(violationName));
    const normalizedStatus = normalizeText(status);
    if (normalizedStatus === DRAFT_STATUS) {
      continue;
    }
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
    const normalizedObjectName = normalizeKey(objectName);
    const objectIdentifierCandidates = buildObjectIdentifierCandidates(
      violationObjectIdColumn ? record[violationObjectIdColumn] : '',
      violationMapping.objectName ? record[violationMapping.objectName] : '',
    );
    const primaryObjectIdentifier = getPrimaryObjectIdentifier(objectIdentifierCandidates);
    const fallbackObjectIdentifier = primaryObjectIdentifier || (normalizedObjectName ? `name:${normalizedObjectName}` : '');

    if (isWithinPeriod(inspectionDate, periods.current)) {
      if (fallbackObjectIdentifier) {
        entry.inspectedObjects.add(fallbackObjectIdentifier);
      }
      if (
        inspectionResultColumn &&
        fallbackObjectIdentifier &&
        normalizedInspectionResult === INSPECTION_RESULT_VIOLATION &&
        matchesViolationFilter
      ) {
        entry.objectsWithDetectedViolations.add(fallbackObjectIdentifier);
      }
      if (
        hasViolationName &&
        matchesViolationFilter &&
        inspectionResultColumn &&
        normalizedInspectionResult === INSPECTION_RESULT_VIOLATION
      ) {
        entry.currentViolationsCount += 1;
      }
      if (hasViolationName && matchesViolationFilter && normalizedStatus === RESOLVED_STATUS) {
        entry.resolvedCount += 1;
      }
      if (hasViolationName && matchesViolationFilter && CURRENT_CONTROL_STATUS_SET.has(normalizedStatus)) {
        entry.currentControlStatusesCount += 1;
      }
    }
    if (isWithinPeriod(inspectionDate, periods.previous)) {
      if (hasViolationName && matchesViolationFilter && CONTROL_STATUS_SET.has(normalizedStatus)) {
        entry.previousControlCount += 1;
      }
    }
  }

  const rows = [];
  const totals = {
    totalObjects: 0,
    inspectedObjects: 0,
    objectsWithDetectedViolations: 0,
    totalViolations: 0,
    currentViolations: 0,
    previousControl: 0,
    resolved: 0,
    onControl: 0,
  };

  const sortedEntries = Array.from(districtData.values()).sort(compareDistrictEntries);
  for (const entry of sortedEntries) {
    const totalObjectsCount = entry.totalObjects.size;
    const inspectedCount = entry.inspectedObjects.size;
    const objectsWithDetectedViolationsCount = entry.objectsWithDetectedViolations.size;
    const currentViolationsCount = entry.currentViolationsCount;
    const previousControlCount = entry.previousControlCount;
    const resolvedCount = entry.resolvedCount;
    const currentControlStatusesCount = entry.currentControlStatusesCount;
    const totalViolationsCount = currentViolationsCount + previousControlCount;
    const onControlTotal = previousControlCount + currentControlStatusesCount;

    const shouldSkipRow =
      state.typeMode === 'custom' &&
      totalObjectsCount === 0 &&
      inspectedCount === 0 &&
      objectsWithDetectedViolationsCount === 0 &&
      totalViolationsCount === 0 &&
      previousControlCount === 0 &&
      resolvedCount === 0 &&
      onControlTotal === 0;
    if (shouldSkipRow) {
      continue;
    }

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
      onControl: onControlTotal,
    });

    totals.totalObjects += totalObjectsCount;
    totals.inspectedObjects += inspectedCount;
    totals.objectsWithDetectedViolations += objectsWithDetectedViolationsCount;
    totals.currentViolations += currentViolationsCount;
    totals.previousControl += previousControlCount;
    totals.resolved += resolvedCount;
    totals.onControl += onControlTotal;
    totals.totalViolations += totalViolationsCount;
  }

  const totalRow = {
    label: 'ИТОГО',
    totalObjects: totals.totalObjects,
    inspectedObjects: totals.inspectedObjects,
    inspectedPercent: computePercent(totals.inspectedObjects, totals.totalObjects),
    violationPercent: computePercent(totals.objectsWithDetectedViolations, totals.inspectedObjects),
    totalViolations: totals.totalViolations,
    currentViolations: totals.currentViolations,
    previousControl: totals.previousControl,
    resolved: totals.resolved,
    onControl: totals.onControl,
  };

  return { rows, totalRow };
}

function shouldMergeWithTiNao(key) {
  return TINAO_COMBINATION_KEYS.has(key);
}

function getDistrictSortIndex(label) {
  const normalized = normalizeKey(label);
  if (DISTRICT_SORT_ORDER.has(normalized)) {
    return DISTRICT_SORT_ORDER.get(normalized);
  }
  if (normalized === 'нао') {
    const baseIndex = DISTRICT_SORT_ORDER.get(TINAO_AGGREGATED_KEY);
    return baseIndex !== undefined ? baseIndex + 0.05 : DISTRICT_SORT_ORDER.size;
  }
  if (normalized === 'тао') {
    const baseIndex = DISTRICT_SORT_ORDER.get(TINAO_AGGREGATED_KEY);
    return baseIndex !== undefined ? baseIndex + 0.1 : DISTRICT_SORT_ORDER.size;
  }
  return DISTRICT_SORT_ORDER.size + 1;
}

function compareDistrictEntries(a, b) {
  const indexA = getDistrictSortIndex(a.label);
  const indexB = getDistrictSortIndex(b.label);
  if (indexA !== indexB) {
    return indexA - indexB;
  }
  return a.label.localeCompare(b.label, 'ru');
}

/**
 * Создаёт HTML-таблицу предпросмотра, придерживаясь структуры thead/tbody.
 * Одновременно формирует подпись таблицы с описанием диапазона и источника данных.
 *
 * @param {{rows: Array<object>, totalRow: object}} report Агрегированные строки отчёта.
 * @param {{current: {start: Date, end: Date}, previous: {start: Date, end: Date}}} periods Выбранные периоды.
 */
function renderReportTable(report, periods, overrides = {}) {
  const context = resolveReportContext(overrides);
  const headers = buildTableHeaders(periods, context);
  const table = elements.reportTable;
  table.innerHTML = '';

  const caption = document.createElement('caption');
  caption.textContent = buildTableCaption(report, periods, context);
  table.append(caption);

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
  tbody.append(createRowElement(report.totalRow, { isTotal: true }));
  for (const row of report.rows) {
    tbody.append(createRowElement(row));
  }
  table.append(tbody);
}

/**
 * Создаёт строку таблицы на основе агрегированной записи отчёта.
 *
 * @param {object} row Данные строки отчёта.
 * @param {{isTotal?: boolean}} [options] Дополнительные настройки строки.
 * @returns {HTMLTableRowElement} DOM-элемент строки.
 */
function createRowElement(row, options = {}) {
  const { isTotal = false } = options;
  const tr = document.createElement('tr');
  if (isTotal) {
    tr.classList.add('table-row--total');
  }
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

/**
 * Создаёт ячейку таблицы с текстовым содержимым.
 *
 * @param {string} value Значение внутри ячейки.
 * @param {boolean} [isHeader=false] true, если требуется ячейка th.
 * @returns {HTMLTableCellElement} DOM-элемент ячейки.
 */
function createCell(value, isHeader = false) {
  const cell = document.createElement(isHeader ? 'th' : 'td');
  cell.textContent = value ?? '';
  return cell;
}

// Формирую заголовки таблицы, включая динамический текст по выбранному периоду.
function buildTableHeaders(periods, context = resolveReportContext()) {
  const totalHeader = buildTotalHeader(context);
  const rangeHeader = `Проверено объектов контроля с  ${formatDateDisplay(periods.current.start)} по ${formatDateDisplay(periods.current.end)}`;
  return [
    'Округ',
    totalHeader,
    rangeHeader,
    '% проверенных объектов от общего количества объектов',
    '% объектов с нарушениями',
    'Всего нарушений',
    'Нарушения, выявленные за отчётный период',
    'Нарушения, находящиеся на контроле с предыдущей проверки',
    'Устранено нарушений',
    'Нарушения на контроле',
  ];
}

/**
 * Формирует подпись для таблицы предпросмотра с описанием периода, источника данных и количества строк.
 *
 * @param {{rows: Array<object>}} report Агрегированные данные отчёта.
 * @param {{current: {start: Date, end: Date}, previous: {start: Date, end: Date}}} periods Выбранные периоды.
 * @returns {string} Текст подписи для таблицы.
 */
function buildTableCaption(report, periods, context = resolveReportContext()) {
  const currentRange = formatPeriodRange(periods.current);
  const rowsDescription = `Строк в таблице: ${numberFormatter.format(report.rows.length)}.`;
  const dataSourceDescription = describeSelectedDataSource(context.dataSourceOption);
  const base = currentRange ? `Отчёт за период ${currentRange}.` : 'Отчётный период не выбран.';
  return `${base} ${dataSourceDescription}. ${rowsDescription}`.trim();
}

/**
 * Возвращает название столбца «Всего» с учётом выбранных типов объектов.
 *
 * @returns {string} Подпись столбца.
 */
function buildTotalHeader(context = resolveReportContext()) {
  if (context.typeMode === 'all' || !context.customTypes.length) {
    return 'Всего ОДХ';
  }
  if (context.customTypes.length === 1) {
    return `Всего ${context.customTypes[0]}`;
  }
  return `Всего (${context.customTypes.join(', ')})`;
}

/**
 * Формирует строку с номерами показателей для выгрузки в Excel.
 *
 * @param {number} columnCount Количество колонок в таблице.
 * @returns {Array<string | number>} Список обозначений.
 */
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

function buildViolationSelectionDescription(context = resolveReportContext()) {
  if (context.violationMode === 'custom' && context.customViolations.length) {
    const seen = new Set();
    const names = [];
    for (const value of context.customViolations) {
      const text = getValueAsString(value);
      const key = normalizeKey(text);
      if (!key || seen.has(key)) {
        continue;
      }
      seen.add(key);
      names.push(text);
    }
    if (names.length) {
      return `Наименования нарушений: ${names.join(', ')}`;
    }
  }
  return 'Наименования нарушений: все (фильтр не применялся).';
}

/**
 * Возвращает описание выбранных типов объектов для заголовка Excel.
 *
 * @param {{typeMode: string, customTypes: Array<string>}} context Контекст отчёта.
 * @returns {string} Текст для подстановки в заголовок.
 */
function describeTypeSelectionForTitle(context = resolveReportContext()) {
  if (context.typeMode !== 'custom' || !Array.isArray(context.customTypes) || !context.customTypes.length) {
    return 'ОДХ';
  }
  const uniqueTypes = [];
  const seen = new Set();
  for (const value of context.customTypes) {
    const text = getValueAsString(value);
    const key = normalizeKey(text);
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    uniqueTypes.push(text);
  }
  if (!uniqueTypes.length) {
    return 'ОДХ';
  }
  if (uniqueTypes.length === 1) {
    return uniqueTypes[0];
  }
  return uniqueTypes.join(', ');
}

/**
 * Собирает заголовок листа Excel с учётом выбранного периода и источника данных.
 *
 * @param {{current?: {start: Date, end: Date}}} periods Выбранные периоды.
 * @returns {string} Заголовок листа Excel.
 */
function buildExcelTitle(periods, context = resolveReportContext(), options = {}) {
  const { titlePrefix } = options;
  const periodRange = formatTitlePeriod(periods?.current);
  const defaultPrefix = `Нарушения на ${describeTypeSelectionForTitle(context)}`;
  const parts = [titlePrefix || defaultPrefix];
  if (periodRange) {
    parts[0] += ` (отчёт за ${periodRange})`;
  }
  const dataSource = describeDataSourceForTitle(context.dataSourceOption);
  if (dataSource) {
    parts.push(dataSource);
  }
  return parts.join(', ');
}

/**
 * Возвращает текстовое описание источников данных для заголовка Excel.
 *
 * @returns {string} Фраза для заголовка.
 */
function describeDataSourceForTitle(option = state.dataSourceOption) {
  switch (option) {
    case 'oati':
      return 'накопленные только ОАТИ';
    case 'cafap':
      return 'накопленные только ЦАФАП';
    default:
      return 'выявленные ОАТИ и ЦАФАП';
  }
}

/**
 * Удаляет все строки из таблицы предпросмотра.
 */
function clearTable() {
  elements.reportTable.innerHTML = '';
}

/**
 * Отображает текстовое сообщение под таблицей предпросмотра.
 *
 * @param {string} message Текст сообщения для пользователя.
 */
function showPreviewMessage(message) {
  elements.previewMessage.textContent = message;
}

/**
 * Проверяет, что значение является корректным объектом Date.
 *
 * @param {unknown} value Проверяемое значение.
 * @returns {boolean} true, если дата валидна.
 */
function isValidDate(value) {
  return value instanceof Date && !Number.isNaN(value.getTime());
}

/**
 * Преобразует ISO-строку в объект Date без учёта временной части.
 *
 * @param {string} iso Строка формата YYYY-MM-DD.
 * @returns {Date | null} Объект даты или null при ошибке.
 */
function parseIsoDate(iso) {
  if (typeof iso !== 'string') {
    return null;
  }
  const trimmed = iso.trim();
  if (!trimmed) {
    return null;
  }
  const match = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!match) {
    return null;
  }
  const [, yearText, monthText, dayText] = match;
  const year = Number.parseInt(yearText, 10);
  const month = Number.parseInt(monthText, 10);
  const day = Number.parseInt(dayText, 10);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
    return null;
  }
  const date = new Date(year, month - 1, day);
  if (!isValidDate(date) || date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return null;
  }
  return date;
}

/**
 * Возвращает предикат для фильтрации записей по типам объектов.
 *
 * @returns {(value: string) => boolean} Функция-предикат.
 */
function createTypePredicate(context = resolveReportContext()) {
  if (context.typeMode === 'all' || !context.customTypes.length) {
    return () => true;
  }
  const allowed = new Set(context.customTypes.map((value) => normalizeKey(value)));
  return (value) => allowed.has(normalizeKey(value));
}

/**
 * Возвращает предикат для фильтрации по наименованиям нарушений.
 *
 * @returns {(value: string) => boolean} Функция-предикат.
 */
function createViolationPredicate(context = resolveReportContext()) {
  if (context.violationMode === 'all' || !context.customViolations.length) {
    return () => true;
  }
  const allowed = new Set(context.customViolations.map((value) => normalizeKey(value)));
  return (value) => allowed.has(normalizeKey(value));
}

/**
 * Возвращает предикат для фильтрации записей по источнику данных.
 *
 * @returns {(record: object) => boolean} Функция-предикат.
 */
function createDataSourcePredicate(context = resolveReportContext()) {
  const column = state.violationMapping.dataSource;
  if (!column) {
    return () => true;
  }
  const mode = context.dataSourceOption;
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

/**
 * Унифицированная нормализация строк: обрезает пробелы и приводит к нижнему регистру.
 *
 * @param {unknown} value Значение для нормализации.
 * @returns {string} Нормализованный ключ.
 */
function normalizeKey(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return value.toString().trim().toLowerCase();
}

/**
 * Обёртка над normalizeKey для лучшей читаемости вызовов выше по коду.
 *
 * @param {unknown} value Значение для нормализации.
 * @returns {string} Нормализованное значение.
 */
function normalizeText(value) {
  return normalizeKey(value);
}

/**
 * Преобразует значение в строку, пригодную для отображения в интерфейсе или логах.
 *
 * @param {unknown} value Исходное значение.
 * @returns {string} Подготовленная строка.
 */
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

/**
 * Определяет категорию источника данных по тексту ячейки.
 *
 * @param {unknown} value Значение из колонки источника данных.
 * @returns {'oati' | 'cafap' | 'both' | ''} Выбранная категория или пустая строка.
 */
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

/**
 * Извлекает уникальные даты из указанной колонки и возвращает их в отсортированном виде.
 *
 * @param {object[]} records Массив записей.
 * @param {string} column Название колонки с датами.
 * @returns {Array<{iso: string, date: Date}>} Упорядоченный массив дат.
 */
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
    if (!iso || seen.has(iso)) {
      continue;
    }
    seen.add(iso);
    dates.push({ iso, date });
  }
  dates.sort((a, b) => a.iso.localeCompare(b.iso));
  return dates;
}

/**
 * Конвертирует значение из Excel в объект Date, поддерживая число, строку или Date.
 *
 * @param {unknown} value Значение ячейки.
 * @returns {Date | null} Объект даты или null, если преобразование невозможно.
 */
function parseDateValue(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }
  if (value instanceof Date) {
    return isValidDate(value) ? new Date(value.getTime()) : null;
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    const parsed = XLSX.SSF?.parse_date_code?.(value);
    if (!parsed) {
      return null;
    }
    const date = new Date(parsed.y, parsed.m - 1, parsed.d);
    return isValidDate(date) ? date : null;
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return null;
    }
    // Поддерживаю формат «дд.мм.гггг» с необязательным временем.
    const dateTimeMatch = trimmed.match(/^(\d{1,2})[.](\d{1,2})[.](\d{4})(?:\s+(\d{1,2}):(\d{2}))?$/);
    if (dateTimeMatch) {
      const [, dayRaw, monthRaw, yearRaw] = dateTimeMatch;
      const day = Number.parseInt(dayRaw, 10);
      const month = Number.parseInt(monthRaw, 10);
      const year = Number.parseInt(yearRaw, 10);
      if (Number.isFinite(day) && Number.isFinite(month) && Number.isFinite(year)) {
        const date = new Date(year, month - 1, day);
        if (
          isValidDate(date)
          && date.getFullYear() === year
          && date.getMonth() === month - 1
          && date.getDate() === day
        ) {
          return date;
        }
      }
      return null;
    }
    const isoDate = Date.parse(trimmed);
    if (Number.isFinite(isoDate)) {
      const date = new Date(isoDate);
      return isValidDate(date) ? date : null;
    }
  }
  return null;
}

/**
 * Форматирует объект Date в ISO-строку без времени (YYYY-MM-DD).
 *
 * @param {Date} date Объект даты.
 * @returns {string} ISO-строка.
 */
function formatIsoDate(date) {
  if (!isValidDate(date)) {
    return '';
  }
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Форматирует дату в строку вида «дд.мм.гггг».
 *
 * @param {Date} date Объект даты.
 * @returns {string} Отформатированная строка.
 */
function formatDateDisplay(date) {
  if (!isValidDate(date)) {
    return '';
  }
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
}

/**
 * Форматирует дату в короткий вид «дд.мм».
 *
 * @param {Date} date Объект даты.
 * @returns {string} Строка без года.
 */
function formatDateDisplayShort(date) {
  if (!isValidDate(date)) {
    return '';
  }
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  return `${day}.${month}`;
}

/**
 * Возвращает список уникальных значений из указанной колонки, сохраняя оригинальное написание.
 *
 * @param {object[]} records Набор записей.
 * @param {string} column Название колонки.
 * @returns {string[]} Уникальные значения.
 */
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

/**
 * Проверяет, попадает ли дата в диапазон (без учёта времени).
 *
 * @param {Date} date Проверяемая дата.
 * @param {{start: Date, end: Date}} period Диапазон дат.
 * @returns {boolean} true, если дата внутри диапазона.
 */
function isWithinPeriod(date, period) {
  if (!isValidDate(date) || !period || !isValidDate(period.start) || !isValidDate(period.end)) {
    return false;
  }
  const dateOnly = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  return dateOnly >= period.start && dateOnly <= period.end;
}

/**
 * Рассчитывает процентное значение с защитой от деления на ноль.
 *
 * @param {number} part Числитель.
 * @param {number} total Знаменатель.
 * @returns {number} Процент.
 */
function computePercent(part, total) {
  if (!total || !Number.isFinite(part)) {
    return 0;
  }
  return (part / total) * 100;
}

/**
 * Форматирует целое число с пробелами в качестве разделителей тысяч.
 *
 * @param {number} value Число.
 * @returns {string} Отформатированная строка.
 */
function formatInteger(value) {
  if (!Number.isFinite(value)) {
    return '0';
  }
  return numberFormatter.format(Math.round(value));
}

/**
 * Форматирует процент с точностью до одного десятичного знака.
 *
 * @param {number} value Значение в процентах.
 * @returns {string} Строка с процентом.
 */
function formatPercent(value) {
  if (!Number.isFinite(value)) {
    return '0';
  }
  return value.toFixed(1);
}

/**
 * Генерирует Excel-файл на основе агрегированного отчёта и инициирует скачивание пользователю.
 *
 * @param {{rows: Array<object>, totalRow: object}} report Данные отчёта.
 * @param {{current: {start: Date, end: Date}, previous: {start: Date, end: Date}}} periods Периоды отчёта.
 */
function createWorksheetForReport(report, periods, options = {}) {
  const { context: overrides = {}, sheetName = 'Отчёт', titlePrefix } = options;
  const context = resolveReportContext(overrides);
  const headers = buildTableHeaders(periods, context);
  const columnCount = headers.length;
  const headerNumbers = buildExcelHeaderNumbers(columnCount);
  const title = buildExcelTitle(periods, context, { titlePrefix });
  const violationDescription = buildViolationSelectionDescription(context);

  const createEmptyRow = () => Array.from({ length: columnCount }, () => null);
  const aoa = [];

  aoa.push(createEmptyRow());
  const titleRowIndex = aoa.length;
  aoa.push([title, ...Array(Math.max(0, columnCount - 1)).fill(null)]);
  let descriptionRowIndex = null;
  if (violationDescription) {
    descriptionRowIndex = aoa.length;
    aoa.push([violationDescription, ...Array(Math.max(0, columnCount - 1)).fill(null)]);
  }
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
  const merges = [
    { s: { r: titleRowIndex, c: 0 }, e: { r: titleRowIndex, c: columnCount - 1 } },
  ];
  if (descriptionRowIndex !== null) {
    merges.push({ s: { r: descriptionRowIndex, c: 0 }, e: { r: descriptionRowIndex, c: columnCount - 1 } });
  }
  sheet['!merges'] = merges;

  const baseColumnWidths = [26, 16, 22, 28, 22, 20, 34, 34, 24, 24];
  sheet['!cols'] = Array.from({ length: columnCount }, (_, index) => ({
    wch: baseColumnWidths[index] ?? 20,
  }));

  sheet['!rows'] = sheet['!rows'] ?? [];
  sheet['!rows'][titleRowIndex] = { hpt: 26 };
  if (descriptionRowIndex !== null) {
    sheet['!rows'][descriptionRowIndex] = { hpt: 20 };
  }
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
  const descriptionStyle = {
    font: { ...baseFont, sz: 11 },
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    fill: { fgColor: { rgb: 'FFF8FAFC' } },
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
    if (descriptionRowIndex !== null) {
      setCellStyle(sheet, descriptionRowIndex, colIndex, descriptionStyle);
    }
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

  return { sheet, sheetName: sanitizeWorksheetName(sheetName) };
}

function exportReportToExcel(report, periods) {
  try {
    const { sheet, sheetName } = createWorksheetForReport(report, periods);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
    const fileName = buildExportFileName(periods);
    XLSX.writeFile(workbook, fileName);
  } catch (error) {
    console.error('Не удалось выгрузить таблицу.', error);
    showPreviewMessage('Не удалось выгрузить таблицу. Повторите попытку.');
  }
}

function sanitizeWorksheetName(name) {
  const fallback = 'Отчёт';
  const raw = getValueAsString(name) || fallback;
  const cleaned = raw.replace(/[\\/?*\[\]:]+/g, '_').trim();
  if (!cleaned) {
    return fallback;
  }
  if (cleaned.length > 31) {
    return cleaned.slice(0, 31);
  }
  return cleaned;
}

/**
 * Преобразует строку отчёта в массив значений для записи в Excel.
 *
 * @param {object} row Строка отчёта.
 * @returns {Array<string | number>} Значения, готовые к записи в XLSX.
 */
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

/**
 * Гарантирует числовое значение для Excel, заменяя NaN и Infinity на ноль.
 *
 * @param {unknown} value Исходное значение.
 * @returns {number} Безопасное число.
 */
function safeNumber(value) {
  if (!Number.isFinite(value)) {
    return 0;
  }
  return Number(value);
}

/**
 * Применяет набор стилевых объектов к конкретной ячейке листа.
 *
 * @param {XLSX.WorkSheet} sheet Лист Excel.
 * @param {number} rowIndex Индекс строки.
 * @param {number} colIndex Индекс столбца.
 * @param {...object} styles Стилевые объекты для объединения.
 */
function setCellStyle(sheet, rowIndex, colIndex, ...styles) {
  const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const cell = sheet[address];
  if (!cell) {
    return;
  }
  cell.s = mergeStyles(cell.s ?? {}, ...styles);
}

/**
 * Устанавливает числовой формат ячейки (например, проценты или целые числа).
 *
 * @param {XLSX.WorkSheet} sheet Лист Excel.
 * @param {number} rowIndex Индекс строки.
 * @param {number} colIndex Индекс столбца.
 * @param {string} format Формат для XLSX.
 */
function setNumberFormat(sheet, rowIndex, colIndex, format) {
  const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  const cell = sheet[address];
  if (!cell) {
    return;
  }
  cell.z = format;
}

/**
 * Глубоко объединяет несколько стилевых объектов XLSX.
 *
 * @param {...object} styles Стилевые объекты.
 * @returns {object} Объединённый стиль.
 */
function mergeStyles(...styles) {
  const result = {};
  for (const style of styles) {
    if (style) {
      deepMerge(result, style);
    }
  }
  return result;
}

/**
 * Возвращает человекочитаемое описание выбранных источников данных.
 *
 * @returns {string} Описание источника.
 */
function describeSelectedDataSource(option = state.dataSourceOption) {
  switch (option) {
    case 'oati':
      return 'Источник данных: ОАТИ';
    case 'cafap':
      return 'Источник данных: ЦАФАП';
    default:
      return 'Источник данных: ОАТИ и ЦАФАП';
  }
}

/**
 * Рекурсивно объединяет свойства объектов без использования внешних зависимостей.
 *
 * @param {object} target Целевой объект.
 * @param {object} source Источник свойств.
 * @returns {object} Изменённый целевой объект.
 */
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

/**
 * Возвращает объект с настройками рамки для ячейки Excel.
 *
 * @param {string} color Цвет рамки.
 * @returns {object} Объект настроек рамки.
 */
function buildBorderStyle(color) {
  return {
    top: { style: 'thin', color: { rgb: color } },
    bottom: { style: 'thin', color: { rgb: color } },
    left: { style: 'thin', color: { rgb: color } },
    right: { style: 'thin', color: { rgb: color } },
  };
}

/**
 * Формирует строку диапазона дат для описательной части отчёта.
 *
 * @param {{start: Date, end: Date}} period Диапазон дат.
 * @returns {string} Строка с датами.
 */
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

/**
 * Формирует компактное представление диапазона для заголовков.
 *
 * @param {{start: Date, end: Date}} period Диапазон дат.
 * @returns {string} Краткое описание периода.
 */
function formatTitlePeriod(period) {
  if (!period) {
    return '';
  }
  const sameYear =
    isValidDate(period.start) &&
    isValidDate(period.end) &&
    period.start.getFullYear() === period.end.getFullYear();
  const start = sameYear ? formatDateDisplayShort(period.start) : formatDateDisplay(period.start);
  const end = sameYear ? formatDateDisplayShort(period.end) : formatDateDisplay(period.end);
  if (!start || !end) {
    return '';
  }
  return `${start} — ${end}`;
}

/**
 * Подготавливает дату для включения в имя файла.
 *
 * @param {Date} date Объект даты.
 * @returns {string} Строка вида YYYY-MM-DD.
 */
function formatDateForFileName(date) {
  if (!isValidDate(date)) {
    return '0000-00-00';
  }
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Генерирует безопасное имя файла для выгрузки отчёта.
 *
 * @param {{current: {start: Date, end: Date}}} periods Диапазон отчётного периода.
 * @returns {string} Имя файла.
 */
function buildExportFileName(periods) {
  const start = formatDateForFileName(periods.current.start);
  const end = formatDateForFileName(periods.current.end);
  const sanitized = `Отчёт ОАТИ ${start}_${end}.xlsx`.replace(/[\\/:*?"<>|]+/g, '_');
  return sanitized;
}

initializeBulkExportFeature();
