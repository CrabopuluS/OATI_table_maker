const STORAGE_KEY = 'feline-vitae-archive-v1';

const defaultArchive = Object.freeze({
  cats: [
    {
      id: 'aurora',
      name: 'Аврора',
      breed: 'Невская маскарадная',
      birthdate: '2019-04-18',
      accentColor: '#d2a45b',
      notes: 'Плановые осмотры дважды в год. Последняя вакцинация проведена в ноябре 2024 года.',
      documents: [
        {
          id: 'aurora-lab-2024',
          title: 'Результаты биохимического анализа',
          description: 'Расширенный профиль крови, показатели в пределах нормы. Рекомендовано повторить через 6 месяцев.',
          documentDate: '2024-11-02',
          uploadedAt: '2024-11-03T09:15:00.000Z',
          fileName: 'aurora_biochemistry.txt',
          fileType: 'text/plain',
          size: 112,
          dataUrl: 'data:text/plain;base64,0JjQvdGC0L7QsdC40YLQtdC70LjQutC+INC/0YDQvtGB0YLRjCDQutC+0LzQtdCy0YPRgNCwINCx0L7QtNC10YLRjCDQv9GA0L7RgdGC0LXRgNC+0YDQvtCyLCDQv9C+0LPQviDQvtGC0L7Qs9C+INC/0YDQvtGB0YLRjCDQutC+0LzQtdCy0LjRj9GC0LXRgNC+0LIu',
        },
      ],
    },
    {
      id: 'miro',
      name: 'Миро',
      breed: 'Абиссинская',
      birthdate: '2021-02-11',
      accentColor: '#9b7ad3',
      notes: 'Повышенная активность, требуется контроль веса и регулярная проверка суставов.',
      documents: [
        {
          id: 'miro-dental-2025',
          title: 'Отчет стоматологического осмотра',
          description: 'Профилактическая чистка зубов и осмотр десен. Назначен повторный визит через 9 месяцев.',
          documentDate: '2025-01-14',
          uploadedAt: '2025-01-14T18:32:00.000Z',
          fileName: 'miro_dental_summary.txt',
          fileType: 'text/plain',
          size: 126,
          dataUrl: 'data:text/plain;base64,0KHQvtGA0LPQvtCy0LjQt9C40YHRgtCwINC40Lcg0LLQtdC30L3QsNGPINGB0YLRgNC+0L3QsNC70YzQvdGL0Lkg0YHQvtC+0YHRgtC+0L3QsCDQv9C+0LvRjNC90YvRhSDQsiDQtNCw0LrRgtC+0LLRgdGC0LLQvi4=',
        },
      ],
    },
  ],
  selectedCatId: 'aurora',
  lastUpdated: '2025-01-14T18:32:00.000Z',
});

const state = {
  archive: loadArchive(),
  filters: {
    catSearch: '',
    catFilter: 'all',
    documentSearch: '',
    documentFilter: 'all',
  },
  selectedDocumentId: null,
};

const hasCrypto = typeof crypto !== 'undefined' && typeof crypto.getRandomValues === 'function';

const elements = {
  catGrid: document.getElementById('cat-grid'),
  catEmpty: document.getElementById('cat-empty'),
  catSearch: document.getElementById('cat-search'),
  catFilter: document.getElementById('cat-filter'),
  heroCatsCount: document.getElementById('hero-cats-count'),
  heroDocsCount: document.getElementById('hero-docs-count'),
  heroLastUpdate: document.getElementById('hero-last-update'),
  documentTableBody: document.getElementById('document-table-body'),
  documentEmpty: document.getElementById('document-empty'),
  documentSearch: document.getElementById('document-search'),
  documentFilter: document.getElementById('document-filter'),
  detail: document.getElementById('document-detail'),
  documentsSubtitle: document.getElementById('documents-subtitle'),
  openAdmin: document.getElementById('open-admin'),
  adminModal: document.getElementById('admin-modal'),
  closeAdmin: document.getElementById('close-admin'),
  securityModal: document.getElementById('security-modal'),
  securityBriefing: document.getElementById('security-briefing'),
  footerYear: document.getElementById('footer-year'),
  adminForm: document.getElementById('admin-form'),
  catNameInput: document.getElementById('cat-name'),
  catBreedInput: document.getElementById('cat-breed'),
  catBirthdateInput: document.getElementById('cat-birthdate'),
  catColorInput: document.getElementById('cat-color'),
  createCatButton: document.getElementById('create-cat'),
  documentCatSelect: document.getElementById('document-cat'),
  documentTitleInput: document.getElementById('document-title'),
  documentDateInput: document.getElementById('document-date'),
  documentFileInput: document.getElementById('document-file'),
  documentDescriptionInput: document.getElementById('document-description'),
  uploadDocumentButton: document.getElementById('upload-document'),
  exportArchiveButton: document.getElementById('export-archive'),
  importArchiveInput: document.getElementById('import-archive'),
};

initialize();

function initialize() {
  updateFooterYear();
  syncArchiveWithStorage();
  wireControls();
  render();
}

function wireControls() {
  elements.catSearch.addEventListener('input', (event) => {
    state.filters.catSearch = sanitizeText(event.target.value);
    renderCatGrid();
  });

  elements.catFilter.addEventListener('change', (event) => {
    state.filters.catFilter = event.target.value;
    renderCatGrid();
  });

  elements.documentSearch.addEventListener('input', (event) => {
    state.filters.documentSearch = sanitizeText(event.target.value);
    renderDocumentTable();
  });

  elements.documentFilter.addEventListener('change', (event) => {
    state.filters.documentFilter = event.target.value;
    renderDocumentTable();
  });

  elements.openAdmin.addEventListener('click', () => {
    refreshAdminCatList();
    elements.adminModal.showModal();
  });

  elements.closeAdmin.addEventListener('click', () => {
    elements.adminModal.close();
  });

  elements.createCatButton.addEventListener('click', handleCreateCat);
  elements.uploadDocumentButton.addEventListener('click', handleUploadDocument);

  elements.securityBriefing.addEventListener('click', () => {
    elements.securityModal.showModal();
  });

  elements.adminModal.addEventListener('cancel', () => {
    elements.adminForm.reset();
  });

  elements.exportArchiveButton.addEventListener('click', handleExportArchive);
  elements.importArchiveInput.addEventListener('change', handleImportArchive);
}

function render() {
  renderCatGrid();
  renderDocumentTable();
  renderHeroStatistics();
}

function renderCatGrid() {
  const cats = getFilteredCats();
  elements.catGrid.innerHTML = '';

  if (!cats.length) {
    elements.catEmpty.hidden = false;
    return;
  }

  elements.catEmpty.hidden = true;

  cats.forEach((cat) => {
    const card = document.createElement('article');
    card.className = 'cat-card';
    card.setAttribute('role', 'listitem');
    card.setAttribute('tabindex', '0');
    card.dataset.catId = cat.id;
    card.setAttribute('aria-selected', String(cat.id === state.archive.selectedCatId));

    card.addEventListener('click', () => selectCat(cat.id));
    card.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        selectCat(cat.id);
      }
    });

    const name = document.createElement('h3');
    name.className = 'cat-card__name';
    name.textContent = cat.name;

    const meta = document.createElement('p');
    meta.className = 'cat-card__meta';
    meta.textContent = formatCatMeta(cat);

    const status = document.createElement('p');
    status.className = 'cat-card__status';
    status.style.color = cat.accentColor || 'var(--accent-strong)';

    const indicator = document.createElement('span');
    indicator.className = 'cat-card__indicator';
    indicator.style.backgroundColor = cat.accentColor || 'var(--accent-strong)';

    const statusText = document.createElement('span');
    statusText.textContent = cat.documents.length ? 'Досье активно' : 'Нет документов';

    status.append(indicator, statusText);

    const notes = document.createElement('p');
    notes.className = 'cat-card__notes';
    notes.textContent = cat.notes || 'Добавьте комментарий через меню администрирования.';

    card.append(name, meta, status, notes);
    elements.catGrid.append(card);
  });
}

function renderDocumentTable() {
  const cat = state.archive.cats.find(({ id }) => id === state.archive.selectedCatId);
  const rows = [];

  if (!cat) {
    elements.documentTableBody.innerHTML = '';
    elements.documentEmpty.hidden = false;
    elements.documentsSubtitle.textContent = 'Выберите питомца, чтобы просмотреть доступные материалы.';
    state.selectedDocumentId = null;
    renderDocumentDetail();
    return;
  }

  elements.documentsSubtitle.textContent = `Документы, связанные с питомцем ${cat.name}.`;

  const documents = getFilteredDocuments(cat.documents);

  if (!documents.length) {
    elements.documentTableBody.innerHTML = '';
    elements.documentEmpty.hidden = false;
    state.selectedDocumentId = null;
    renderDocumentDetail();
    return;
  }

  elements.documentEmpty.hidden = true;

  if (!documents.some((documentRecord) => documentRecord.id === state.selectedDocumentId)) {
    state.selectedDocumentId = documents[0].id;
  }

  documents.forEach((documentRecord) => {
    const row = document.createElement('tr');

    const titleCell = document.createElement('td');
    titleCell.textContent = documentRecord.title;

    const descriptionCell = document.createElement('td');
    descriptionCell.textContent = documentRecord.description || '—';

    const dateCell = document.createElement('td');
    dateCell.textContent = formatDocumentDate(documentRecord.documentDate);

    const sizeCell = document.createElement('td');
    sizeCell.textContent = formatFileSize(documentRecord.size);

    const actionsCell = document.createElement('td');
    actionsCell.className = 'document-actions';

    const viewButton = document.createElement('button');
    viewButton.type = 'button';
    viewButton.className = 'button button--outline';
    viewButton.textContent = 'Детали';
    viewButton.addEventListener('click', () => {
      state.selectedDocumentId = documentRecord.id;
      renderDocumentDetail();
    });

    const downloadButton = document.createElement('button');
    downloadButton.type = 'button';
    downloadButton.className = 'button';
    downloadButton.textContent = 'Скачать';
    downloadButton.addEventListener('click', () => downloadDocument(documentRecord));

    actionsCell.append(viewButton, downloadButton);
    row.append(titleCell, descriptionCell, dateCell, sizeCell, actionsCell);
    rows.push(row);
  });

  elements.documentTableBody.replaceChildren(...rows);
  renderDocumentDetail();
}

function renderDocumentDetail() {
  const cat = state.archive.cats.find(({ id }) => id === state.archive.selectedCatId);
  const documentRecord = cat?.documents.find(({ id }) => id === state.selectedDocumentId);

  const detailRows = elements.detail.querySelectorAll('dd');
  if (!documentRecord) {
    detailRows[0].textContent = '—';
    detailRows[1].textContent = '—';
    detailRows[2].textContent = '—';
    detailRows[3].textContent = '—';
    detailRows[4].textContent = 'Выберите документ, чтобы увидеть подробности.';
    return;
  }

  detailRows[0].textContent = documentRecord.title;
  detailRows[1].textContent = formatDocumentDate(documentRecord.documentDate);
  detailRows[2].textContent = formatTimestamp(documentRecord.uploadedAt);
  detailRows[3].textContent = `${documentRecord.fileType || '—'} · ${formatFileSize(documentRecord.size)}`;
  detailRows[4].textContent = documentRecord.description || 'Описание отсутствует.';
}

function renderHeroStatistics() {
  const totalCats = state.archive.cats.length;
  const totalDocuments = state.archive.cats.reduce((sum, cat) => sum + cat.documents.length, 0);
  elements.heroCatsCount.textContent = String(totalCats);
  elements.heroDocsCount.textContent = String(totalDocuments);
  elements.heroLastUpdate.textContent = formatTimestamp(state.archive.lastUpdated) || '—';
}

function selectCat(catId) {
  if (state.archive.selectedCatId === catId) {
    return;
  }

  state.archive.selectedCatId = catId;
  state.selectedDocumentId = null;
  saveArchive();
  render();
}

function handleCreateCat() {
  const name = sanitizeText(elements.catNameInput.value);
  const breed = sanitizeText(elements.catBreedInput.value);
  const birthdate = elements.catBirthdateInput.value || '';
  const accentColor = elements.catColorInput.value || '#d2a45b';

  if (!name) {
    window.alert('Введите имя питомца.');
    return;
  }

  const newCat = {
    id: generateId(name),
    name,
    breed,
    birthdate,
    accentColor,
    notes: '',
    documents: [],
  };

  state.archive.cats.push(newCat);
  state.archive.selectedCatId = newCat.id;
  state.selectedDocumentId = null;
  updateLastUpdated();
  saveArchive();
  elements.adminForm.reset();
  refreshAdminCatList();
  elements.documentCatSelect.value = newCat.id;
  render();
}

function handleUploadDocument() {
  const catId = elements.documentCatSelect.value;
  const title = sanitizeText(elements.documentTitleInput.value);
  const documentDate = elements.documentDateInput.value;
  const description = sanitizeText(elements.documentDescriptionInput.value);
  const file = elements.documentFileInput.files?.[0];

  if (!catId) {
    window.alert('Выберите питомца, к которому относится документ.');
    return;
  }

  if (!title || !documentDate || !file) {
    window.alert('Заполните все обязательные поля и выберите файл.');
    return;
  }

  if (file.size > 25 * 1024 * 1024) {
    window.alert('Размер файла превышает 25 МБ. Сократите документ перед загрузкой.');
    return;
  }

  const reader = new FileReader();
  reader.onload = () => {
    const dataUrl = typeof reader.result === 'string' ? reader.result : '';
    if (!dataUrl) {
      window.alert('Не удалось прочитать файл.');
      return;
    }

    const documentRecord = {
      id: generateId(`${catId}-${title}-${Date.now()}`),
      title,
      description,
      documentDate,
      uploadedAt: new Date().toISOString(),
      fileName: file.name,
      fileType: file.type || 'application/octet-stream',
      size: file.size,
      dataUrl,
    };

    const cat = state.archive.cats.find(({ id }) => id === catId);
    if (!cat) {
      window.alert('Не удалось найти карточку питомца.');
      return;
    }

    cat.documents.push(documentRecord);
    state.archive.selectedCatId = catId;
    state.selectedDocumentId = documentRecord.id;
    updateLastUpdated();
    saveArchive();
    elements.adminForm.reset();
    refreshAdminCatList();
    elements.documentCatSelect.value = catId;
    render();
  };

  reader.onerror = () => {
    window.alert('Произошла ошибка при чтении файла.');
  };

  reader.readAsDataURL(file);
}

function downloadDocument(documentRecord) {
  const link = document.createElement('a');
  link.href = documentRecord.dataUrl;
  link.download = documentRecord.fileName || `${documentRecord.title}.dat`;
  document.body.append(link);
  link.click();
  link.remove();
}

function refreshAdminCatList() {
  elements.documentCatSelect.innerHTML = '';
  state.archive.cats.forEach((cat) => {
    const option = document.createElement('option');
    option.value = cat.id;
    option.textContent = cat.name;
    if (cat.id === state.archive.selectedCatId) {
      option.selected = true;
    }
    elements.documentCatSelect.append(option);
  });
}

function updateFooterYear() {
  elements.footerYear.textContent = new Date().getFullYear().toString();
}

function getFilteredCats() {
  const search = state.filters.catSearch.toLowerCase();
  return state.archive.cats
    .filter((cat) => {
      if (!search) {
        return true;
      }
      return (
        cat.name.toLowerCase().includes(search) ||
        (cat.breed && cat.breed.toLowerCase().includes(search))
      );
    })
    .filter((cat) => {
      if (state.filters.catFilter === 'active') {
        return cat.documents.length > 0;
      }
      if (state.filters.catFilter === 'archived') {
        return cat.documents.length === 0;
      }
      return true;
    });
}

function getFilteredDocuments(documents) {
  const search = state.filters.documentSearch.toLowerCase();
  return documents
    .filter((documentRecord) => {
      if (!search) {
        return true;
      }
      return (
        documentRecord.title.toLowerCase().includes(search) ||
        (documentRecord.description && documentRecord.description.toLowerCase().includes(search))
      );
    })
    .filter((documentRecord) => {
      switch (state.filters.documentFilter) {
        case 'pdf':
          return documentRecord.fileType.includes('pdf');
        case 'image':
          return documentRecord.fileType.startsWith('image/');
        case 'other':
          return !documentRecord.fileType.includes('pdf') && !documentRecord.fileType.startsWith('image/');
        default:
          return true;
      }
    })
    .sort(
      (first, second) =>
        (new Date(second.documentDate).getTime() || 0) - (new Date(first.documentDate).getTime() || 0),
    );
}

function formatCatMeta(cat) {
  const fragments = [];
  if (cat.breed) {
    fragments.push(cat.breed);
  }
  if (cat.birthdate) {
    fragments.push(`Дата рождения: ${formatDocumentDate(cat.birthdate)}`);
  }
  return fragments.join(' · ');
}

function formatDocumentDate(value) {
  if (!value) {
    return '—';
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '—';
  }
  return date.toLocaleDateString('ru-RU', { year: 'numeric', month: 'long', day: 'numeric' });
}

function formatTimestamp(value) {
  if (!value) {
    return '';
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return '';
  }
  return date.toLocaleString('ru-RU', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });
}

function formatFileSize(bytes) {
  if (typeof bytes !== 'number' || Number.isNaN(bytes) || bytes <= 0) {
    return '—';
  }
  const units = ['Б', 'КБ', 'МБ', 'ГБ'];
  const exponent = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** exponent;
  return `${value.toFixed(value >= 10 || exponent === 0 ? 0 : 1)} ${units[exponent]}`;
}

function generateId(seed) {
  const normalizedSeed = sanitizeText(seed).toLowerCase().replace(/\s+/g, '-') || 'item';
  if (hasCrypto) {
    const buffer = new Uint32Array(1);
    crypto.getRandomValues(buffer);
    return `${normalizedSeed}-${buffer[0].toString(16)}`;
  }
  return `${normalizedSeed}-${Math.random().toString(16).slice(2, 10)}`;
}

function sanitizeText(value) {
  if (typeof value !== 'string') {
    return '';
  }
  return value.replace(/\s+/g, ' ').trim();
}

function loadArchive() {
  try {
    if (typeof localStorage === 'undefined') {
      return cloneArchive(defaultArchive);
    }
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return cloneArchive(defaultArchive);
    }
    const parsed = JSON.parse(raw);
    return normalizeArchive(parsed);
  } catch (error) {
    console.warn('Не удалось прочитать архив. Используется резервный набор данных.', error);
    return cloneArchive(defaultArchive);
  }
}

function normalizeArchive(value) {
  if (!value || typeof value !== 'object') {
    return cloneArchive(defaultArchive);
  }

  const cats = Array.isArray(value.cats) ? value.cats : [];
  const selectedCatId = typeof value.selectedCatId === 'string' ? value.selectedCatId : cats[0]?.id;
  const lastUpdated = typeof value.lastUpdated === 'string' ? value.lastUpdated : new Date().toISOString();

  const normalizedCats = cats.map((cat) => ({
      id: String(cat.id || generateId('cat')),
      name: sanitizeText(cat.name) || 'Безымянный питомец',
      breed: sanitizeText(cat.breed),
      birthdate: cat.birthdate || '',
      accentColor: cat.accentColor || '#d2a45b',
      notes: sanitizeText(cat.notes),
      documents: Array.isArray(cat.documents)
        ? cat.documents.map((documentRecord) => ({
            id: String(documentRecord.id || generateId('document')),
            title: sanitizeText(documentRecord.title) || 'Без названия',
            description: sanitizeText(documentRecord.description),
            documentDate: documentRecord.documentDate || '',
            uploadedAt: documentRecord.uploadedAt || new Date().toISOString(),
            fileName: sanitizeText(documentRecord.fileName) || 'document.dat',
            fileType: sanitizeText(documentRecord.fileType) || 'application/octet-stream',
            size: Number(documentRecord.size) || 0,
            dataUrl: typeof documentRecord.dataUrl === 'string' ? documentRecord.dataUrl : '',
          }))
        : [],
    }));

  const resolvedSelectedCatId =
    typeof value.selectedCatId === 'string' && normalizedCats.some((cat) => cat.id === value.selectedCatId)
      ? value.selectedCatId
      : normalizedCats[0]?.id || null;

  return {
    cats: normalizedCats,
    selectedCatId: resolvedSelectedCatId,
    lastUpdated,
  };
}

function saveArchive() {
  try {
    if (typeof localStorage === 'undefined') {
      return;
    }
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.archive));
  } catch (error) {
    console.error('Не удалось сохранить архив.', error);
  }
}

function syncArchiveWithStorage() {
  const normalized = normalizeArchive(state.archive);
  state.archive = normalized;
  saveArchive();
}

function updateLastUpdated() {
  state.archive.lastUpdated = new Date().toISOString();
}

async function handleExportArchive() {
  if (!window.crypto?.subtle) {
    window.alert('Браузер не поддерживает необходимое шифрование.');
    return;
  }

  const password = sanitizeText(window.prompt('Укажите пароль для шифрования архива:'));
  if (!password) {
    window.alert('Экспорт отменен: пароль не задан.');
    return;
  }

  try {
    const payload = await encryptArchive(state.archive, password);
    const blob = new Blob([JSON.stringify(payload)], { type: 'application/json' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `feline-vitae-archive-${new Date().toISOString().slice(0, 10)}.json`;
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
  } catch (error) {
    console.error('Ошибка экспорта архива', error);
    window.alert('Не удалось экспортировать архив. Проверьте пароль и повторите попытку.');
  }
}

async function handleImportArchive(event) {
  const file = event.target.files?.[0];
  event.target.value = '';

  if (!file) {
    return;
  }

  if (!window.crypto?.subtle) {
    window.alert('Браузер не поддерживает необходимое шифрование.');
    return;
  }

  const password = sanitizeText(window.prompt('Введите пароль, использованный при экспорте архива:'));
  if (!password) {
    window.alert('Импорт отменен: пароль не указан.');
    return;
  }

  try {
    const fileContent = await file.text();
    const payload = JSON.parse(fileContent);
    const archive = await decryptArchive(payload, password);
    state.archive = normalizeArchive(archive);
    state.selectedDocumentId = null;
    saveArchive();
    refreshAdminCatList();
    render();
  } catch (error) {
    console.error('Ошибка импорта архива', error);
    window.alert('Не удалось импортировать архив. Убедитесь в правильности файла и пароля.');
  }
}

async function encryptArchive(archive, password) {
  const encoder = new TextEncoder();
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const keyMaterial = await crypto.subtle.importKey('raw', encoder.encode(password), 'PBKDF2', false, ['deriveKey']);
  const key = await crypto.subtle.deriveKey(
    {
      name: 'PBKDF2',
      salt,
      iterations: 150000,
      hash: 'SHA-256',
    },
    keyMaterial,
    {
      name: 'AES-GCM',
      length: 256,
    },
    false,
    ['encrypt'],
  );

  const plaintext = encoder.encode(JSON.stringify(archive));
  const ciphertext = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, plaintext);

  return {
    version: 1,
    salt: bufferToBase64(salt),
    iv: bufferToBase64(iv),
    payload: bufferToBase64(new Uint8Array(ciphertext)),
  };
}

async function decryptArchive(payload, password) {
  if (!payload || payload.version !== 1) {
    throw new Error('Неверный формат архива.');
  }

  const encoder = new TextEncoder();
  const decoder = new TextDecoder();
  const salt = base64ToBuffer(payload.salt);
  const iv = base64ToBuffer(payload.iv);
  const cipherBuffer = base64ToBuffer(payload.payload);

  const keyMaterial = await crypto.subtle.importKey('raw', encoder.encode(password), 'PBKDF2', false, ['deriveKey']);
  const key = await crypto.subtle.deriveKey(
    {
      name: 'PBKDF2',
      salt,
      iterations: 150000,
      hash: 'SHA-256',
    },
    keyMaterial,
    {
      name: 'AES-GCM',
      length: 256,
    },
    false,
    ['decrypt'],
  );

  const decrypted = await crypto.subtle.decrypt({ name: 'AES-GCM', iv }, key, cipherBuffer);
  return JSON.parse(decoder.decode(new Uint8Array(decrypted)));
}

function bufferToBase64(buffer) {
  let binary = '';
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  bytes.forEach((byte) => {
    binary += String.fromCharCode(byte);
  });
  return btoa(binary);
}

function base64ToBuffer(value) {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function cloneArchive(value) {
  if (typeof structuredClone === 'function') {
    return structuredClone(value);
  }
  return JSON.parse(JSON.stringify(value));
}
