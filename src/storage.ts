export type ZoomMode = 'fitPage' | 'fitWidth' | 'manual';
export type Theme = 'light' | 'dark';
export type SortMode = 'recent' | 'title' | 'added';

export type CopyPacketOptions = {
  includePageImage: boolean;
  includeComments: boolean;
  includeTranscript: boolean;
  includeTrack: boolean;
  includeTimeline: boolean;
  includeFixedPrompt: boolean;
};

export type VisibleCopySettingKey = 'includePageImage' | 'includeComments';
export type DocumentSourceKind = 'local' | 'drive';

export type DriveAuthSettingsState = {
  hasGrantedFileAccess: boolean;
  hasGrantedAppDataAccess: boolean;
};

export type DriveSyncSettingsState = {
  enabled: boolean;
  deviceId: string;
  lastSyncedAt: number | null;
  lastRemoteModifiedTime: string | null;
  pendingSubjectTombstones: Record<string, number>;
  pendingDocumentTombstones: Record<string, number>;
  pendingCommentTombstones: Record<string, number>;
  lastPageUpdatedAt: Record<string, number>;
  bookmarkUpdatedAt: Record<string, number>;
};

export type StoredSubject = {
  id: string;
  name: string;
  createdAt: number;
  updatedAt: number;
};

export type DocumentSourceMetadata = {
  sourceKind: DocumentSourceKind;
  pageCount: number;
  fileName: string;
  fileSize: number;
  fileLastModified: number;
  driveFileId: string | null;
  driveName: string | null;
  driveModifiedTime: string | null;
  driveSize: number | null;
};

export type DocumentReaderState = {
  lastPageIndex: number;
  zoomMode: ZoomMode;
  manualZoom: number;
};

export type DocumentLibraryMetadata = {
  key: string;
  title: string;
  subjectId: string | null;
  createdAt: number;
  updatedAt: number;
};

export type StoredDocument = DocumentLibraryMetadata & DocumentSourceMetadata & DocumentReaderState;

export type StoredComment = {
  id: string;
  docKey: string;
  pageIndex: number;
  body: string;
  createdAt: number;
  updatedAt: number;
};

export type StudyDataState = {
  comments: StoredComment[];
  bookmarks: Record<string, number[]>;
};

export type AppSettingsState = {
  theme: Theme;
  selectedDocKey: string | null;
  selectedSubjectId: string | null;
  sortMode: SortMode;
  copySettings: CopyPacketOptions;
  driveAuth: DriveAuthSettingsState;
  driveSync: DriveSyncSettingsState;
};

export type AppState = {
  schemaVersion: 5;
  settings: AppSettingsState;
  subjects: Record<string, StoredSubject>;
  documents: Record<string, DocumentLibraryMetadata>;
  documentSources: Record<string, DocumentSourceMetadata>;
  readerStates: Record<string, DocumentReaderState>;
  studyData: StudyDataState;
};

export type AppBackupEnvelope = {
  app: 'slide-study';
  schemaVersion: 1;
  exportedAt: string;
  includes: {
    settings: true;
    pdfFiles: false;
    readerStates: true;
    studyData: true;
  };
  syncPolicy: {
    settings: 'backup-only';
    pdfFiles: 'external-source';
    merge: 'replace-only';
  };
  futureStudyDataSlots: typeof FutureStudyDataSlots[number][];
  data: AppState;
};

type AppStateRecord = Record<string, unknown>;
type SettingsStoreRecord = AppSettingsState & { id: string };
type DocumentSourceStoreRecord = DocumentSourceMetadata & { docKey: string };
type DocumentReaderStoreRecord = DocumentReaderState & { docKey: string };
type BookmarkStoreRecord = { docKey: string; pages: number[] };

export const AppStorageKey = 'slide-study-web-v2';
export const AppIndexedDbName = 'slide-study-web';
export const AppIndexedDbVersion = 1;
export const AppStateSchemaVersion = 5;
export const AppBackupSchemaVersion = 1;
export const UnfiledSubjectId = '__unfiled__';
export const DefaultZoomMode: ZoomMode = 'fitPage';
export const DefaultManualZoom = 1.15;
export const MinManualZoom = 0.25;
export const MaxStoredManualZoom = 4;
export const FutureStudyDataSlots = ['track', 'transcript', 'timeline'] as const;

export const DefaultCopyPacketOptions: CopyPacketOptions = {
  includePageImage: true,
  includeComments: true,
  includeTranscript: false,
  includeTrack: false,
  includeTimeline: false,
  includeFixedPrompt: false,
};

export const DefaultDriveAuthSettings: DriveAuthSettingsState = {
  hasGrantedFileAccess: false,
  hasGrantedAppDataAccess: false,
};

export const DefaultDriveSyncSettings: DriveSyncSettingsState = {
  enabled: false,
  deviceId: '',
  lastSyncedAt: null,
  lastRemoteModifiedTime: null,
  pendingSubjectTombstones: {},
  pendingDocumentTombstones: {},
  pendingCommentTombstones: {},
  lastPageUpdatedAt: {},
  bookmarkUpdatedAt: {},
};

const SettingsStoreKey = 'settings';
const StoreNames = {
  settings: 'settings',
  subjects: 'subjects',
  documents: 'documents',
  documentSources: 'documentSources',
  readerStates: 'readerStates',
  comments: 'comments',
  bookmarks: 'bookmarks',
} as const;
const AllStoreNames = Object.values(StoreNames);
let saveQueue: Promise<void> = Promise.resolve();

export function createInitialAppState(): AppState {
  return {
    schemaVersion: AppStateSchemaVersion,
    settings: {
      theme: 'light',
      selectedDocKey: null,
      selectedSubjectId: null,
      sortMode: 'recent',
      copySettings: DefaultCopyPacketOptions,
      driveAuth: DefaultDriveAuthSettings,
      driveSync: {
        ...DefaultDriveSyncSettings,
        deviceId: makeDeviceId(),
      },
    },
    subjects: {},
    documents: {},
    documentSources: {},
    readerStates: {},
    studyData: {
      comments: [],
      bookmarks: {},
    },
  };
}

export async function loadAppState(): Promise<AppState> {
  try {
    const indexedDbState = await loadIndexedDbAppState();
    if (indexedDbState) return indexedDbState;

    const localStorageState = loadLocalStorageAppState();
    await saveIndexedDbAppState(localStorageState).catch(() => undefined);
    return localStorageState;
  } catch {
    return loadLocalStorageAppState();
  }
}

export async function saveAppState(state: AppState): Promise<void> {
  const normalized = normalizeAppState(state);
  saveQueue = saveQueue
    .catch(() => undefined)
    .then(async () => {
      try {
        await saveIndexedDbAppState(normalized);
        saveLocalStorageAppState(normalized);
      } catch {
        saveLocalStorageAppState(normalized);
      }
    });
  return saveQueue;
}

function loadLocalStorageAppState(): AppState {
  try {
    const raw = localStorage.getItem(AppStorageKey);
    if (!raw) return createInitialAppState();
    return normalizeAppState(JSON.parse(raw) as unknown);
  } catch {
    return createInitialAppState();
  }
}

function saveLocalStorageAppState(state: AppState): void {
  localStorage.setItem(AppStorageKey, JSON.stringify(state));
}

export function normalizeAppState(raw: unknown, now = Date.now()): AppState {
  const parsed = isRecord(raw) ? raw : {};
  const subjects = normalizeSubjects(parsed.subjects, now);
  const sources = normalizeDocumentSources(parsed.documentSources, now);
  const documents = normalizeDocuments(parsed.documents, sources, subjects, now);
  const readerStates = normalizeReaderStates(parsed.readerStates, sources);
  const settings = normalizeAppSettings(parsed.settings, subjects, documents);
  const studyData = normalizeStudyData(parsed.studyData, sources, now);

  return {
    schemaVersion: AppStateSchemaVersion,
    settings,
    subjects,
    documents,
    documentSources: sources,
    readerStates,
    studyData,
  };
}

export function createAppBackupEnvelope(state: AppState, exportedAt = new Date()): AppBackupEnvelope {
  return {
    app: 'slide-study',
    schemaVersion: AppBackupSchemaVersion,
    exportedAt: exportedAt.toISOString(),
    includes: {
      settings: true,
      pdfFiles: false,
      readerStates: true,
      studyData: true,
    },
    syncPolicy: {
      settings: 'backup-only',
      pdfFiles: 'external-source',
      merge: 'replace-only',
    },
    futureStudyDataSlots: [...FutureStudyDataSlots],
    data: normalizeAppState(state, exportedAt.getTime()),
  };
}

export function serializeAppBackup(state: AppState): string {
  return JSON.stringify(createAppBackupEnvelope(state), null, 2);
}

export function restoreAppStateFromBackup(raw: unknown, now = Date.now()): AppState {
  if (!isRecord(raw)) {
    throw new Error('Backup file is not valid JSON data.');
  }

  if (raw.app !== 'slide-study') {
    throw new Error('This is not a Slide Study backup file.');
  }

  if (raw.schemaVersion !== AppBackupSchemaVersion) {
    throw new Error('This backup version is not supported.');
  }

  if (!isRecord(raw.data)) {
    throw new Error('Backup file is missing app data.');
  }

  return normalizeAppState(raw.data, now);
}

export function composeStoredDocument(
  library: DocumentLibraryMetadata,
  source: DocumentSourceMetadata,
  reader: DocumentReaderState,
): StoredDocument {
  return {
    ...library,
    ...source,
    ...reader,
  };
}

export function getStoredDocument(state: AppState, docKey: string | null): StoredDocument | null {
  if (!docKey) return null;
  const library = state.documents[docKey];
  const source = state.documentSources[docKey];
  const reader = state.readerStates[docKey];
  if (!library || !source || !reader) return null;
  return composeStoredDocument(library, source, reader);
}

export function getAllStoredDocuments(state: AppState): StoredDocument[] {
  return Object.keys(state.documents)
    .map((key) => getStoredDocument(state, key))
    .filter((doc): doc is StoredDocument => Boolean(doc));
}

export function titleFromFileName(fileName: string): string {
  return fileName.replace(/\.pdf$/i, '').trim() || 'Lecture PDF';
}

export function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

async function loadIndexedDbAppState(): Promise<AppState | null> {
  const db = await openAppDb();
  try {
    const transaction = db.transaction(AllStoreNames, 'readonly');
    const settingsPromise = getStoreValue<SettingsStoreRecord>(transaction, StoreNames.settings, SettingsStoreKey);
    const subjectsPromise = getAllStoreValues<StoredSubject>(transaction, StoreNames.subjects);
    const documentsPromise = getAllStoreValues<DocumentLibraryMetadata>(transaction, StoreNames.documents);
    const documentSourcesPromise = getAllStoreValues<DocumentSourceStoreRecord>(transaction, StoreNames.documentSources);
    const readerStatesPromise = getAllStoreValues<DocumentReaderStoreRecord>(transaction, StoreNames.readerStates);
    const commentsPromise = getAllStoreValues<StoredComment>(transaction, StoreNames.comments);
    const bookmarksPromise = getAllStoreValues<BookmarkStoreRecord>(transaction, StoreNames.bookmarks);
    const [
      settingsRecord,
      subjects,
      documents,
      documentSources,
      readerStates,
      comments,
      bookmarks,
    ] = await Promise.all([
      settingsPromise,
      subjectsPromise,
      documentsPromise,
      documentSourcesPromise,
      readerStatesPromise,
      commentsPromise,
      bookmarksPromise,
    ]);
    await waitForTransaction(transaction);

    const hasStoredData = Boolean(settingsRecord)
      || subjects.length > 0
      || documents.length > 0
      || documentSources.length > 0
      || readerStates.length > 0
      || comments.length > 0
      || bookmarks.length > 0;
    if (!hasStoredData) return null;

    return normalizeAppState({
      schemaVersion: AppStateSchemaVersion,
      settings: settingsRecord ? stripRecordId(settingsRecord) : undefined,
      subjects: Object.fromEntries(subjects.map((subject) => [subject.id, subject])),
      documents: Object.fromEntries(documents.map((document) => [document.key, document])),
      documentSources: Object.fromEntries(documentSources.map(({ docKey, ...source }) => [docKey, source])),
      readerStates: Object.fromEntries(readerStates.map(({ docKey, ...readerState }) => [docKey, readerState])),
      studyData: {
        comments,
        bookmarks: Object.fromEntries(bookmarks.map((bookmark) => [bookmark.docKey, bookmark.pages])),
      },
    });
  } finally {
    db.close();
  }
}

async function saveIndexedDbAppState(state: AppState): Promise<void> {
  const db = await openAppDb();
  try {
    const transaction = db.transaction(AllStoreNames, 'readwrite');
    AllStoreNames.forEach((storeName) => {
      transaction.objectStore(storeName).clear();
    });

    transaction.objectStore(StoreNames.settings).put({ id: SettingsStoreKey, ...state.settings });
    Object.values(state.subjects).forEach((subject) => {
      transaction.objectStore(StoreNames.subjects).put(subject);
    });
    Object.values(state.documents).forEach((document) => {
      transaction.objectStore(StoreNames.documents).put(document);
    });
    Object.entries(state.documentSources).forEach(([docKey, source]) => {
      transaction.objectStore(StoreNames.documentSources).put({ docKey, ...source });
    });
    Object.entries(state.readerStates).forEach(([docKey, readerState]) => {
      transaction.objectStore(StoreNames.readerStates).put({ docKey, ...readerState });
    });
    state.studyData.comments.forEach((comment) => {
      transaction.objectStore(StoreNames.comments).put(comment);
    });
    Object.entries(state.studyData.bookmarks).forEach(([docKey, pages]) => {
      transaction.objectStore(StoreNames.bookmarks).put({ docKey, pages });
    });

    await waitForTransaction(transaction);
  } finally {
    db.close();
  }
}

function openAppDb(): Promise<IDBDatabase> {
  if (typeof indexedDB === 'undefined') {
    return Promise.reject(new Error('IndexedDB is not available.'));
  }

  return new Promise((resolve, reject) => {
    const request = indexedDB.open(AppIndexedDbName, AppIndexedDbVersion);

    request.onupgradeneeded = () => {
      const db = request.result;
      ensureObjectStore(db, StoreNames.settings, 'id');
      ensureObjectStore(db, StoreNames.subjects, 'id');
      ensureObjectStore(db, StoreNames.documents, 'key');
      ensureObjectStore(db, StoreNames.documentSources, 'docKey');
      ensureObjectStore(db, StoreNames.readerStates, 'docKey');
      ensureObjectStore(db, StoreNames.comments, 'id');
      ensureObjectStore(db, StoreNames.bookmarks, 'docKey');
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error('Could not open IndexedDB.'));
  });
}

function ensureObjectStore(db: IDBDatabase, name: string, keyPath: string): void {
  if (!db.objectStoreNames.contains(name)) {
    db.createObjectStore(name, { keyPath });
  }
}

function getStoreValue<T>(transaction: IDBTransaction, storeName: string, key: IDBValidKey): Promise<T | undefined> {
  return new Promise((resolve, reject) => {
    const request = transaction.objectStore(storeName).get(key);
    request.onsuccess = () => resolve(request.result as T | undefined);
    request.onerror = () => reject(request.error ?? new Error(`Could not read ${storeName}.`));
  });
}

function getAllStoreValues<T>(transaction: IDBTransaction, storeName: string): Promise<T[]> {
  return new Promise((resolve, reject) => {
    const request = transaction.objectStore(storeName).getAll();
    request.onsuccess = () => resolve(request.result as T[]);
    request.onerror = () => reject(request.error ?? new Error(`Could not read ${storeName}.`));
  });
}

function waitForTransaction(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onabort = () => reject(transaction.error ?? new Error('IndexedDB transaction aborted.'));
    transaction.onerror = () => reject(transaction.error ?? new Error('IndexedDB transaction failed.'));
  });
}

function stripRecordId(record: SettingsStoreRecord): AppSettingsState {
  const { id: _id, ...settings } = record;
  return settings;
}

function isRecord(value: unknown): value is AppStateRecord {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function textOr(value: unknown, fallback: string): string {
  return typeof value === 'string' && value.trim() ? value.trim() : fallback;
}

function optionalText(value: unknown): string | null {
  return typeof value === 'string' && value.trim() ? value.trim() : null;
}

function finiteNumberOr(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : fallback;
}

function timestampOr(value: unknown, fallback: number): number {
  const timestamp = finiteNumberOr(value, fallback);
  return timestamp > 0 ? timestamp : fallback;
}

function nullableTimestamp(value: unknown): number | null {
  const timestamp = finiteNumberOr(value, 0);
  return timestamp > 0 ? timestamp : null;
}

function normalizeTimestampMap(value: unknown): Record<string, number> {
  if (!isRecord(value)) return {};
  return Object.fromEntries(
    Object.entries(value)
      .flatMap(([key, rawTimestamp]) => {
        const timestamp = nullableTimestamp(rawTimestamp);
        return timestamp ? [[key, timestamp]] : [];
      }),
  );
}

function makeDeviceId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return `device-${crypto.randomUUID()}`;
  }
  return `device-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function integerInRange(value: unknown, fallback: number, min: number, max: number): number {
  const boundedMax = Math.max(min, max);
  return Math.floor(clamp(finiteNumberOr(value, fallback), min, boundedMax));
}

function normalizeZoomMode(value: unknown): ZoomMode {
  return value === 'fitPage' || value === 'fitWidth' || value === 'manual'
    ? value
    : DefaultZoomMode;
}

function normalizeSortMode(value: unknown): SortMode {
  return value === 'title' || value === 'added' || value === 'recent'
    ? value
    : 'recent';
}

function normalizeCopySettings(settings?: AppStateRecord): CopyPacketOptions {
  return {
    includePageImage: typeof settings?.includePageImage === 'boolean' ? settings.includePageImage : DefaultCopyPacketOptions.includePageImage,
    includeComments: typeof settings?.includeComments === 'boolean' ? settings.includeComments : DefaultCopyPacketOptions.includeComments,
    includeTranscript: typeof settings?.includeTranscript === 'boolean' ? settings.includeTranscript : DefaultCopyPacketOptions.includeTranscript,
    includeTrack: typeof settings?.includeTrack === 'boolean' ? settings.includeTrack : DefaultCopyPacketOptions.includeTrack,
    includeTimeline: typeof settings?.includeTimeline === 'boolean' ? settings.includeTimeline : DefaultCopyPacketOptions.includeTimeline,
    includeFixedPrompt: typeof settings?.includeFixedPrompt === 'boolean' ? settings.includeFixedPrompt : DefaultCopyPacketOptions.includeFixedPrompt,
  };
}

function normalizeSubjects(rawSubjects: unknown, now: number): Record<string, StoredSubject> {
  if (!isRecord(rawSubjects)) return {};
  const subjects: Record<string, StoredSubject> = {};

  Object.entries(rawSubjects).forEach(([fallbackId, value]) => {
    if (!isRecord(value)) return;
    const id = textOr(value.id, fallbackId);
    subjects[id] = {
      id,
      name: textOr(value.name, 'Untitled subject'),
      createdAt: timestampOr(value.createdAt, timestampOr(value.updatedAt, now)),
      updatedAt: timestampOr(value.updatedAt, now),
    };
  });

  return subjects;
}

function normalizeDocumentSources(rawSources: unknown, now: number): Record<string, DocumentSourceMetadata> {
  if (!isRecord(rawSources)) return {};
  const sources: Record<string, DocumentSourceMetadata> = {};

  Object.entries(rawSources).forEach(([docKey, value]) => {
    if (!isRecord(value)) return;
    const sourceKind: DocumentSourceKind = value.sourceKind === 'drive' ? 'drive' : 'local';
    const driveName = optionalText(value.driveName);
    const fileName = textOr(value.fileName, driveName ?? 'Lecture PDF.pdf');
    const driveSize = sourceKind === 'drive'
      ? integerInRange(value.driveSize ?? value.fileSize, 0, 0, Number.MAX_SAFE_INTEGER)
      : null;
    sources[docKey] = {
      sourceKind,
      pageCount: integerInRange(value.pageCount, 1, 1, 100000),
      fileName,
      fileSize: integerInRange(value.fileSize, 0, 0, Number.MAX_SAFE_INTEGER),
      fileLastModified: timestampOr(value.fileLastModified, now),
      driveFileId: sourceKind === 'drive' ? optionalText(value.driveFileId) : null,
      driveName: sourceKind === 'drive' ? driveName ?? fileName : null,
      driveModifiedTime: sourceKind === 'drive' ? optionalText(value.driveModifiedTime) : null,
      driveSize,
    };
  });

  return sources;
}

function normalizeDocuments(
  rawDocuments: unknown,
  sources: Record<string, DocumentSourceMetadata>,
  subjects: Record<string, StoredSubject>,
  now: number,
): Record<string, DocumentLibraryMetadata> {
  if (!isRecord(rawDocuments)) return {};
  const documents: Record<string, DocumentLibraryMetadata> = {};

  Object.entries(rawDocuments).forEach(([fallbackKey, value]) => {
    if (!isRecord(value) || !sources[fallbackKey]) return;
    const key = textOr(value.key, fallbackKey);
    if (!sources[key]) return;
    documents[key] = {
      key,
      title: textOr(value.title, titleFromFileName(sources[key].fileName)),
      subjectId: typeof value.subjectId === 'string' && subjects[value.subjectId] ? value.subjectId : null,
      createdAt: timestampOr(value.createdAt, timestampOr(value.updatedAt, now)),
      updatedAt: timestampOr(value.updatedAt, now),
    };
  });

  return documents;
}

function normalizeReaderStates(
  rawReaderStates: unknown,
  sources: Record<string, DocumentSourceMetadata>,
): Record<string, DocumentReaderState> {
  if (!isRecord(rawReaderStates)) return {};
  const readerStates: Record<string, DocumentReaderState> = {};

  Object.entries(sources).forEach(([docKey, source]) => {
    const value = rawReaderStates[docKey];
    const record = isRecord(value) ? value : {};
    readerStates[docKey] = {
      lastPageIndex: integerInRange(record.lastPageIndex, 0, 0, source.pageCount - 1),
      zoomMode: normalizeZoomMode(record.zoomMode),
      manualZoom: clamp(finiteNumberOr(record.manualZoom, DefaultManualZoom), MinManualZoom, MaxStoredManualZoom),
    };
  });

  return readerStates;
}

function normalizeAppSettings(
  rawSettings: unknown,
  subjects: Record<string, StoredSubject>,
  documents: Record<string, DocumentLibraryMetadata>,
): AppSettingsState {
  const settings = isRecord(rawSettings) ? rawSettings : {};
  const selectedDocKey = typeof settings.selectedDocKey === 'string' && documents[settings.selectedDocKey]
    ? settings.selectedDocKey
    : null;
  const selectedSubjectId = settings.selectedSubjectId === UnfiledSubjectId
    ? UnfiledSubjectId
    : typeof settings.selectedSubjectId === 'string' && subjects[settings.selectedSubjectId]
      ? settings.selectedSubjectId
      : null;

  return {
    theme: settings.theme === 'dark' ? 'dark' : 'light',
    selectedDocKey,
    selectedSubjectId,
    sortMode: normalizeSortMode(settings.sortMode),
    copySettings: normalizeCopySettings(isRecord(settings.copySettings) ? settings.copySettings : undefined),
    driveAuth: normalizeDriveAuthSettings(isRecord(settings.driveAuth) ? settings.driveAuth : undefined),
    driveSync: normalizeDriveSyncSettings(isRecord(settings.driveSync) ? settings.driveSync : undefined),
  };
}

function normalizeDriveAuthSettings(rawDriveAuth: Record<string, unknown> | undefined): DriveAuthSettingsState {
  return {
    hasGrantedFileAccess: rawDriveAuth?.hasGrantedFileAccess === true,
    hasGrantedAppDataAccess: rawDriveAuth?.hasGrantedAppDataAccess === true,
  };
}

function normalizeDriveSyncSettings(rawDriveSync: Record<string, unknown> | undefined): DriveSyncSettingsState {
  return {
    enabled: rawDriveSync?.enabled === true,
    deviceId: textOr(rawDriveSync?.deviceId, makeDeviceId()),
    lastSyncedAt: nullableTimestamp(rawDriveSync?.lastSyncedAt),
    lastRemoteModifiedTime: optionalText(rawDriveSync?.lastRemoteModifiedTime),
    pendingSubjectTombstones: normalizeTimestampMap(rawDriveSync?.pendingSubjectTombstones),
    pendingDocumentTombstones: normalizeTimestampMap(rawDriveSync?.pendingDocumentTombstones),
    pendingCommentTombstones: normalizeTimestampMap(rawDriveSync?.pendingCommentTombstones),
    lastPageUpdatedAt: normalizeTimestampMap(rawDriveSync?.lastPageUpdatedAt),
    bookmarkUpdatedAt: normalizeTimestampMap(rawDriveSync?.bookmarkUpdatedAt),
  };
}

function normalizeStudyData(
  rawStudyData: unknown,
  sources: Record<string, DocumentSourceMetadata>,
  now: number,
): StudyDataState {
  const studyData = isRecord(rawStudyData) ? rawStudyData : {};
  return {
    comments: normalizeComments(studyData.comments, sources, now),
    bookmarks: normalizeBookmarks(studyData.bookmarks, sources),
  };
}

function normalizeBookmarks(
  rawBookmarks: unknown,
  sources: Record<string, DocumentSourceMetadata>,
): Record<string, number[]> {
  if (!isRecord(rawBookmarks)) return {};
  const bookmarks: Record<string, number[]> = {};

  Object.entries(sources).forEach(([docKey, source]) => {
    const rawPages = rawBookmarks[docKey];
    if (!Array.isArray(rawPages)) return;
    const pages = Array.from(new Set(
      rawPages
        .filter((page): page is number => typeof page === 'number' && Number.isFinite(page))
        .map((page) => Math.floor(page))
        .filter((page) => page >= 0 && page < source.pageCount),
    )).sort((a, b) => a - b);
    if (pages.length > 0) bookmarks[docKey] = pages;
  });

  return bookmarks;
}

function normalizeComments(
  rawComments: unknown,
  sources: Record<string, DocumentSourceMetadata>,
  now: number,
): StoredComment[] {
  if (!Array.isArray(rawComments)) return [];
  const seenIds = new Set<string>();

  return rawComments.flatMap((value, index) => {
    if (!isRecord(value) || typeof value.docKey !== 'string') return [];
    const source = sources[value.docKey];
    const body = typeof value.body === 'string' ? value.body.trim() : '';
    if (!source || !body) return [];

    const fallbackId = `comment-${value.docKey}-${index}`;
    const baseId = textOr(value.id, fallbackId);
    const id = seenIds.has(baseId) ? `${baseId}-${index}` : baseId;
    seenIds.add(id);

    return [{
      id,
      docKey: value.docKey,
      pageIndex: integerInRange(value.pageIndex, 0, 0, source.pageCount - 1),
      body,
      createdAt: timestampOr(value.createdAt, timestampOr(value.updatedAt, now)),
      updatedAt: timestampOr(value.updatedAt, now),
    }];
  });
}
