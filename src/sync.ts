import {
  DefaultManualZoom,
  DefaultZoomMode,
  type AppState,
  type DocumentLibraryMetadata,
  type DocumentReaderState,
  type DocumentSourceMetadata,
  MaxStoredManualZoom,
  MinManualZoom,
  clamp,
  normalizeAppState,
  type StoredComment,
  type StoredSubject,
} from './storage';

export const DriveSyncFileName = 'slide-study-sync-v1.json';
export const DriveSyncSchemaVersion = 1;

type TimestampMap = Record<string, number>;

type DriveSyncDocumentRecord = DocumentLibraryMetadata & {
  source: DocumentSourceMetadata;
  lastPageIndex: number;
  lastPageUpdatedAt: number;
};

type DriveSyncBookmarkRecord = {
  pages: number[];
  updatedAt: number;
};

export type DriveSyncEnvelope = {
  app: 'slide-study-drive-sync';
  schemaVersion: typeof DriveSyncSchemaVersion;
  updatedAt: number;
  updatedBy: string;
  subjects: Record<string, StoredSubject>;
  subjectTombstones: TimestampMap;
  documents: Record<string, DriveSyncDocumentRecord>;
  documentTombstones: TimestampMap;
  comments: StoredComment[];
  commentTombstones: TimestampMap;
  bookmarks: Record<string, DriveSyncBookmarkRecord>;
};

type DriveSyncPayload = Omit<DriveSyncEnvelope, 'updatedAt' | 'updatedBy'>;

export function createDriveSyncEnvelope(state: AppState, now = Date.now()): DriveSyncEnvelope {
  return {
    ...createDriveSyncPayload(state, now),
    updatedAt: now,
    updatedBy: state.settings.driveSync.deviceId,
  };
}

export function getDriveSyncFingerprint(state: AppState): string {
  return JSON.stringify(createDriveSyncPayload(state, 0));
}

export function mergeDriveSyncEnvelope(
  localState: AppState,
  remoteRaw: unknown,
  now = Date.now(),
): AppState {
  const remote = normalizeDriveSyncEnvelope(remoteRaw);
  if (!remote) return normalizeAppState(localState, now);

  const subjectTombstones = mergeTimestampMaps(
    localState.settings.driveSync.pendingSubjectTombstones,
    remote.subjectTombstones,
  );
  const documentTombstones = mergeTimestampMaps(
    localState.settings.driveSync.pendingDocumentTombstones,
    remote.documentTombstones,
  );
  const commentTombstones = mergeTimestampMaps(
    localState.settings.driveSync.pendingCommentTombstones,
    remote.commentTombstones,
  );

  const nextState: AppState = {
    ...localState,
    subjects: mergeSubjects(localState.subjects, remote.subjects, subjectTombstones),
    documents: { ...localState.documents },
    documentSources: { ...localState.documentSources },
    readerStates: { ...localState.readerStates },
    studyData: {
      comments: [],
      bookmarks: {},
    },
    settings: {
      ...localState.settings,
      driveSync: {
        ...localState.settings.driveSync,
        pendingSubjectTombstones: subjectTombstones,
        pendingDocumentTombstones: documentTombstones,
        pendingCommentTombstones: commentTombstones,
        lastPageUpdatedAt: { ...localState.settings.driveSync.lastPageUpdatedAt },
        bookmarkUpdatedAt: { ...localState.settings.driveSync.bookmarkUpdatedAt },
      },
    },
  };

  mergeDocuments(nextState, localState, remote, documentTombstones, now);
  nextState.studyData.comments = mergeComments(localState, remote, documentTombstones, commentTombstones);
  nextState.studyData.bookmarks = mergeBookmarks(nextState, localState, remote, documentTombstones);

  return normalizeAppState(nextState, now);
}

export function parseDriveSyncEnvelope(raw: unknown): DriveSyncEnvelope | null {
  return normalizeDriveSyncEnvelope(raw);
}

function createDriveSyncPayload(state: AppState, now: number): DriveSyncPayload {
  const driveDocKeys = getDriveDocKeys(state);
  const driveDocKeySet = new Set(driveDocKeys);
  const documents = Object.fromEntries(
    driveDocKeys.flatMap((docKey) => {
      const record = createDriveDocumentRecord(state, docKey, now);
      return record ? [[docKey, record]] : [];
    }),
  );
  const bookmarks = Object.fromEntries(
    driveDocKeys.flatMap((docKey) => {
      const pages = state.studyData.bookmarks[docKey] ?? [];
      const updatedAt = state.settings.driveSync.bookmarkUpdatedAt[docKey]
        ?? state.documents[docKey]?.updatedAt
        ?? now;
      return pages.length > 0 || state.settings.driveSync.bookmarkUpdatedAt[docKey]
        ? [[docKey, { pages, updatedAt }]]
        : [];
    }),
  );

  return {
    app: 'slide-study-drive-sync',
    schemaVersion: DriveSyncSchemaVersion,
    subjects: state.subjects,
    subjectTombstones: state.settings.driveSync.pendingSubjectTombstones,
    documents,
    documentTombstones: state.settings.driveSync.pendingDocumentTombstones,
    comments: state.studyData.comments.filter((comment) => driveDocKeySet.has(comment.docKey)),
    commentTombstones: state.settings.driveSync.pendingCommentTombstones,
    bookmarks,
  };
}

function createDriveDocumentRecord(state: AppState, docKey: string, now: number): DriveSyncDocumentRecord | null {
  const library = state.documents[docKey];
  const source = state.documentSources[docKey];
  const reader = state.readerStates[docKey];
  if (!library || !source || !reader || source.sourceKind !== 'drive' || !source.driveFileId) return null;
  return {
    ...library,
    source,
    lastPageIndex: reader.lastPageIndex,
    lastPageUpdatedAt: state.settings.driveSync.lastPageUpdatedAt[docKey] ?? library.updatedAt ?? now,
  };
}

function mergeSubjects(
  localSubjects: Record<string, StoredSubject>,
  remoteSubjects: Record<string, StoredSubject>,
  tombstones: TimestampMap,
): Record<string, StoredSubject> {
  const subjects = { ...localSubjects };
  Object.values(remoteSubjects).forEach((remoteSubject) => {
    const localSubject = subjects[remoteSubject.id];
    if (!localSubject || remoteSubject.updatedAt > localSubject.updatedAt) {
      subjects[remoteSubject.id] = remoteSubject;
    }
  });
  Object.entries(tombstones).forEach(([subjectId, deletedAt]) => {
    const subject = subjects[subjectId];
    if (subject && deletedAt >= subject.updatedAt) {
      delete subjects[subjectId];
    }
  });
  return subjects;
}

function mergeDocuments(
  nextState: AppState,
  localState: AppState,
  remote: DriveSyncEnvelope,
  tombstones: TimestampMap,
  now: number,
): void {
  const localDriveRecords = Object.fromEntries(
    getDriveDocKeys(localState).flatMap((docKey) => {
      const record = createDriveDocumentRecord(localState, docKey, now);
      return record ? [[docKey, record]] : [];
    }),
  );
  const docKeys = new Set([...Object.keys(localDriveRecords), ...Object.keys(remote.documents), ...Object.keys(tombstones)]);

  docKeys.forEach((docKey) => {
    const localRecord = localDriveRecords[docKey];
    const remoteRecord = remote.documents[docKey];
    const newestUpdatedAt = Math.max(localRecord?.updatedAt ?? 0, remoteRecord?.updatedAt ?? 0);
    const deletedAt = tombstones[docKey] ?? 0;
    if (deletedAt >= newestUpdatedAt) {
      removeDocument(nextState, docKey);
      return;
    }

    const libraryRecord = chooseNewerLibrary(localRecord, remoteRecord);
    const sourceRecord = libraryRecord === remoteRecord ? remoteRecord?.source : localRecord?.source ?? remoteRecord?.source;
    if (!libraryRecord || !sourceRecord) return;

    const localLastPageUpdatedAt = localRecord?.lastPageUpdatedAt ?? 0;
    const remoteLastPageUpdatedAt = remoteRecord?.lastPageUpdatedAt ?? 0;
    const lastPageIndex = remoteLastPageUpdatedAt > localLastPageUpdatedAt
      ? remoteRecord?.lastPageIndex ?? 0
      : localRecord?.lastPageIndex ?? 0;
    const lastPageUpdatedAt = Math.max(localLastPageUpdatedAt, remoteLastPageUpdatedAt, libraryRecord.updatedAt);
    const existingReader = nextState.readerStates[docKey] ?? localState.readerStates[docKey];

    nextState.documents[docKey] = {
      key: libraryRecord.key,
      title: libraryRecord.title,
      subjectId: libraryRecord.subjectId,
      createdAt: libraryRecord.createdAt,
      updatedAt: libraryRecord.updatedAt,
    };
    nextState.documentSources[docKey] = sourceRecord;
    nextState.readerStates[docKey] = {
      lastPageIndex,
      zoomMode: existingReader?.zoomMode ?? DefaultZoomMode,
      manualZoom: clamp(existingReader?.manualZoom ?? DefaultManualZoom, MinManualZoom, MaxStoredManualZoom),
    };
    nextState.settings.driveSync.lastPageUpdatedAt[docKey] = lastPageUpdatedAt;
  });
}

function mergeComments(
  localState: AppState,
  remote: DriveSyncEnvelope,
  documentTombstones: TimestampMap,
  commentTombstones: TimestampMap,
): StoredComment[] {
  const localNonDriveComments = localState.studyData.comments
    .filter((comment) => localState.documentSources[comment.docKey]?.sourceKind !== 'drive');
  const commentsById: Record<string, StoredComment> = {};

  [...localState.studyData.comments, ...remote.comments].forEach((comment) => {
    const documentDeletedAt = documentTombstones[comment.docKey] ?? 0;
    if (documentDeletedAt > 0) return;
    const deletedAt = commentTombstones[comment.id] ?? 0;
    if (deletedAt >= comment.updatedAt) return;
    const existing = commentsById[comment.id];
    if (!existing || comment.updatedAt > existing.updatedAt) {
      commentsById[comment.id] = comment;
    }
  });

  return [
    ...localNonDriveComments,
    ...Object.values(commentsById)
      .filter((comment) => localState.documentSources[comment.docKey]?.sourceKind === 'drive' || remote.documents[comment.docKey])
      .sort((a, b) => a.createdAt - b.createdAt),
  ];
}

function mergeBookmarks(
  nextState: AppState,
  localState: AppState,
  remote: DriveSyncEnvelope,
  documentTombstones: TimestampMap,
): Record<string, number[]> {
  const bookmarks: Record<string, number[]> = {};
  Object.entries(localState.studyData.bookmarks).forEach(([docKey, pages]) => {
    if (localState.documentSources[docKey]?.sourceKind !== 'drive') {
      bookmarks[docKey] = pages;
    }
  });

  const driveDocKeys = new Set([...getDriveDocKeys(nextState), ...Object.keys(remote.documents)]);
  driveDocKeys.forEach((docKey) => {
    if (documentTombstones[docKey]) return;
    const localPages = localState.studyData.bookmarks[docKey] ?? [];
    const localUpdatedAt = localState.settings.driveSync.bookmarkUpdatedAt[docKey] ?? localState.documents[docKey]?.updatedAt ?? 0;
    const remoteBookmark = remote.bookmarks[docKey];
    const remoteUpdatedAt = remoteBookmark?.updatedAt ?? 0;
    const pages = remoteUpdatedAt > localUpdatedAt
      ? remoteBookmark?.pages ?? []
      : localPages;
    if (pages.length > 0) bookmarks[docKey] = pages;
    if (localUpdatedAt > 0 || remoteUpdatedAt > 0) {
      nextState.settings.driveSync.bookmarkUpdatedAt[docKey] = Math.max(localUpdatedAt, remoteUpdatedAt);
    }
  });

  return bookmarks;
}

function normalizeDriveSyncEnvelope(raw: unknown): DriveSyncEnvelope | null {
  if (!isRecord(raw) || raw.app !== 'slide-study-drive-sync' || raw.schemaVersion !== DriveSyncSchemaVersion) {
    return null;
  }
  return {
    app: 'slide-study-drive-sync',
    schemaVersion: DriveSyncSchemaVersion,
    updatedAt: timestampOr(raw.updatedAt, 0),
    updatedBy: textOr(raw.updatedBy, 'unknown-device'),
    subjects: normalizeSubjects(raw.subjects),
    subjectTombstones: normalizeTimestampMap(raw.subjectTombstones),
    documents: normalizeSyncDocuments(raw.documents),
    documentTombstones: normalizeTimestampMap(raw.documentTombstones),
    comments: normalizeSyncComments(raw.comments),
    commentTombstones: normalizeTimestampMap(raw.commentTombstones),
    bookmarks: normalizeSyncBookmarks(raw.bookmarks),
  };
}

function normalizeSubjects(rawSubjects: unknown): Record<string, StoredSubject> {
  if (!isRecord(rawSubjects)) return {};
  return Object.fromEntries(
    Object.entries(rawSubjects).flatMap(([fallbackId, rawSubject]) => {
      if (!isRecord(rawSubject)) return [];
      const id = textOr(rawSubject.id, fallbackId);
      const updatedAt = timestampOr(rawSubject.updatedAt, timestampOr(rawSubject.createdAt, 0));
      if (!id || updatedAt <= 0) return [];
      return [[id, {
        id,
        name: textOr(rawSubject.name, 'Untitled subject'),
        createdAt: timestampOr(rawSubject.createdAt, updatedAt),
        updatedAt,
      } satisfies StoredSubject]];
    }),
  );
}

function normalizeSyncDocuments(rawDocuments: unknown): Record<string, DriveSyncDocumentRecord> {
  if (!isRecord(rawDocuments)) return {};
  return Object.fromEntries(
    Object.entries(rawDocuments).flatMap(([fallbackKey, rawDocument]) => {
      if (!isRecord(rawDocument) || !isRecord(rawDocument.source)) return [];
      const key = textOr(rawDocument.key, fallbackKey);
      const source = normalizeDriveSource(rawDocument.source);
      if (!source) return [];
      const updatedAt = timestampOr(rawDocument.updatedAt, timestampOr(rawDocument.createdAt, 0));
      if (!key || updatedAt <= 0) return [];
      return [[key, {
        key,
        title: textOr(rawDocument.title, source.fileName.replace(/\.pdf$/i, '') || 'Drive PDF'),
        subjectId: optionalText(rawDocument.subjectId),
        createdAt: timestampOr(rawDocument.createdAt, updatedAt),
        updatedAt,
        source,
        lastPageIndex: integerInRange(rawDocument.lastPageIndex, 0, 0, source.pageCount - 1),
        lastPageUpdatedAt: timestampOr(rawDocument.lastPageUpdatedAt, updatedAt),
      } satisfies DriveSyncDocumentRecord]];
    }),
  );
}

function normalizeDriveSource(rawSource: Record<string, unknown>): DocumentSourceMetadata | null {
  const driveFileId = optionalText(rawSource.driveFileId);
  if (!driveFileId) return null;
  const driveName = optionalText(rawSource.driveName);
  const fileName = textOr(rawSource.fileName, driveName ?? 'Drive PDF.pdf');
  return {
    sourceKind: 'drive',
    pageCount: integerInRange(rawSource.pageCount, 1, 1, 100000),
    fileName,
    fileSize: integerInRange(rawSource.fileSize, 0, 0, Number.MAX_SAFE_INTEGER),
    fileLastModified: timestampOr(rawSource.fileLastModified, 1),
    driveFileId,
    driveName: driveName ?? fileName,
    driveModifiedTime: optionalText(rawSource.driveModifiedTime),
    driveSize: integerInRange(rawSource.driveSize ?? rawSource.fileSize, 0, 0, Number.MAX_SAFE_INTEGER),
  };
}

function normalizeSyncComments(rawComments: unknown): StoredComment[] {
  if (!Array.isArray(rawComments)) return [];
  const seenIds = new Set<string>();
  return rawComments.flatMap((rawComment) => {
    if (!isRecord(rawComment)) return [];
    const id = optionalText(rawComment.id);
    const docKey = optionalText(rawComment.docKey);
    const body = optionalText(rawComment.body);
    const updatedAt = timestampOr(rawComment.updatedAt, timestampOr(rawComment.createdAt, 0));
    if (!id || !docKey || !body || updatedAt <= 0 || seenIds.has(id)) return [];
    seenIds.add(id);
    return [{
      id,
      docKey,
      pageIndex: integerInRange(rawComment.pageIndex, 0, 0, 100000),
      body,
      createdAt: timestampOr(rawComment.createdAt, updatedAt),
      updatedAt,
    } satisfies StoredComment];
  });
}

function normalizeSyncBookmarks(rawBookmarks: unknown): Record<string, DriveSyncBookmarkRecord> {
  if (!isRecord(rawBookmarks)) return {};
  return Object.fromEntries(
    Object.entries(rawBookmarks).flatMap(([docKey, rawBookmark]) => {
      if (!isRecord(rawBookmark) || !Array.isArray(rawBookmark.pages)) return [];
      const updatedAt = timestampOr(rawBookmark.updatedAt, 0);
      if (updatedAt <= 0) return [];
      const pages = Array.from(new Set(
        rawBookmark.pages
          .filter((page): page is number => typeof page === 'number' && Number.isFinite(page))
          .map((page) => Math.floor(page))
          .filter((page) => page >= 0),
      )).sort((a, b) => a - b);
      return [[docKey, { pages, updatedAt } satisfies DriveSyncBookmarkRecord]];
    }),
  );
}

function chooseNewerLibrary(
  localRecord: DriveSyncDocumentRecord | undefined,
  remoteRecord: DriveSyncDocumentRecord | undefined,
): DriveSyncDocumentRecord | undefined {
  if (!localRecord) return remoteRecord;
  if (!remoteRecord) return localRecord;
  return remoteRecord.updatedAt > localRecord.updatedAt ? remoteRecord : localRecord;
}

function removeDocument(state: AppState, docKey: string): void {
  delete state.documents[docKey];
  delete state.documentSources[docKey];
  delete state.readerStates[docKey];
  delete state.studyData.bookmarks[docKey];
  state.studyData.comments = state.studyData.comments.filter((comment) => comment.docKey !== docKey);
  if (state.settings.selectedDocKey === docKey) {
    state.settings.selectedDocKey = null;
  }
}

function getDriveDocKeys(state: AppState): string[] {
  return Object.keys(state.documents).filter((docKey) => {
    const source = state.documentSources[docKey];
    return source?.sourceKind === 'drive' && Boolean(source.driveFileId);
  });
}

function mergeTimestampMaps(...maps: TimestampMap[]): TimestampMap {
  const merged: TimestampMap = {};
  maps.forEach((map) => {
    Object.entries(map).forEach(([key, timestamp]) => {
      merged[key] = Math.max(merged[key] ?? 0, timestamp);
    });
  });
  return merged;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function textOr(value: unknown, fallback: string): string {
  return typeof value === 'string' && value.trim() ? value.trim() : fallback;
}

function optionalText(value: unknown): string | null {
  return typeof value === 'string' && value.trim() ? value.trim() : null;
}

function timestampOr(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : fallback;
}

function normalizeTimestampMap(value: unknown): TimestampMap {
  if (!isRecord(value)) return {};
  return Object.fromEntries(
    Object.entries(value).flatMap(([key, rawTimestamp]) => {
      const timestamp = timestampOr(rawTimestamp, 0);
      return timestamp > 0 ? [[key, timestamp]] : [];
    }),
  );
}

function integerInRange(value: unknown, fallback: number, min: number, max: number): number {
  const parsed = typeof value === 'number' && Number.isFinite(value) ? value : fallback;
  return Math.floor(clamp(parsed, min, Math.max(min, max)));
}
