import {
  DefaultManualZoom,
  DefaultZoomMode,
  type AppState,
  type DocumentLibraryMetadata,
  type DocumentReaderState,
  type DocumentSourceMetadata,
  type DriveSyncOperation,
  MaxStoredManualZoom,
  MinManualZoom,
  clamp,
  normalizeAppState,
  type StoredComment,
  type StoredSubject,
} from './storage';

export const DriveSyncFileName = 'slide-study-sync-v1.json';
export const DriveSyncSchemaVersion = 2;
const LegacyDriveSyncSchemaVersion = 1;
const MaxStoredDriveSyncOperations = 2000;

type TimestampMap = Record<string, number>;

type DriveSyncDocumentRecord = DocumentLibraryMetadata & {
  source: DocumentSourceMetadata;
  lastPageIndex: number;
  lastPageUpdatedAt: number;
};

type DriveSyncBookmarkRecord = {
  pages: number[];
  updatedAt: number;
  pageUpdatedAt: TimestampMap;
};

export type DriveSyncEnvelope = {
  app: 'slide-study-drive-sync';
  schemaVersion: typeof DriveSyncSchemaVersion;
  updatedAt: number;
  updatedBy: string;
  operations: DriveSyncOperation[];
  subjects: Record<string, StoredSubject>;
  subjectTombstones: TimestampMap;
  documents: Record<string, DriveSyncDocumentRecord>;
  documentTombstones: TimestampMap;
  comments: StoredComment[];
  commentTombstones: TimestampMap;
  bookmarks: Record<string, DriveSyncBookmarkRecord>;
};

type DriveSyncPayload = Omit<DriveSyncEnvelope, 'updatedAt' | 'updatedBy'>;

export function createDriveSyncEnvelope(
  state: AppState,
  now = Date.now(),
  operations = state.settings.driveSync.pendingOperations,
): DriveSyncEnvelope {
  return {
    ...createDriveSyncPayload(state, now, operations),
    updatedAt: now,
    updatedBy: state.settings.driveSync.deviceId,
  };
}

export function getDriveSyncFingerprint(state: AppState): string {
  return JSON.stringify(createDriveSyncPayload(state, 0, state.settings.driveSync.pendingOperations));
}

export function mergeDriveSyncEnvelope(
  localState: AppState,
  remoteRaw: unknown,
  now = Date.now(),
  operations = mergeDriveSyncOperations(localState, remoteRaw),
): AppState {
  const remote = normalizeDriveSyncEnvelope(remoteRaw);
  if (!remote) return applyDriveSyncOperations(normalizeAppState(localState, now), operations, now);

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

  return applyDriveSyncOperations(normalizeAppState(nextState, now), operations, now);
}

export function parseDriveSyncEnvelope(raw: unknown): DriveSyncEnvelope | null {
  return normalizeDriveSyncEnvelope(raw);
}

export function mergeDriveSyncOperations(localState: AppState, remoteRaw: unknown): DriveSyncOperation[] {
  const remote = normalizeDriveSyncEnvelope(remoteRaw);
  return compactDriveSyncOperations([
    ...(remote?.operations ?? []),
    ...localState.settings.driveSync.pendingOperations,
  ]);
}

function createDriveSyncPayload(
  state: AppState,
  now: number,
  operations: DriveSyncOperation[],
): DriveSyncPayload {
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
      const pageUpdatedAt = createBookmarkPageTimestampMap(state, docKey, pages, now);
      const updatedAt = Math.max(
        state.settings.driveSync.bookmarkUpdatedAt[docKey] ?? 0,
        maxTimestamp(pageUpdatedAt),
      );
      return pages.length > 0 || updatedAt > 0
        ? [[docKey, { pages, updatedAt, pageUpdatedAt }]]
        : [];
    }),
  );

  return {
    app: 'slide-study-drive-sync',
    schemaVersion: DriveSyncSchemaVersion,
    operations: compactDriveSyncOperations(operations),
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

function createBookmarkPageTimestampMap(
  state: AppState,
  docKey: string,
  pages: number[],
  now: number,
): TimestampMap {
  const existing = state.settings.driveSync.bookmarkPageUpdatedAt[docKey] ?? {};
  const fallback = state.settings.driveSync.bookmarkUpdatedAt[docKey] ?? state.documents[docKey]?.updatedAt ?? now;
  const pageUpdatedAt = { ...existing };
  pages.forEach((page) => {
    const key = String(page);
    pageUpdatedAt[key] = pageUpdatedAt[key] ?? fallback;
  });
  return pageUpdatedAt;
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
    const lastPageUpdatedAt = Math.max(localLastPageUpdatedAt, remoteLastPageUpdatedAt);
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

  const driveDocKeys = new Set([
    ...getDriveDocKeys(nextState),
    ...Object.keys(remote.documents),
    ...Object.keys(remote.bookmarks),
  ]);
  driveDocKeys.forEach((docKey) => {
    if (documentTombstones[docKey]) return;
    if (!nextState.documents[docKey] && !remote.documents[docKey]) return;
    const localPages = localState.studyData.bookmarks[docKey] ?? [];
    const remoteBookmark = remote.bookmarks[docKey];
    const localUpdatedAt = localState.settings.driveSync.bookmarkUpdatedAt[docKey] ?? localState.documents[docKey]?.updatedAt ?? 0;
    const localPageUpdatedAt = createBookmarkPageTimestampMap(localState, docKey, localPages, localUpdatedAt);
    const remotePages = remoteBookmark?.pages ?? [];
    const remoteUpdatedAt = remoteBookmark?.updatedAt ?? 0;
    const remotePageUpdatedAt = remoteBookmark?.pageUpdatedAt ?? {};
    const pageCount = nextState.documentSources[docKey]?.pageCount ?? remote.documents[docKey]?.source.pageCount ?? Number.MAX_SAFE_INTEGER;
    const localPageSet = new Set(localPages);
    const remotePageSet = new Set(remotePages);
    const pageIndexes = new Set([
      ...localPages,
      ...remotePages,
      ...Object.keys(localPageUpdatedAt).map(Number),
      ...Object.keys(remotePageUpdatedAt).map(Number),
    ]);
    const pages: number[] = [];
    const nextPageUpdatedAt: TimestampMap = {};

    pageIndexes.forEach((page) => {
      if (!Number.isInteger(page) || page < 0 || page >= pageCount) return;
      const key = String(page);
      const localPageTimestamp = localPageUpdatedAt[key] ?? (localPageSet.has(page) ? localUpdatedAt : 0);
      const remotePageTimestamp = remotePageUpdatedAt[key] ?? (remotePageSet.has(page) ? remoteUpdatedAt : 0);
      const useRemote = remotePageTimestamp > localPageTimestamp;
      const bookmarked = useRemote ? remotePageSet.has(page) : localPageSet.has(page);
      const updatedAt = Math.max(localPageTimestamp, remotePageTimestamp);
      if (bookmarked) pages.push(page);
      if (updatedAt > 0) nextPageUpdatedAt[key] = updatedAt;
    });

    pages.sort((a, b) => a - b);
    if (pages.length > 0) bookmarks[docKey] = pages;
    const updatedAt = maxTimestamp(nextPageUpdatedAt);
    if (updatedAt > 0) {
      nextState.settings.driveSync.bookmarkUpdatedAt[docKey] = updatedAt;
      nextState.settings.driveSync.bookmarkPageUpdatedAt[docKey] = nextPageUpdatedAt;
    } else {
      delete nextState.settings.driveSync.bookmarkUpdatedAt[docKey];
      delete nextState.settings.driveSync.bookmarkPageUpdatedAt[docKey];
    }
  });

  return bookmarks;
}

function applyDriveSyncOperations(state: AppState, operations: DriveSyncOperation[], now: number): AppState {
  const nextState: AppState = {
    ...state,
    subjects: { ...state.subjects },
    documents: { ...state.documents },
    documentSources: { ...state.documentSources },
    readerStates: { ...state.readerStates },
    studyData: {
      comments: [...state.studyData.comments],
      bookmarks: Object.fromEntries(
        Object.entries(state.studyData.bookmarks).map(([docKey, pages]) => [docKey, [...pages]]),
      ),
    },
    settings: {
      ...state.settings,
      driveSync: {
        ...state.settings.driveSync,
        pendingSubjectTombstones: { ...state.settings.driveSync.pendingSubjectTombstones },
        pendingDocumentTombstones: { ...state.settings.driveSync.pendingDocumentTombstones },
        pendingCommentTombstones: { ...state.settings.driveSync.pendingCommentTombstones },
        lastPageUpdatedAt: { ...state.settings.driveSync.lastPageUpdatedAt },
        bookmarkUpdatedAt: { ...state.settings.driveSync.bookmarkUpdatedAt },
        bookmarkPageUpdatedAt: Object.fromEntries(
          Object.entries(state.settings.driveSync.bookmarkPageUpdatedAt)
            .map(([docKey, map]) => [docKey, { ...map }]),
        ),
        pendingOperations: [...state.settings.driveSync.pendingOperations],
      },
    },
  };

  const sortedOperations = compactDriveSyncOperations(operations)
    .sort(compareDriveSyncOperations);
  sortedOperations.forEach((operation) => {
    applyDriveSyncOperation(nextState, operation, now);
  });

  return normalizeAppState(nextState, now);
}

function applyDriveSyncOperation(state: AppState, operation: DriveSyncOperation, now: number): void {
  switch (operation.type) {
    case 'subjectUpsert':
      applySubjectUpsertOperation(state, operation.subject);
      return;
    case 'subjectDelete':
      state.settings.driveSync.pendingSubjectTombstones[operation.subjectId] = Math.max(
        state.settings.driveSync.pendingSubjectTombstones[operation.subjectId] ?? 0,
        operation.createdAt,
      );
      delete state.subjects[operation.subjectId];
      Object.entries(state.documents).forEach(([docKey, document]) => {
        if (document.subjectId === operation.subjectId) {
          state.documents[docKey] = { ...document, subjectId: null, updatedAt: Math.max(document.updatedAt, operation.createdAt) };
        }
      });
      if (state.settings.selectedSubjectId === operation.subjectId) {
        state.settings.selectedSubjectId = null;
      }
      return;
    case 'documentUpsert':
      applyDocumentUpsertOperation(state, operation.document, operation.source, operation.reader, operation.createdAt);
      return;
    case 'documentDelete':
      state.settings.driveSync.pendingDocumentTombstones[operation.docKey] = Math.max(
        state.settings.driveSync.pendingDocumentTombstones[operation.docKey] ?? 0,
        operation.createdAt,
      );
      removeDocument(state, operation.docKey);
      return;
    case 'lastPageSet':
      applyLastPageSetOperation(state, operation.docKey, operation.pageIndex, operation.createdAt);
      return;
    case 'bookmarkSet':
      applyBookmarkSetOperation(state, operation.docKey, operation.pageIndex, operation.bookmarked, operation.createdAt);
      return;
    case 'commentUpsert':
      applyCommentUpsertOperation(state, operation.comment, operation.createdAt);
      return;
    case 'commentDelete':
      state.settings.driveSync.pendingCommentTombstones[operation.commentId] = Math.max(
        state.settings.driveSync.pendingCommentTombstones[operation.commentId] ?? 0,
        operation.createdAt,
      );
      state.studyData.comments = state.studyData.comments.filter((comment) => {
        if (comment.id !== operation.commentId) return true;
        return comment.updatedAt > operation.createdAt;
      });
      return;
    default:
      return;
  }
}

function applySubjectUpsertOperation(state: AppState, subject: StoredSubject): void {
  const deletedAt = state.settings.driveSync.pendingSubjectTombstones[subject.id] ?? 0;
  if (deletedAt >= subject.updatedAt) return;
  const existing = state.subjects[subject.id];
  if (!existing || subject.updatedAt >= existing.updatedAt) {
    state.subjects[subject.id] = subject;
  }
}

function applyDocumentUpsertOperation(
  state: AppState,
  document: DocumentLibraryMetadata,
  source: DocumentSourceMetadata,
  reader: DocumentReaderState,
  operationAt: number,
): void {
  if (source.sourceKind !== 'drive' || !source.driveFileId) return;
  const deletedAt = state.settings.driveSync.pendingDocumentTombstones[document.key] ?? 0;
  if (deletedAt >= Math.max(document.updatedAt, operationAt)) return;
  const existing = state.documents[document.key];
  if (!existing || document.updatedAt >= existing.updatedAt) {
    state.documents[document.key] = document;
    state.documentSources[document.key] = source;
  }
  if (!state.readerStates[document.key]) {
    state.readerStates[document.key] = {
      lastPageIndex: clamp(reader.lastPageIndex, 0, source.pageCount - 1),
      zoomMode: DefaultZoomMode,
      manualZoom: DefaultManualZoom,
    };
    state.settings.driveSync.lastPageUpdatedAt[document.key] = operationAt;
  }
}

function applyLastPageSetOperation(state: AppState, docKey: string, pageIndex: number, operationAt: number): void {
  const source = state.documentSources[docKey];
  const reader = state.readerStates[docKey];
  if (!source || !reader || source.sourceKind !== 'drive') return;
  const previousUpdatedAt = state.settings.driveSync.lastPageUpdatedAt[docKey] ?? 0;
  if (operationAt < previousUpdatedAt) return;
  state.readerStates[docKey] = {
    ...reader,
    lastPageIndex: clamp(pageIndex, 0, source.pageCount - 1),
  };
  state.settings.driveSync.lastPageUpdatedAt[docKey] = operationAt;
}

function applyBookmarkSetOperation(
  state: AppState,
  docKey: string,
  pageIndex: number,
  bookmarked: boolean,
  operationAt: number,
): void {
  const source = state.documentSources[docKey];
  if (!source || source.sourceKind !== 'drive' || pageIndex < 0 || pageIndex >= source.pageCount) return;
  const bookmarkPageUpdatedAt = {
    ...(state.settings.driveSync.bookmarkPageUpdatedAt[docKey] ?? {}),
  };
  const pageKey = String(pageIndex);
  const previousUpdatedAt = bookmarkPageUpdatedAt[pageKey] ?? 0;
  if (operationAt < previousUpdatedAt) return;

  const pages = new Set(state.studyData.bookmarks[docKey] ?? []);
  if (bookmarked) {
    pages.add(pageIndex);
  } else {
    pages.delete(pageIndex);
  }
  const nextPages = Array.from(pages).sort((a, b) => a - b);
  if (nextPages.length > 0) {
    state.studyData.bookmarks[docKey] = nextPages;
  } else {
    delete state.studyData.bookmarks[docKey];
  }
  bookmarkPageUpdatedAt[pageKey] = operationAt;
  state.settings.driveSync.bookmarkPageUpdatedAt[docKey] = bookmarkPageUpdatedAt;
  state.settings.driveSync.bookmarkUpdatedAt[docKey] = Math.max(
    state.settings.driveSync.bookmarkUpdatedAt[docKey] ?? 0,
    operationAt,
  );
}

function applyCommentUpsertOperation(state: AppState, comment: StoredComment, operationAt: number): void {
  const source = state.documentSources[comment.docKey];
  if (!source || source.sourceKind !== 'drive') return;
  const deletedAt = state.settings.driveSync.pendingCommentTombstones[comment.id] ?? 0;
  if (deletedAt >= Math.max(comment.updatedAt, operationAt)) return;
  const nextComment: StoredComment = {
    ...comment,
    pageIndex: clamp(comment.pageIndex, 0, source.pageCount - 1),
  };
  const existingIndex = state.studyData.comments.findIndex((existing) => existing.id === comment.id);
  if (existingIndex >= 0) {
    const existing = state.studyData.comments[existingIndex];
    if (nextComment.updatedAt >= existing.updatedAt) {
      state.studyData.comments[existingIndex] = nextComment;
    }
    return;
  }
  state.studyData.comments.push(nextComment);
}

function normalizeDriveSyncEnvelope(raw: unknown): DriveSyncEnvelope | null {
  if (
    !isRecord(raw)
    || raw.app !== 'slide-study-drive-sync'
    || (raw.schemaVersion !== DriveSyncSchemaVersion && raw.schemaVersion !== LegacyDriveSyncSchemaVersion)
  ) {
    return null;
  }
  return {
    app: 'slide-study-drive-sync',
    schemaVersion: DriveSyncSchemaVersion,
    updatedAt: timestampOr(raw.updatedAt, 0),
    updatedBy: textOr(raw.updatedBy, 'unknown-device'),
    operations: normalizeSyncOperations(raw.operations),
    subjects: normalizeSubjects(raw.subjects),
    subjectTombstones: normalizeTimestampMap(raw.subjectTombstones),
    documents: normalizeSyncDocuments(raw.documents),
    documentTombstones: normalizeTimestampMap(raw.documentTombstones),
    comments: normalizeSyncComments(raw.comments),
    commentTombstones: normalizeTimestampMap(raw.commentTombstones),
    bookmarks: normalizeSyncBookmarks(raw.bookmarks),
  };
}

function normalizeSyncOperations(rawOperations: unknown): DriveSyncOperation[] {
  if (!Array.isArray(rawOperations)) return [];
  const seenIds = new Set<string>();

  return rawOperations.flatMap((rawOperation): DriveSyncOperation[] => {
    if (!isRecord(rawOperation)) return [];
    const type = optionalText(rawOperation.type);
    const deviceId = textOr(rawOperation.deviceId, 'unknown-device');
    const seq = integerInRange(rawOperation.seq, 0, 0, Number.MAX_SAFE_INTEGER);
    const createdAt = timestampOr(rawOperation.createdAt, 0);
    const id = textOr(rawOperation.id, `${deviceId}:${seq}:${type ?? 'operation'}`);
    if (createdAt <= 0 || seenIds.has(id)) return [];
    seenIds.add(id);
    const base = { id, deviceId, seq, createdAt };

    if (type === 'subjectUpsert' && isRecord(rawOperation.subject)) {
      const subject = Object.values(normalizeSubjects({ subject: rawOperation.subject }))[0];
      return subject ? [{ ...base, type, subject }] : [];
    }

    if (type === 'subjectDelete') {
      const subjectId = optionalText(rawOperation.subjectId);
      return subjectId ? [{ ...base, type, subjectId }] : [];
    }

    if (type === 'documentUpsert' && isRecord(rawOperation.document) && isRecord(rawOperation.source)) {
      const docKey = textOr(rawOperation.docKey, textOr(rawOperation.document.key, ''));
      const source = normalizeDriveSource(rawOperation.source);
      if (!docKey || !source) return [];
      const updatedAt = timestampOr(rawOperation.document.updatedAt, timestampOr(rawOperation.document.createdAt, createdAt));
      const rawReader = isRecord(rawOperation.reader) ? rawOperation.reader : {};
      return [{
        ...base,
        type,
        document: {
          key: docKey,
          title: textOr(rawOperation.document.title, source.fileName.replace(/\.pdf$/i, '') || 'Drive PDF'),
          subjectId: optionalText(rawOperation.document.subjectId),
          createdAt: timestampOr(rawOperation.document.createdAt, updatedAt),
          updatedAt,
        },
        source,
        reader: {
          lastPageIndex: integerInRange(rawReader.lastPageIndex, 0, 0, source.pageCount - 1),
          zoomMode: DefaultZoomMode,
          manualZoom: DefaultManualZoom,
        },
      }];
    }

    if (type === 'documentDelete') {
      const docKey = optionalText(rawOperation.docKey);
      return docKey ? [{ ...base, type, docKey }] : [];
    }

    if (type === 'lastPageSet') {
      const docKey = optionalText(rawOperation.docKey);
      return docKey
        ? [{ ...base, type, docKey, pageIndex: integerInRange(rawOperation.pageIndex, 0, 0, Number.MAX_SAFE_INTEGER) }]
        : [];
    }

    if (type === 'bookmarkSet') {
      const docKey = optionalText(rawOperation.docKey);
      return docKey
        ? [{
          ...base,
          type,
          docKey,
          pageIndex: integerInRange(rawOperation.pageIndex, 0, 0, Number.MAX_SAFE_INTEGER),
          bookmarked: rawOperation.bookmarked === true,
        }]
        : [];
    }

    if (type === 'commentUpsert' && isRecord(rawOperation.comment)) {
      const commentId = optionalText(rawOperation.comment.id);
      const docKey = optionalText(rawOperation.comment.docKey);
      const body = optionalText(rawOperation.comment.body);
      if (!commentId || !docKey || !body) return [];
      const updatedAt = timestampOr(rawOperation.comment.updatedAt, timestampOr(rawOperation.comment.createdAt, createdAt));
      return [{
        ...base,
        type,
        comment: {
          id: commentId,
          docKey,
          pageIndex: integerInRange(rawOperation.comment.pageIndex, 0, 0, Number.MAX_SAFE_INTEGER),
          body,
          createdAt: timestampOr(rawOperation.comment.createdAt, updatedAt),
          updatedAt,
        },
      }];
    }

    if (type === 'commentDelete') {
      const commentId = optionalText(rawOperation.commentId);
      const docKey = optionalText(rawOperation.docKey);
      return commentId && docKey ? [{ ...base, type, commentId, docKey }] : [];
    }

    return [];
  });
}

function compactDriveSyncOperations(operations: DriveSyncOperation[]): DriveSyncOperation[] {
  const byId = new Map<string, DriveSyncOperation>();
  operations.forEach((operation) => {
    byId.set(operation.id, operation);
  });
  return Array.from(byId.values())
    .sort(compareDriveSyncOperations)
    .slice(-MaxStoredDriveSyncOperations);
}

function compareDriveSyncOperations(a: DriveSyncOperation, b: DriveSyncOperation): number {
  if (a.createdAt !== b.createdAt) return a.createdAt - b.createdAt;
  if (a.deviceId !== b.deviceId) return a.deviceId.localeCompare(b.deviceId);
  if (a.seq !== b.seq) return a.seq - b.seq;
  return a.id.localeCompare(b.id);
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
      const pageUpdatedAt = normalizePageTimestampMap(rawBookmark.pageUpdatedAt);
      pages.forEach((page) => {
        const key = String(page);
        pageUpdatedAt[key] = pageUpdatedAt[key] ?? updatedAt;
      });
      return [[docKey, { pages, updatedAt, pageUpdatedAt } satisfies DriveSyncBookmarkRecord]];
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
  delete state.settings.driveSync.lastPageUpdatedAt[docKey];
  delete state.settings.driveSync.bookmarkUpdatedAt[docKey];
  delete state.settings.driveSync.bookmarkPageUpdatedAt[docKey];
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

function maxTimestamp(map: TimestampMap): number {
  return Object.values(map).reduce((max, timestamp) => Math.max(max, timestamp), 0);
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

function normalizePageTimestampMap(value: unknown): TimestampMap {
  if (!isRecord(value)) return {};
  return Object.fromEntries(
    Object.entries(value).flatMap(([rawPage, rawTimestamp]) => {
      const page = Number(rawPage);
      const timestamp = timestampOr(rawTimestamp, 0);
      return Number.isInteger(page) && page >= 0 && timestamp > 0
        ? [[String(page), timestamp]]
        : [];
    }),
  );
}

function integerInRange(value: unknown, fallback: number, min: number, max: number): number {
  const parsed = typeof value === 'number' && Number.isFinite(value) ? value : fallback;
  return Math.floor(clamp(parsed, min, Math.max(min, max)));
}
