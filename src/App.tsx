import { useCallback, useEffect, useMemo, useRef, useState, type FormEvent, type SetStateAction } from 'react';
import {
  ArrowLeft,
  ArrowUpDown,
  Bookmark,
  ChevronLeft,
  ChevronRight,
  Cloud,
  Copy,
  Download,
  Edit3,
  FileText,
  Folder,
  FolderPlus,
  Grid2X2,
  Link2,
  Maximize2,
  MessageSquare,
  Moon,
  MoreHorizontal,
  Power,
  PanelRightClose,
  PanelRightOpen,
  Plus,
  RefreshCcw,
  Search,
  Settings,
  StretchHorizontal,
  Sun,
  Trash2,
  Upload,
  ZoomIn,
  ZoomOut,
} from 'lucide-react';
import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy, RenderTask } from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.mjs?url';
import {
  type DriveAuthOptions,
  downloadDrivePdf,
  getDriveConfigStatus,
  pickDrivePdf,
  preloadDriveApis,
  readDriveAppDataJsonFile,
  saveDriveAppDataJsonFile,
  type DrivePdfFile,
} from './drive';
import {
  DefaultManualZoom,
  DefaultZoomMode,
  MinManualZoom,
  UnfiledSubjectId,
  clamp,
  composeStoredDocument,
  createInitialAppState,
  getAllStoredDocuments,
  getStoredDocument,
  loadAppState,
  restoreAppStateFromBackup,
  saveAppState,
  serializeAppBackup,
  titleFromFileName,
  type AppState,
  type CopyPacketOptions,
  type DocumentLibraryMetadata,
  type DocumentReaderState,
  type DocumentSourceMetadata,
  type DocumentSourceKind,
  type SortMode,
  type StoredComment,
  type StoredDocument,
  type StoredSubject,
  type Theme,
  type VisibleCopySettingKey,
  type ZoomMode,
} from './storage';
import {
  DriveSyncFileName,
  createDriveSyncEnvelope,
  getDriveSyncFingerprint,
  mergeDriveSyncEnvelope,
  parseDriveSyncEnvelope,
} from './sync';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

type AppScreen = 'library' | 'reader';

type ZoomMetrics = {
  effectiveZoom: number;
  minManualZoom: number;
  maxManualZoom: number;
};

type CopyImageResult = {
  blob: Blob;
  width: number;
  height: number;
};

type DialogState =
  | {
    type: 'text';
    title: string;
    description: string;
    label: string;
    initialValue: string;
    confirmLabel: string;
    onConfirm: (value: string) => void;
  }
  | {
    type: 'confirm';
    title: string;
    description: string;
    confirmLabel: string;
    danger?: boolean;
    onConfirm: () => void;
  };

type SyncStatusKind = 'off' | 'paused' | 'idle' | 'syncing' | 'synced' | 'failed';

type SyncStatusState = {
  kind: SyncStatusKind;
  label: string;
};

type SyncIssueKind = 'auth' | 'drive-read' | 'drive-save' | 'remote-data' | 'network' | 'config' | 'unknown';

type SubjectMenuPosition = {
  top: number;
  left: number;
};

type RuntimePdf = {
  key: string;
  file: File;
  document: PDFDocumentProxy;
};

type LoadedDrivePdf = {
  driveFile: DrivePdfFile;
  file: File;
  document: PDFDocumentProxy;
  source: DocumentSourceMetadata;
};

type PendingDriveImport = LoadedDrivePdf & {
  title: string;
  subjectId: string | null;
};

type DriveImportChoice = {
  title: string;
  subjectId: string | null;
};

type SourceConnectionInfo = {
  label: string;
  tone: 'ready' | 'needs-file' | 'drive' | 'pending';
  description: string;
};

const ZoomStep = 0.05;
const SubjectMenuWidth = 176;
const SubjectMenuEstimatedHeight = 92;
const MaxClipboardHistoryItemBytes = 4 * 1024 * 1024;
const ClipboardHistoryCommitDelayMs = 350;
const CopyImageAttempts = [
  { maxWidth: 1600, maxHeight: 2200 },
  { maxWidth: 1280, maxHeight: 1800 },
  { maxWidth: 1024, maxHeight: 1440 },
  { maxWidth: 768, maxHeight: 1080 },
] as const;

function makeFallbackDocKey(file: File): string {
  return `local:${file.name}:${file.size}:${file.lastModified}`;
}

function makeDriveDocKey(fileId: string): string {
  return `drive:${fileId}`;
}

function makeId(prefix: string): string {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function formatDate(timestamp: number): string {
  return new Intl.DateTimeFormat(undefined, { month: 'short', day: 'numeric' }).format(new Date(timestamp));
}

function formatTime(timestamp: number): string {
  return new Intl.DateTimeFormat(undefined, {
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
  }).format(new Date(timestamp));
}

function formatBackupFileName(date = new Date()): string {
  const stamp = date.toISOString().replace(/[:.]/g, '-').slice(0, 19);
  return `slide-study-backup-${stamp}.json`;
}

function countBookmarkedPages(bookmarks: Record<string, number[]>): number {
  return Object.values(bookmarks).reduce((sum, pages) => sum + pages.length, 0);
}

function createLocalDocumentSource(file: File, pageCount: number): DocumentSourceMetadata {
  return {
    sourceKind: 'local',
    pageCount,
    fileName: file.name,
    fileSize: file.size,
    fileLastModified: file.lastModified,
    driveFileId: null,
    driveName: null,
    driveModifiedTime: null,
    driveSize: null,
  };
}

function createDriveDocumentSource(driveFile: DrivePdfFile, file: File, pageCount: number): DocumentSourceMetadata {
  return {
    sourceKind: 'drive',
    pageCount,
    fileName: file.name,
    fileSize: file.size,
    fileLastModified: file.lastModified,
    driveFileId: driveFile.id,
    driveName: driveFile.name,
    driveModifiedTime: driveFile.modifiedTime,
    driveSize: driveFile.size ?? file.size,
  };
}

function driveFileFromStoredDocument(doc: StoredDocument): DrivePdfFile | null {
  if (doc.sourceKind !== 'drive' || !doc.driveFileId) return null;
  return {
    id: doc.driveFileId,
    name: doc.driveName || doc.fileName,
    mimeType: 'application/pdf',
    size: doc.driveSize ?? doc.fileSize,
    modifiedTime: doc.driveModifiedTime,
  };
}

function normalizeDrivePdfName(name: string): string {
  const trimmed = name.trim() || 'Drive PDF.pdf';
  return trimmed.toLowerCase().endsWith('.pdf') ? trimmed : `${trimmed}.pdf`;
}

function parseDriveTimestamp(modifiedTime: string | null): number {
  if (!modifiedTime) return Date.now();
  const parsed = Date.parse(modifiedTime);
  return Number.isFinite(parsed) ? parsed : Date.now();
}

function drivePdfFileFromBlob(driveFile: DrivePdfFile, blob: Blob): File {
  return new File([blob], normalizeDrivePdfName(driveFile.name), {
    type: 'application/pdf',
    lastModified: parseDriveTimestamp(driveFile.modifiedTime),
  });
}

async function loadDrivePdf(driveFile: DrivePdfFile, authOptions: DriveAuthOptions): Promise<LoadedDrivePdf> {
  const blob = await downloadDrivePdf(driveFile.id, authOptions);
  const file = drivePdfFileFromBlob(driveFile, blob);
  const data = await file.arrayBuffer();
  const document = await pdfjsLib.getDocument({ data }).promise;
  const source = createDriveDocumentSource(driveFile, file, document.numPages);
  return {
    driveFile,
    file,
    document,
    source,
  };
}

function errorMessage(error: unknown, fallback: string): string {
  return error instanceof Error && error.message ? error.message : fallback;
}

function driveImportMessage(error: unknown, fallback: string): string {
  const message = errorMessage(error, fallback);
  const normalized = message.toLowerCase();
  if (normalized.includes('missing config')) {
    return 'Google Drive is not configured for this build.';
  }
  if (isGooglePopupOrAuthPause(normalized)) {
    return 'Google sign-in was closed. Try Drive again when you are ready.';
  }
  if (normalized.includes('could not load https://accounts.google.com') || normalized.includes('could not load https://apis.google.com')) {
    return 'Could not load Google Drive. Check the connection and try again.';
  }
  if (normalized.includes('(404)')) {
    return 'Drive PDF was not found.';
  }
  if (normalized.includes('(401)') || normalized.includes('(403)')) {
    return 'Could not download this Drive PDF. Check that this Google account can open the file.';
  }
  if (normalized.includes('could not download the drive pdf')) {
    return 'Could not download this Drive PDF.';
  }
  if (normalized.includes('failed to fetch') || normalized.includes('network')) {
    return 'Network or Drive request failed.';
  }
  if (normalized.includes('google drive picker could not be loaded')) {
    return 'Google Drive Picker could not be loaded. Refresh the page and try again.';
  }
  if (normalized.includes('could not read the selected drive pdf')) {
    return 'Could not read the selected Drive PDF. Choose the file again.';
  }
  return message;
}

function driveSyncMessage(error: unknown, fallback: string): string {
  const message = errorMessage(error, fallback);
  const normalized = message.toLowerCase();
  if (isGooglePopupOrAuthPause(normalized)) {
    return 'Sync paused. Click Sync now to reconnect.';
  }
  if (normalized.includes('could not read drive sync file')) {
    return 'Could not read Drive sync data. Try Sync again.';
  }
  if (normalized.includes('could not download drive sync file')) {
    return 'Could not download Drive sync data. Try Sync again.';
  }
  if (normalized.includes('could not save drive sync file')) {
    return 'Could not save Drive sync data. Try Sync again.';
  }
  if (normalized.includes('not a supported slide study sync file')) {
    return 'Drive sync data could not be read. You can use Reset remote sync to replace it with this browser data.';
  }
  if (normalized.includes('google identity services could not be loaded')) {
    return 'Could not load Google sign-in. Check the connection and try again.';
  }
  return message;
}

function syncIssueKind(error: unknown, message: string): SyncIssueKind {
  const rawMessage = errorMessage(error, message).toLowerCase();
  const normalizedMessage = message.toLowerCase();
  if (isGooglePopupOrAuthPause(rawMessage) || isGooglePopupOrAuthPause(normalizedMessage)) return 'auth';
  if (rawMessage.includes('missing config') || normalizedMessage.includes('not configured')) return 'config';
  if (rawMessage.includes('not a supported slide study sync file') || normalizedMessage.includes('reset remote')) return 'remote-data';
  if (rawMessage.includes('could not save drive sync file') || normalizedMessage.includes('save drive sync data')) return 'drive-save';
  if (
    rawMessage.includes('could not read drive sync file')
    || rawMessage.includes('could not download drive sync file')
    || normalizedMessage.includes('read drive sync data')
    || normalizedMessage.includes('download drive sync data')
  ) return 'drive-read';
  if (
    rawMessage.includes('failed to fetch')
    || rawMessage.includes('network')
    || normalizedMessage.includes('check the connection')
  ) return 'network';
  return 'unknown';
}

function syncIssueLabel(issue: SyncIssueKind): string {
  switch (issue) {
    case 'auth':
      return 'Google reconnect needed';
    case 'drive-read':
      return 'Drive could not read data';
    case 'drive-save':
      return 'Drive could not save data';
    case 'remote-data':
      return 'Remote data needs reset';
    case 'network':
      return 'Network or Drive request failed';
    case 'config':
      return 'Drive config missing';
    default:
      return 'Sync could not finish';
  }
}

function isGooglePopupOrAuthPause(normalizedMessage: string): boolean {
  return normalizedMessage.includes('sign-in was closed')
    || normalizedMessage.includes('popup_closed')
    || normalizedMessage.includes('popup_failed_to_open')
    || normalizedMessage.includes('failed to open popup')
    || normalizedMessage.includes('popup window')
    || normalizedMessage.includes('blocked by the browser')
    || normalizedMessage.includes('sign-in was closed or blocked')
    || normalizedMessage.includes('interaction_required');
}

function syncStatusDetail(status: SyncStatusState, lastSyncedAt: number | null): string {
  if (status.kind === 'paused') {
    return lastSyncedAt
      ? `Reconnect to sync. Last sync ${formatTime(lastSyncedAt)}`
      : 'Connect to start syncing.';
  }
  if (status.kind === 'off') return 'Drive sync is off.';
  if (status.kind === 'syncing') return 'Working with Google Drive...';
  if (status.kind === 'idle') return 'Waiting to upload changes...';
  if (status.kind === 'failed') return lastSyncedAt
    ? `Last sync ${formatTime(lastSyncedAt)}`
    : 'No sync completed yet.';
  return lastSyncedAt ? `Last sync ${formatTime(lastSyncedAt)}` : 'No sync yet';
}

function resetWindowScroll(): void {
  document.documentElement.scrollTop = 0;
  document.body.scrollTop = 0;
  window.scrollTo(0, 0);
}

function restoreAppViewport(focusTarget: HTMLElement | null): void {
  window.setTimeout(() => {
    resetWindowScroll();
    window.requestAnimationFrame(() => {
      resetWindowScroll();
      if (focusTarget && document.contains(focusTarget)) {
        focusTarget.focus({ preventScroll: true });
      }
    });
  }, 0);
}

function getSourceKindLabel(sourceKind: DocumentSourceKind): string {
  return sourceKind === 'drive' ? 'Drive PDF' : 'Local PDF';
}

function getSourceConnectionInfo(doc: StoredDocument, runtimeDocKey: string | null): SourceConnectionInfo {
  if (doc.sourceKind === 'drive') {
    return doc.driveFileId
      ? {
        label: runtimeDocKey === doc.key ? 'Drive ready' : 'Drive linked',
        tone: 'drive',
        description: runtimeDocKey === doc.key
          ? 'This Drive PDF is open in the current browser session.'
          : 'This PDF can be opened from its Google Drive source.',
      }
      : {
        label: 'Drive pending',
        tone: 'pending',
        description: 'This PDF has a Drive source shape but no Drive file ID yet.',
      };
  }

  if (runtimeDocKey === doc.key) {
    return {
      label: 'Local ready',
      tone: 'ready',
      description: 'The local PDF file is open in this browser session.',
    };
  }

  return {
    label: 'Needs file',
    tone: 'needs-file',
    description: 'Select the local PDF file again to render pages in this browser.',
  };
}

function describeReconnectMismatch(expected: StoredDocument, source: DocumentSourceMetadata): string | null {
  if (expected.pageCount !== source.pageCount) {
    return `Selected PDF has ${source.pageCount} pages; expected ${expected.pageCount}.`;
  }

  if (expected.fileSize > 0 && source.fileSize > 0 && expected.fileSize !== source.fileSize) {
    return 'File size differs from the saved source.';
  }

  if (expected.fileName !== source.fileName) {
    return 'File name differs from the saved source.';
  }

  return null;
}

function getSubjectMenuPosition(anchor: DOMRect, minLeft = 0): SubjectMenuPosition {
  const gutter = 8;
  const rightSideLeft = Math.max(anchor.right + gutter, minLeft + gutter);
  const leftSideLeft = anchor.left - SubjectMenuWidth - gutter;
  const hasRightSpace = rightSideLeft + SubjectMenuWidth <= window.innerWidth - gutter;
  const left = clamp(
    hasRightSpace ? rightSideLeft : leftSideLeft,
    gutter,
    Math.max(gutter, window.innerWidth - SubjectMenuWidth - gutter),
  );
  const top = clamp(
    anchor.top - 4,
    gutter,
    Math.max(gutter, window.innerHeight - SubjectMenuEstimatedHeight - gutter),
  );

  return { top, left };
}

function buildCopyPacketText(
  doc: StoredDocument,
  pageIndex: number,
  comments: StoredComment[],
  options: CopyPacketOptions,
): string {
  const lines = [
    `PDF: ${doc.title}`,
    `Page: ${pageIndex + 1} / ${doc.pageCount}`,
  ];

  if (options.includeComments && comments.length > 0) {
    lines.push('', 'Comments:');
    comments.forEach((comment, index) => {
      lines.push(comments.length === 1 ? comment.body : `Comment ${index + 1}\n${comment.body}`);
      if (index < comments.length - 1) lines.push('');
    });
  }

  return lines.join('\n').trim();
}

function canvasToPngBlob(canvas: HTMLCanvasElement): Promise<Blob | null> {
  return new Promise((resolve) => {
    canvas.toBlob((blob) => resolve(blob), 'image/png');
  });
}

async function renderPdfPageClipboardImage(
  pdf: PDFDocumentProxy,
  pageIndex: number,
): Promise<CopyImageResult | null> {
  const page = await pdf.getPage(pageIndex + 1);
  const baseViewport = page.getViewport({ scale: 1 });

  for (const attempt of CopyImageAttempts) {
    const scale = Math.min(
      attempt.maxWidth / baseViewport.width,
      attempt.maxHeight / baseViewport.height,
    );
    const viewport = page.getViewport({ scale: Math.max(scale, 0.1) });
    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    const context = canvas.getContext('2d');
    if (!context) return null;

    context.fillStyle = '#ffffff';
    context.fillRect(0, 0, canvas.width, canvas.height);

    await page.render({
      canvasContext: context,
      viewport,
    }).promise;

    const blob = await canvasToPngBlob(canvas);
    if (blob && blob.size <= MaxClipboardHistoryItemBytes) {
      return {
        blob,
        width: canvas.width,
        height: canvas.height,
      };
    }
  }

  return null;
}

function canWriteClipboardImage(): boolean {
  if (!navigator.clipboard?.write || typeof ClipboardItem === 'undefined') return false;
  if ('supports' in ClipboardItem && !ClipboardItem.supports('image/png')) return false;
  return true;
}

function wait(ms: number): Promise<void> {
  return new Promise((resolve) => {
    window.setTimeout(resolve, ms);
  });
}

function App() {
  const [stored, setStoredState] = useState<AppState>(() => createInitialAppState());
  const [storageReady, setStorageReady] = useState(false);
  const [screen, setScreen] = useState<AppScreen>('library');
  const [runtimePdf, setRuntimePdf] = useState<RuntimePdf | null>(null);
  const [pageIndex, setPageIndex] = useState(0);
  const [zoomMetrics, setZoomMetrics] = useState<ZoomMetrics>({
    effectiveZoom: DefaultManualZoom,
    minManualZoom: MinManualZoom,
    maxManualZoom: DefaultManualZoom,
  });
  const [thumbnailsOpen, setThumbnailsOpen] = useState(false);
  const [bookmarkedOnly, setBookmarkedOnly] = useState(false);
  const [studyOpen, setStudyOpen] = useState(false);
  const [composerOpen, setComposerOpen] = useState(false);
  const [commentDraft, setCommentDraft] = useState('');
  const [editingCommentId, setEditingCommentId] = useState<string | null>(null);
  const [statusText, setStatusText] = useState('Ready');
  const [query, setQuery] = useState('');
  const [dialog, setDialog] = useState<DialogState | null>(null);
  const [pendingDriveImport, setPendingDriveImport] = useState<PendingDriveImport | null>(null);
  const [drivePickerOpen, setDrivePickerOpen] = useState(false);
  const [driveBusy, setDriveBusy] = useState(false);
  const [syncBusy, setSyncBusy] = useState(false);
  const [syncStatus, setSyncStatus] = useState<SyncStatusState>({ kind: 'off', label: 'Not connected' });
  const [syncIssue, setSyncIssue] = useState<string | null>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const backupInputRef = useRef<HTMLInputElement | null>(null);
  const pendingReconnectDocKeyRef = useRef<string | null>(null);
  const driveBusyRef = useRef(false);
  const syncBusyRef = useRef(false);
  const syncSessionReadyRef = useRef(false);
  const silentReconnectAttemptedRef = useRef(false);
  const syncDebounceRef = useRef<number | null>(null);
  const storedRef = useRef(stored);
  const storedRevisionRef = useRef(0);
  const syncFingerprintRef = useRef('');
  const setStored = useCallback((nextState: SetStateAction<AppState>) => {
    const previous = storedRef.current;
    const resolved = typeof nextState === 'function'
      ? (nextState as (value: AppState) => AppState)(previous)
      : nextState;
    storedRef.current = resolved;
    if (!Object.is(resolved, previous)) {
      storedRevisionRef.current += 1;
    }
    setStoredState(resolved);
  }, []);
  const driveConfigStatus = useMemo(() => getDriveConfigStatus(), []);

  useEffect(() => {
    preloadDriveApis();
  }, []);

  useEffect(() => {
    storedRef.current = stored;
  }, [stored]);

  useEffect(() => {
    const previousScrollRestoration = window.history.scrollRestoration;
    window.history.scrollRestoration = 'manual';
    resetWindowScroll();
    return () => {
      window.history.scrollRestoration = previousScrollRestoration;
    };
  }, []);

  const markDriveFileAccessGranted = useCallback(() => {
    setStored((prev) => (
      prev.settings.driveAuth.hasGrantedFileAccess && prev.settings.driveAuth.hasGrantedAppDataAccess
        ? prev
        : {
          ...prev,
          settings: {
            ...prev.settings,
            driveAuth: {
              ...prev.settings.driveAuth,
              hasGrantedFileAccess: true,
              hasGrantedAppDataAccess: true,
            },
          },
        }
    ));
  }, []);

  useEffect(() => {
    let cancelled = false;
    void loadAppState().then((state) => {
      if (cancelled) return;
      setStored(state);
      setStorageReady(true);
    });

    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!storageReady) return;
    void saveAppState(stored).catch(() => {
      setStatusText('Could not save study data.');
    });
  }, [storageReady, stored]);

  useEffect(() => {
    document.body.classList.toggle('dark', stored.settings.theme === 'dark');
  }, [stored.settings.theme]);

  useEffect(() => {
    document.body.classList.toggle('drive-picker-open', drivePickerOpen);
    if (!drivePickerOpen) {
      restoreAppViewport(null);
    }
    return () => {
      document.body.classList.remove('drive-picker-open');
    };
  }, [drivePickerOpen]);

  useEffect(() => {
    const selected = getStoredDocument(stored, stored.settings.selectedDocKey);
    if (selected) {
      setPageIndex(selected.lastPageIndex);
    }
  }, [stored, stored.settings.selectedDocKey]);

  const selectedDoc = getStoredDocument(stored, stored.settings.selectedDocKey);
  const currentRuntimePdf = runtimePdf && runtimePdf.key === stored.settings.selectedDocKey ? runtimePdf : null;
  const currentZoomMode = selectedDoc?.zoomMode ?? DefaultZoomMode;
  const currentManualZoom = selectedDoc?.manualZoom ?? DefaultManualZoom;
  const renderScalePercent = Math.round(zoomMetrics.effectiveZoom * 100);
  const currentComments = useMemo(
    () => stored.studyData.comments
      .filter((comment) => comment.docKey === stored.settings.selectedDocKey && comment.pageIndex === pageIndex)
      .sort((a, b) => b.updatedAt - a.updatedAt),
    [stored.studyData.comments, stored.settings.selectedDocKey, pageIndex],
  );
  const docBookmarks = useMemo(
    () => (stored.settings.selectedDocKey ? stored.studyData.bookmarks[stored.settings.selectedDocKey] ?? [] : []),
    [stored.studyData.bookmarks, stored.settings.selectedDocKey],
  );
  const currentPageBookmarked = docBookmarks.includes(pageIndex);
  const driveAuthOptions = useMemo<DriveAuthOptions>(() => ({
    hasGrantedFileAccess: stored.settings.driveAuth.hasGrantedFileAccess,
    hasGrantedAppDataAccess: stored.settings.driveAuth.hasGrantedAppDataAccess,
    onTokenGranted: markDriveFileAccessGranted,
  }), [markDriveFileAccessGranted, stored.settings.driveAuth.hasGrantedAppDataAccess, stored.settings.driveAuth.hasGrantedFileAccess]);

  const runDriveSync = useCallback(async ({
    userInitiated = false,
    enable = false,
    resetRemote = false,
    silent = false,
  }: { userInitiated?: boolean; enable?: boolean; resetRemote?: boolean; silent?: boolean } = {}) => {
    if (syncBusyRef.current) {
      if (userInitiated) setStatusText('Drive sync is already running.');
      return;
    }
    if (!driveConfigStatus.configured) {
      const message = `Drive sync is not configured for this build: ${driveConfigStatus.missing.join(', ')}.`;
      if (!silent) {
        setSyncStatus({ kind: 'failed', label: 'Config missing' });
        setSyncIssue(syncIssueLabel('config'));
        setStatusText(message);
      }
      return;
    }

    syncBusyRef.current = true;
    setSyncBusy(true);
    setSyncStatus({ kind: 'syncing', label: silent ? 'Reconnecting...' : 'Syncing...' });
    if (!silent) setSyncIssue(null);
    if (!silent) {
      setStatusText(userInitiated ? 'Syncing with Google Drive...' : 'Auto-syncing with Google Drive...');
    }

    try {
      const withSyncEnabled = (state: AppState): AppState => (
        enable || !state.settings.driveSync.enabled
          ? {
            ...state,
            settings: {
              ...state.settings,
              driveSync: {
                ...state.settings.driveSync,
                enabled: true,
              },
            },
          }
          : state
      );
      const stateForAuth = withSyncEnabled(storedRef.current);
      const authOptions: DriveAuthOptions = {
        ...driveAuthOptions,
        hasGrantedFileAccess: stateForAuth.settings.driveAuth.hasGrantedFileAccess,
        hasGrantedAppDataAccess: stateForAuth.settings.driveAuth.hasGrantedAppDataAccess,
        forceConsent: userInitiated && !stateForAuth.settings.driveAuth.hasGrantedAppDataAccess,
      };
      const remoteFile = await readDriveAppDataJsonFile(DriveSyncFileName, authOptions);
      if (remoteFile && !resetRemote && !parseDriveSyncEnvelope(remoteFile.data)) {
        throw new Error('Drive sync file is not a supported Slide Study sync file. Use Reset remote sync to replace it.');
      }
      let remoteRaw: unknown = remoteFile?.data ?? null;
      let remoteFileId = remoteFile?.id;
      let remoteModifiedTime = remoteFile?.modifiedTime ?? null;
      let finalMerged: AppState | null = null;
      let savedFingerprint = '';
      let needsFollowUpSync = false;

      for (let attempt = 0; attempt < 3; attempt += 1) {
        const sourceState = withSyncEnabled(storedRef.current);
        const sourceRevision = storedRevisionRef.current;
        const merged = resetRemote
          ? sourceState
          : mergeDriveSyncEnvelope(sourceState, remoteRaw);
        const envelope = createDriveSyncEnvelope(merged);
        const saved = await saveDriveAppDataJsonFile(
          DriveSyncFileName,
          envelope,
          authOptions,
          remoteFileId,
        );
        remoteRaw = envelope;
        remoteFileId = saved.id;
        remoteModifiedTime = saved.modifiedTime ?? remoteModifiedTime;
        finalMerged = merged;
        savedFingerprint = getDriveSyncFingerprint(merged);

        const currentState = withSyncEnabled(storedRef.current);
        if (storedRevisionRef.current === sourceRevision) {
          break;
        }

        if (attempt === 2) {
          finalMerged = resetRemote
            ? currentState
            : mergeDriveSyncEnvelope(currentState, remoteRaw);
          needsFollowUpSync = true;
        }
      }

      if (!finalMerged) return;
      const syncedAt = Date.now();
      const finalState: AppState = {
        ...finalMerged,
        settings: {
          ...finalMerged.settings,
          driveAuth: {
            ...finalMerged.settings.driveAuth,
            hasGrantedFileAccess: true,
            hasGrantedAppDataAccess: true,
          },
          driveSync: {
            ...finalMerged.settings.driveSync,
            enabled: true,
            lastSyncedAt: syncedAt,
            lastRemoteModifiedTime: remoteModifiedTime,
          },
        },
      };
      storedRef.current = finalState;
      const finalFingerprint = getDriveSyncFingerprint(finalState);
      const hasUnsavedSyncChanges = needsFollowUpSync && finalFingerprint !== savedFingerprint;
      syncFingerprintRef.current = hasUnsavedSyncChanges ? savedFingerprint : finalFingerprint;
      syncSessionReadyRef.current = true;
      setStored(finalState);
      setSyncStatus(hasUnsavedSyncChanges ? { kind: 'idle', label: 'Sync pending' } : { kind: 'synced', label: 'Synced' });
      setSyncIssue(null);
      if (!silent) {
        setStatusText(hasUnsavedSyncChanges
          ? 'Recent changes will sync next.'
          : resetRemote ? 'Remote sync data was reset.' : 'Drive sync complete.');
      }
    } catch (error) {
      const message = driveSyncMessage(error, 'Could not sync with Google Drive.');
      const normalizedMessage = message.toLowerCase();
      const paused = silent
        || normalizedMessage.includes('sync paused')
        || normalizedMessage.includes('sign-in')
        || normalizedMessage.includes('access')
        || isGooglePopupOrAuthPause(normalizedMessage);
      syncSessionReadyRef.current = false;
      setSyncStatus({ kind: paused ? 'paused' : 'failed', label: paused ? 'Sync paused' : 'Sync failed' });
      setSyncIssue(silent ? null : syncIssueLabel(syncIssueKind(error, message)));
      if (!silent) {
        setStatusText(message);
      }
    } finally {
      syncBusyRef.current = false;
      setSyncBusy(false);
    }
  }, [driveAuthOptions, driveConfigStatus.configured, driveConfigStatus.missing]);

  useEffect(() => {
    if (!storageReady) return;
    syncFingerprintRef.current = getDriveSyncFingerprint(stored);
    setSyncStatus(stored.settings.driveSync.enabled
      ? {
        kind: 'paused',
        label: stored.settings.driveSync.lastSyncedAt ? 'Sync paused' : 'Ready to sync',
      }
      : { kind: 'off', label: 'Not connected' });
    if (!stored.settings.driveSync.enabled) setSyncIssue(null);
  }, [storageReady]);

  useEffect(() => {
    if (!storageReady || silentReconnectAttemptedRef.current || !driveConfigStatus.configured) return undefined;
    const state = storedRef.current;
    if (
      !state.settings.driveSync.enabled
      || !state.settings.driveAuth.hasGrantedFileAccess
      || !state.settings.driveAuth.hasGrantedAppDataAccess
    ) {
      return undefined;
    }
    silentReconnectAttemptedRef.current = true;
    const timeoutId = window.setTimeout(() => {
      void runDriveSync({ silent: true });
    }, 800);
    return () => window.clearTimeout(timeoutId);
  }, [driveConfigStatus.configured, runDriveSync, storageReady]);

  useEffect(() => {
    if (!storageReady || !stored.settings.driveSync.enabled || !syncSessionReadyRef.current || syncBusyRef.current) return undefined;
    const fingerprint = getDriveSyncFingerprint(stored);
    if (!syncFingerprintRef.current) {
      syncFingerprintRef.current = fingerprint;
      return undefined;
    }
    if (fingerprint === syncFingerprintRef.current) return undefined;

    if (syncDebounceRef.current) {
      window.clearTimeout(syncDebounceRef.current);
    }
    setSyncStatus({ kind: 'idle', label: 'Sync pending' });
    setSyncIssue(null);
    syncDebounceRef.current = window.setTimeout(() => {
      syncDebounceRef.current = null;
      void runDriveSync();
    }, 2500);

    return () => {
      if (syncDebounceRef.current) {
        window.clearTimeout(syncDebounceRef.current);
        syncDebounceRef.current = null;
      }
    };
  }, [runDriveSync, storageReady, stored]);

  const visibleDocs = useMemo(() => {
    const docs = getAllStoredDocuments(stored)
      .filter((doc) => {
        if (!stored.settings.selectedSubjectId) return true;
        if (stored.settings.selectedSubjectId === UnfiledSubjectId) return !doc.subjectId;
        return doc.subjectId === stored.settings.selectedSubjectId;
      });
    const normalized = query.trim().toLowerCase();
    const filtered = normalized
      ? docs.filter((doc) => `${doc.title} ${doc.fileName}`.toLowerCase().includes(normalized))
      : docs;
    return filtered.sort((a, b) => {
      if (stored.settings.sortMode === 'title') return a.title.localeCompare(b.title);
      if (stored.settings.sortMode === 'added') return b.createdAt - a.createdAt;
      return b.updatedAt - a.updatedAt;
    });
  }, [stored, stored.settings.selectedSubjectId, stored.settings.sortMode, query]);

  const persistDocPage = useCallback((docKey: string, nextPageIndex: number) => {
    setStored((prev) => {
      const readerState = prev.readerStates[docKey];
      const source = prev.documentSources[docKey];
      if (!readerState || !prev.documents[docKey]) return prev;
      const now = Date.now();
      return {
        ...prev,
        settings: {
          ...prev.settings,
          selectedDocKey: docKey,
          driveSync: source?.sourceKind === 'drive'
            ? {
              ...prev.settings.driveSync,
              lastPageUpdatedAt: {
                ...prev.settings.driveSync.lastPageUpdatedAt,
                [docKey]: now,
              },
            }
            : prev.settings.driveSync,
        },
        readerStates: {
          ...prev.readerStates,
          [docKey]: {
            ...readerState,
            lastPageIndex: nextPageIndex,
          },
        },
      };
    });
  }, []);

  const persistDocZoom = useCallback((docKey: string, zoomMode: ZoomMode, manualZoom?: number) => {
    setStored((prev) => {
      const readerState = prev.readerStates[docKey];
      const documentMeta = prev.documents[docKey];
      if (!readerState || !documentMeta) return prev;
      return {
        ...prev,
        documents: {
          ...prev.documents,
          [docKey]: {
            ...documentMeta,
            updatedAt: Date.now(),
          },
        },
        readerStates: {
          ...prev.readerStates,
          [docKey]: {
            ...readerState,
            zoomMode,
            manualZoom: manualZoom ?? readerState.manualZoom,
          },
        },
      };
    });
  }, []);

  const changeZoomMode = useCallback((zoomMode: ZoomMode) => {
    if (!selectedDoc) return;
    persistDocZoom(selectedDoc.key, zoomMode);
  }, [persistDocZoom, selectedDoc]);

  const changeManualZoom = useCallback((direction: -1 | 1) => {
    if (!selectedDoc) return;
    const { minManualZoom, maxManualZoom, effectiveZoom } = zoomMetrics;
    if (maxManualZoom <= minManualZoom) return;
    const currentZoom = currentZoomMode === 'manual'
      ? clamp(currentManualZoom, minManualZoom, maxManualZoom)
      : clamp(effectiveZoom, minManualZoom, maxManualZoom);
    const nextZoom = clamp(currentZoom + direction * ZoomStep, minManualZoom, maxManualZoom);
    persistDocZoom(selectedDoc.key, 'manual', Number(nextZoom.toFixed(3)));
  }, [currentManualZoom, currentZoomMode, persistDocZoom, selectedDoc, zoomMetrics]);

  const handleZoomMetricsChange = useCallback((nextMetrics: ZoomMetrics) => {
    setZoomMetrics((current) => {
      const sameEffective = Math.abs(current.effectiveZoom - nextMetrics.effectiveZoom) < 0.001;
      const sameMin = Math.abs(current.minManualZoom - nextMetrics.minManualZoom) < 0.001;
      const sameMax = Math.abs(current.maxManualZoom - nextMetrics.maxManualZoom) < 0.001;
      return sameEffective && sameMin && sameMax ? current : nextMetrics;
    });
  }, []);

  const movePage = useCallback((nextPageIndex: number) => {
    if (!selectedDoc) return;
    const bounded = Math.min(Math.max(nextPageIndex, 0), selectedDoc.pageCount - 1);
    setPageIndex(bounded);
    setComposerOpen(false);
    setEditingCommentId(null);
    setCommentDraft('');
    persistDocPage(selectedDoc.key, bounded);
  }, [persistDocPage, selectedDoc]);

  const openPdfFile = useCallback(async (file: File, reconnectDocKey: string | null = null) => {
    if (!file.name.toLowerCase().endsWith('.pdf')) {
      setStatusText('Choose a PDF file.');
      return;
    }
    setStatusText(`Opening ${file.name}...`);
    try {
      const data = await file.arrayBuffer();
      const document = await pdfjsLib.getDocument({ data }).promise;
      const fingerprints = (document as unknown as { fingerprints?: string[] }).fingerprints;
      const detectedKey = fingerprints?.[0] || makeFallbackDocKey(file);
      const reconnectTarget = reconnectDocKey ? getStoredDocument(stored, reconnectDocKey) : null;
      const key = reconnectTarget ? reconnectTarget.key : detectedKey;
      const previous = reconnectTarget ?? getStoredDocument(stored, key);
      const now = Date.now();
      const source = createLocalDocumentSource(file, document.numPages);
      const mismatch = reconnectTarget ? describeReconnectMismatch(reconnectTarget, source) : null;

      if (reconnectTarget && reconnectTarget.pageCount !== source.pageCount) {
        void document.destroy();
        setStatusText(mismatch ?? 'Selected PDF does not match the saved source.');
        return;
      }

      const reader: DocumentReaderState = {
        lastPageIndex: Math.min(previous?.lastPageIndex ?? 0, source.pageCount - 1),
        zoomMode: previous?.zoomMode ?? DefaultZoomMode,
        manualZoom: previous?.manualZoom ?? DefaultManualZoom,
      };
      const library: DocumentLibraryMetadata = {
        key,
        title: previous?.title || titleFromFileName(source.fileName),
        subjectId: previous?.subjectId ?? (stored.settings.selectedSubjectId === UnfiledSubjectId ? null : stored.settings.selectedSubjectId),
        createdAt: previous?.createdAt ?? now,
        updatedAt: now,
      };
      const doc = composeStoredDocument(library, source, reader);
      setRuntimePdf((previousRuntime) => {
        if (previousRuntime?.document && previousRuntime.document !== document) {
          void previousRuntime.document.destroy();
        }
        return { key, file, document };
      });
      setStored((prev) => ({
        ...prev,
        settings: {
          ...prev.settings,
          selectedDocKey: key,
        },
        documents: {
          ...prev.documents,
          [key]: library,
        },
        documentSources: {
          ...prev.documentSources,
          [key]: source,
        },
        readerStates: {
          ...prev.readerStates,
          [key]: reader,
        },
      }));
      setPageIndex(doc.lastPageIndex);
      setScreen('reader');
      setStatusText(reconnectTarget
        ? mismatch
          ? `Reconnected ${doc.title}. ${mismatch}`
          : `Reconnected ${doc.title}.`
        : `Opened ${doc.title}.`);
    } catch {
      setStatusText('Could not open this PDF.');
    }
  }, [stored, stored.settings.selectedSubjectId]);

  const commitLoadedDrivePdf = useCallback((
    loaded: LoadedDrivePdf,
    previous: StoredDocument | null,
    choice: DriveImportChoice | null,
    statusKind: 'added' | 'opened' | 'already-in-library',
  ) => {
    const { driveFile, file, document, source } = loaded;
    const key = previous?.key ?? makeDriveDocKey(driveFile.id);
    const now = Date.now();
    const reader: DocumentReaderState = {
      lastPageIndex: Math.min(previous?.lastPageIndex ?? 0, source.pageCount - 1),
      zoomMode: previous?.zoomMode ?? DefaultZoomMode,
      manualZoom: previous?.manualZoom ?? DefaultManualZoom,
    };
    const library: DocumentLibraryMetadata = {
      key,
      title: choice?.title || previous?.title || titleFromFileName(source.fileName),
      subjectId: previous ? previous.subjectId : choice?.subjectId ?? null,
      createdAt: previous?.createdAt ?? now,
      updatedAt: now,
    };
    const doc = composeStoredDocument(library, source, reader);

    setRuntimePdf((previousRuntime) => {
      if (previousRuntime?.document && previousRuntime.document !== document) {
        void previousRuntime.document.destroy();
      }
      return { key, file, document };
    });
    setStored((prev) => ({
      ...prev,
      settings: {
        ...prev.settings,
        selectedDocKey: key,
      },
      documents: {
        ...prev.documents,
        [key]: library,
      },
      documentSources: {
        ...prev.documentSources,
        [key]: source,
      },
      readerStates: {
        ...prev.readerStates,
        [key]: reader,
      },
    }));
    setPageIndex(doc.lastPageIndex);
    setScreen('reader');
    setStatusText(statusKind === 'already-in-library'
      ? `Already in Library. Opened ${doc.title} from Drive.`
      : statusKind === 'opened'
        ? `Opened ${doc.title} from Drive.`
        : `Added ${doc.title} to Library.`);
  }, []);

  const openDrivePdf = useCallback(async (
    driveFile: DrivePdfFile,
    preferredDocKey: string | null = null,
    statusKind: 'opened' | 'already-in-library' = 'opened',
  ) => {
    setStatusText(`Downloading ${driveFile.name}...`);
    let loaded: LoadedDrivePdf | null = null;

    try {
      loaded = await loadDrivePdf(driveFile, driveAuthOptions);
      const allDocs = getAllStoredDocuments(stored);
      const existingDriveDoc = allDocs.find((doc) => doc.sourceKind === 'drive' && doc.driveFileId === driveFile.id) ?? null;
      const preferredDoc = preferredDocKey ? getStoredDocument(stored, preferredDocKey) : null;
      const previous = preferredDoc ?? existingDriveDoc;
      commitLoadedDrivePdf(loaded, previous, null, previous ? statusKind : 'added');
    } catch (error) {
      if (loaded?.document) void loaded.document.destroy();
      setStatusText(driveImportMessage(error, 'Could not open the Drive PDF.'));
    }
  }, [commitLoadedDrivePdf, driveAuthOptions, stored]);

  const addDrivePdf = useCallback(async () => {
    if (driveBusyRef.current) {
      setStatusText('Google Drive is already working.');
      return;
    }
    if (!driveConfigStatus.configured) {
      setStatusText(`Google Drive is not configured for this build: ${driveConfigStatus.missing.join(', ')}.`);
      return;
    }

    driveBusyRef.current = true;
    setDriveBusy(true);
    const focusTarget = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    setStatusText('Opening Google Drive...');
    try {
      setDrivePickerOpen(true);
      let driveFile: DrivePdfFile | null = null;
      try {
        driveFile = await pickDrivePdf(driveAuthOptions);
      } finally {
        setDrivePickerOpen(false);
        restoreAppViewport(focusTarget);
      }
      if (!driveFile) {
        setStatusText('Drive picker closed.');
        return;
      }

      const existingDriveDoc = getAllStoredDocuments(stored)
        .find((doc) => doc.sourceKind === 'drive' && doc.driveFileId === driveFile.id) ?? null;
      if (existingDriveDoc) {
        await openDrivePdf(driveFile, existingDriveDoc.key, 'already-in-library');
        return;
      }

      setStatusText(`Downloading ${driveFile.name}...`);
      const loaded = await loadDrivePdf(driveFile, driveAuthOptions);
      const defaultSubjectId = stored.settings.selectedSubjectId && stored.settings.selectedSubjectId !== UnfiledSubjectId
        ? stored.settings.selectedSubjectId
        : null;
      setPendingDriveImport((previous) => {
        if (previous?.document && previous.document !== loaded.document) {
          void previous.document.destroy();
        }
        return {
          ...loaded,
          title: titleFromFileName(loaded.source.fileName),
          subjectId: defaultSubjectId,
        };
      });
      setStatusText(`Choose a subject for ${titleFromFileName(loaded.source.fileName)}.`);
    } catch (error) {
      setDrivePickerOpen(false);
      restoreAppViewport(focusTarget);
      setStatusText(driveImportMessage(error, 'Could not open Google Drive.'));
    } finally {
      driveBusyRef.current = false;
      setDriveBusy(false);
    }
  }, [driveAuthOptions, driveConfigStatus.configured, driveConfigStatus.missing, openDrivePdf, stored]);

  const confirmDriveImport = useCallback((choice: DriveImportChoice) => {
    if (!pendingDriveImport) return;
    const existingDriveDoc = getAllStoredDocuments(stored)
      .find((doc) => doc.sourceKind === 'drive' && doc.driveFileId === pendingDriveImport.driveFile.id) ?? null;
    setPendingDriveImport(null);
    commitLoadedDrivePdf(
      pendingDriveImport,
      existingDriveDoc,
      choice,
      existingDriveDoc ? 'already-in-library' : 'added',
    );
  }, [commitLoadedDrivePdf, pendingDriveImport, stored]);

  const cancelDriveImport = useCallback(() => {
    setPendingDriveImport((pending) => {
      if (pending?.document) void pending.document.destroy();
      return null;
    });
    setStatusText('Drive PDF was not added.');
  }, []);

  const requestPdf = useCallback((reconnectDocKey: string | null = null) => {
    pendingReconnectDocKeyRef.current = reconnectDocKey;
    fileInputRef.current?.click();
  }, []);

  const selectStoredDoc = useCallback((doc: StoredDocument) => {
    setStored((prev) => ({
      ...prev,
      settings: {
        ...prev.settings,
        selectedDocKey: doc.key,
      },
    }));
    setPageIndex(doc.lastPageIndex);
    setScreen('reader');
    if (runtimePdf?.key === doc.key) {
      setStatusText(`${doc.title} is ready.`);
      return;
    }

    if (doc.sourceKind === 'drive') {
      setRuntimePdf(null);
      const driveFile = driveFileFromStoredDocument(doc);
      if (!driveFile) {
        setStatusText('This Drive PDF is missing its Drive file ID.');
        return;
      }
      void openDrivePdf(driveFile, doc.key);
      return;
    }

    if (runtimePdf?.key !== doc.key) {
      setRuntimePdf(null);
      setStatusText(getSourceConnectionInfo(doc, null).description);
    }
  }, [openDrivePdf, runtimePdf?.key]);

  const toggleTheme = useCallback(() => {
    setStored((prev) => ({
      ...prev,
      settings: {
        ...prev.settings,
        theme: prev.settings.theme === 'light' ? 'dark' : 'light',
      },
    }));
  }, []);

  const setSelectedSubject = useCallback((subjectId: string | null) => {
    setStored((prev) => ({
      ...prev,
      settings: {
        ...prev.settings,
        selectedSubjectId: subjectId,
      },
    }));
  }, []);

  const setSortMode = useCallback((sortMode: SortMode) => {
    setStored((prev) => ({
      ...prev,
      settings: {
        ...prev.settings,
        sortMode,
      },
    }));
  }, []);

  const updateCopySetting = useCallback((key: VisibleCopySettingKey, value: boolean) => {
    setStored((prev) => ({
      ...prev,
      settings: {
        ...prev.settings,
        copySettings: {
          ...prev.settings.copySettings,
          [key]: value,
        },
      },
    }));
  }, []);

  const exportSyncData = useCallback(() => {
    try {
      const syncEnvelope = createDriveSyncEnvelope(stored);
      const blob = new Blob([JSON.stringify(syncEnvelope, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `slide-study-sync-${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}.json`;
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 0);
      setStatusText('Sync data exported. PDF files are not included.');
    } catch {
      setStatusText('Could not export sync data.');
    }
  }, [stored]);

  const disconnectSync = useCallback(() => {
    setDialog({
      type: 'confirm',
      title: 'Disconnect Drive sync?',
      description: 'Sync will stop only on this browser. Local Library data stays here, and the hidden Drive sync data is not deleted.',
      confirmLabel: 'Disconnect',
      onConfirm: () => {
        syncSessionReadyRef.current = false;
        if (syncDebounceRef.current) {
          window.clearTimeout(syncDebounceRef.current);
          syncDebounceRef.current = null;
        }
        setStored((prev) => ({
          ...prev,
          settings: {
            ...prev.settings,
            driveSync: {
              ...prev.settings.driveSync,
              enabled: false,
            },
          },
        }));
        setSyncStatus({ kind: 'off', label: 'Not connected' });
        setStatusText('Drive sync disconnected on this browser.');
      },
    });
  }, []);

  const resetRemoteSync = useCallback(() => {
    setDialog({
      type: 'confirm',
      title: 'Reset remote sync data?',
      description: 'This replaces the hidden Drive sync data with this browser data. Other devices may receive this version next time they sync. PDF files are not uploaded or deleted.',
      confirmLabel: 'Reset remote',
      danger: true,
      onConfirm: () => {
        void runDriveSync({ userInitiated: true, enable: true, resetRemote: true });
      },
    });
  }, [runDriveSync]);

  const exportBackup = useCallback(() => {
    try {
      const text = serializeAppBackup(stored);
      const blob = new Blob([text], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = formatBackupFileName();
      link.click();
      window.setTimeout(() => URL.revokeObjectURL(url), 0);
      setStatusText('Local backup exported. PDF files are not included.');
    } catch {
      setStatusText('Could not export local backup.');
    }
  }, [stored]);

  const importBackupFile = useCallback(async (file: File) => {
    try {
      const parsed = JSON.parse(await file.text()) as unknown;
      const restored = restoreAppStateFromBackup(parsed);
      const pdfCount = Object.keys(restored.documents).length;
      const commentCount = restored.studyData.comments.length;
      const bookmarkCount = countBookmarkedPages(restored.studyData.bookmarks);

      setDialog({
        type: 'confirm',
        title: 'Import local backup?',
        description: `This will replace current browser data with ${pdfCount} PDFs, ${commentCount} comments, and ${bookmarkCount} bookmarked pages. PDF files are not included.`,
        confirmLabel: 'Import backup',
        danger: true,
        onConfirm: () => {
          if (runtimePdf) {
            void runtimePdf.document.destroy();
          }
          setRuntimePdf(null);
          setScreen('library');
          setStored(restored);
          setStatusText(`Imported local backup with ${pdfCount} PDFs. Reopen local PDF files to render them.`);
        },
      });
    } catch (error) {
      setStatusText(error instanceof Error ? error.message : 'Could not import local backup.');
    }
  }, [runtimePdf]);

  const createSubject = useCallback(() => {
    setDialog({
      type: 'text',
      title: 'Add subject',
      description: 'Create a subject to organize related PDFs.',
      label: 'Subject name',
      initialValue: '',
      confirmLabel: 'Create',
      onConfirm: (value) => {
        const now = Date.now();
        const id = makeId('subject');
        setStored((prev) => ({
          ...prev,
          settings: {
            ...prev.settings,
            selectedSubjectId: id,
          },
          subjects: {
            ...prev.subjects,
            [id]: { id, name: value, createdAt: now, updatedAt: now },
          },
        }));
      },
    });
  }, []);

  const renameSubject = useCallback((subject: StoredSubject) => {
    setDialog({
      type: 'text',
      title: 'Rename subject',
      description: 'Update the subject name. PDFs in this subject will stay linked.',
      label: 'Subject name',
      initialValue: subject.name,
      confirmLabel: 'Save',
      onConfirm: (value) => {
        if (value === subject.name) return;
        setStored((prev) => ({
          ...prev,
          subjects: {
            ...prev.subjects,
            [subject.id]: { ...subject, name: value, updatedAt: Date.now() },
          },
        }));
      },
    });
  }, []);

  const deleteSubject = useCallback((subject: StoredSubject) => {
    setDialog({
      type: 'confirm',
      title: `Delete ${subject.name}?`,
      description: 'The subject will be removed. PDFs stay in Library and move to Unfiled.',
      confirmLabel: 'Delete subject',
      danger: true,
      onConfirm: () => {
        setStored((prev) => {
          const now = Date.now();
          const { [subject.id]: _removed, ...subjects } = prev.subjects;
          const documents = Object.fromEntries(
            Object.entries(prev.documents).map(([key, doc]) => [
              key,
              doc.subjectId === subject.id ? { ...doc, subjectId: null, updatedAt: now } : doc,
            ]),
          );
          return {
            ...prev,
            settings: {
              ...prev.settings,
              selectedSubjectId: prev.settings.selectedSubjectId === subject.id ? null : prev.settings.selectedSubjectId,
              driveSync: {
                ...prev.settings.driveSync,
                pendingSubjectTombstones: {
                  ...prev.settings.driveSync.pendingSubjectTombstones,
                  [subject.id]: now,
                },
              },
            },
            subjects,
            documents,
          };
        });
      },
    });
  }, []);

  const assignDocSubject = useCallback((docKey: string, subjectId: string | null) => {
    setStored((prev) => {
      const doc = prev.documents[docKey];
      if (!doc) return prev;
      return {
        ...prev,
        documents: {
          ...prev.documents,
          [docKey]: { ...doc, subjectId, updatedAt: Date.now() },
        },
      };
    });
  }, []);

  const renameDoc = useCallback((doc: StoredDocument) => {
    setDialog({
      type: 'text',
      title: 'Rename PDF',
      description: 'This changes the Library title only. The original file name is kept.',
      label: 'PDF title',
      initialValue: doc.title,
      confirmLabel: 'Save',
      onConfirm: (value) => {
        if (value === doc.title) return;
        setStored((prev) => {
          const current = prev.documents[doc.key];
          if (!current) return prev;
          return {
            ...prev,
            documents: {
              ...prev.documents,
              [doc.key]: {
                ...current,
                title: value,
                updatedAt: Date.now(),
              },
            },
          };
        });
      },
    });
  }, []);

  const removeDoc = useCallback((doc: StoredDocument) => {
    setDialog({
      type: 'confirm',
      title: `Remove ${doc.title}?`,
      description: 'This removes the PDF from Library and deletes its saved comments and bookmarks in this browser.',
      confirmLabel: 'Remove PDF',
      danger: true,
      onConfirm: () => {
        setStored((prev) => {
          const now = Date.now();
          const { [doc.key]: _removedDoc, ...documents } = prev.documents;
          const { [doc.key]: _removedSource, ...documentSources } = prev.documentSources;
          const { [doc.key]: _removedReaderState, ...readerStates } = prev.readerStates;
          const { [doc.key]: _removedBookmarks, ...bookmarks } = prev.studyData.bookmarks;
          return {
            ...prev,
            settings: {
              ...prev.settings,
              selectedDocKey: prev.settings.selectedDocKey === doc.key ? null : prev.settings.selectedDocKey,
              driveSync: doc.sourceKind === 'drive'
                ? {
                  ...prev.settings.driveSync,
                  pendingDocumentTombstones: {
                    ...prev.settings.driveSync.pendingDocumentTombstones,
                    [doc.key]: now,
                  },
                  lastPageUpdatedAt: Object.fromEntries(
                    Object.entries(prev.settings.driveSync.lastPageUpdatedAt).filter(([key]) => key !== doc.key),
                  ),
                  bookmarkUpdatedAt: Object.fromEntries(
                    Object.entries(prev.settings.driveSync.bookmarkUpdatedAt).filter(([key]) => key !== doc.key),
                  ),
                  bookmarkPageUpdatedAt: Object.fromEntries(
                    Object.entries(prev.settings.driveSync.bookmarkPageUpdatedAt).filter(([key]) => key !== doc.key),
                  ),
                }
                : prev.settings.driveSync,
            },
            documents,
            documentSources,
            readerStates,
            studyData: {
              ...prev.studyData,
              bookmarks,
              comments: prev.studyData.comments.filter((comment) => comment.docKey !== doc.key),
            },
          };
        });
        if (runtimePdf?.key === doc.key) {
          void runtimePdf.document.destroy();
          setRuntimePdf(null);
        }
      },
    });
  }, [runtimePdf]);

  const reconnectDoc = useCallback((doc: StoredDocument) => {
    setStored((prev) => ({
      ...prev,
      settings: {
        ...prev.settings,
        selectedDocKey: doc.key,
      },
    }));
    if (doc.sourceKind === 'drive') {
      const driveFile = driveFileFromStoredDocument(doc);
      if (!driveFile) {
        setStatusText('This Drive PDF is missing its Drive file ID.');
        return;
      }
      void openDrivePdf(driveFile, doc.key);
      return;
    }
    setStatusText(`Choose ${doc.fileName} again.`);
    requestPdf(doc.key);
  }, [openDrivePdf, requestPdf]);

  const toggleBookmark = useCallback(() => {
    if (!stored.settings.selectedDocKey) return;
    setStored((prev) => {
      const docKey = stored.settings.selectedDocKey!;
      const existing = prev.studyData.bookmarks[docKey] ?? [];
      const next = existing.includes(pageIndex)
        ? existing.filter((page) => page !== pageIndex)
        : [...existing, pageIndex].sort((a, b) => a - b);
      const source = prev.documentSources[docKey];
      const now = Date.now();
      return {
        ...prev,
        settings: {
          ...prev.settings,
          driveSync: source?.sourceKind === 'drive'
            ? {
              ...prev.settings.driveSync,
              bookmarkUpdatedAt: {
                ...prev.settings.driveSync.bookmarkUpdatedAt,
                [docKey]: now,
              },
              bookmarkPageUpdatedAt: {
                ...prev.settings.driveSync.bookmarkPageUpdatedAt,
                [docKey]: {
                  ...prev.settings.driveSync.bookmarkPageUpdatedAt[docKey],
                  [String(pageIndex)]: now,
                },
              },
            }
            : prev.settings.driveSync,
        },
        studyData: {
          ...prev.studyData,
          bookmarks: {
            ...prev.studyData.bookmarks,
            [docKey]: next,
          },
        },
      };
    });
  }, [pageIndex, stored.settings.selectedDocKey]);

  const openNewComment = useCallback(() => {
    setStudyOpen(true);
    setEditingCommentId(null);
    setCommentDraft('');
    setComposerOpen(true);
  }, []);

  const saveComment = useCallback(() => {
    if (!stored.settings.selectedDocKey || !commentDraft.trim()) return;
    const normalized = commentDraft.trim();
    setStored((prev) => {
      const now = Date.now();
      if (editingCommentId) {
        return {
          ...prev,
          studyData: {
            ...prev.studyData,
            comments: prev.studyData.comments.map((comment) => (
              comment.id === editingCommentId
                ? { ...comment, body: normalized, updatedAt: now }
                : comment
            )),
          },
        };
      }
      return {
        ...prev,
        studyData: {
          ...prev.studyData,
          comments: [
            ...prev.studyData.comments,
            {
              id: makeId('comment'),
              docKey: stored.settings.selectedDocKey!,
              pageIndex,
              body: normalized,
              createdAt: now,
              updatedAt: now,
            },
          ],
        },
      };
    });
    setComposerOpen(false);
    setEditingCommentId(null);
    setCommentDraft('');
  }, [commentDraft, editingCommentId, pageIndex, stored.settings.selectedDocKey]);

  const editComment = useCallback((comment: StoredComment) => {
    setStudyOpen(true);
    setEditingCommentId(comment.id);
    setCommentDraft(comment.body);
    setComposerOpen(true);
  }, []);

  const deleteComment = useCallback((commentId: string) => {
    setDialog({
      type: 'confirm',
      title: 'Delete comment?',
      description: 'This comment will be removed from the current browser data.',
      confirmLabel: 'Delete comment',
      danger: true,
      onConfirm: () => {
        setStored((prev) => {
          const target = prev.studyData.comments.find((comment) => comment.id === commentId);
          const source = target ? prev.documentSources[target.docKey] : null;
          const now = Date.now();
          return {
            ...prev,
            settings: {
              ...prev.settings,
              driveSync: source?.sourceKind === 'drive'
                ? {
                  ...prev.settings.driveSync,
                  pendingCommentTombstones: {
                    ...prev.settings.driveSync.pendingCommentTombstones,
                    [commentId]: now,
                  },
                }
                : prev.settings.driveSync,
            },
            studyData: {
              ...prev.studyData,
              comments: prev.studyData.comments.filter((comment) => comment.id !== commentId),
            },
          };
        });
        if (editingCommentId === commentId) {
          setEditingCommentId(null);
          setCommentDraft('');
          setComposerOpen(false);
        }
      },
    });
  }, [editingCommentId]);

  const copyPageStudyPacket = useCallback(async () => {
    if (!selectedDoc) return;
    const packet = buildCopyPacketText(
      selectedDoc,
      pageIndex,
      currentComments,
      stored.settings.copySettings,
    );

    if (!navigator.clipboard?.writeText) {
      setStatusText('Clipboard is not available in this browser.');
      return;
    }

    try {
      await navigator.clipboard.writeText(packet);
      await wait(ClipboardHistoryCommitDelayMs);

      if (!stored.settings.copySettings.includePageImage) {
        setStatusText('Copied page text.');
        return;
      }

      if (!currentRuntimePdf) {
        setStatusText('Copied text. Reopen the PDF to copy the page image.');
        return;
      }

      if (!canWriteClipboardImage()) {
        setStatusText('Copied text. Image copy is not supported in this browser.');
        return;
      }

      const image = await renderPdfPageClipboardImage(currentRuntimePdf.document, pageIndex);
      if (!image) {
        setStatusText('Copied text. Page image was too large for clipboard history.');
        return;
      }

      await navigator.clipboard.write([
        new ClipboardItem({ 'image/png': image.blob }),
      ]);
      setStatusText(`Copied text + page image (${image.width}x${image.height}). Use Win+V for both.`);
    } catch {
      setStatusText('Could not copy this page.');
    }
  }, [currentComments, currentRuntimePdf, pageIndex, selectedDoc, stored.settings.copySettings]);

  return (
    <div className="app-shell">
      <div className={`app-content-layer ${drivePickerOpen ? 'picker-active' : ''}`}>
      <input
        ref={fileInputRef}
        type="file"
        accept="application/pdf,.pdf"
        hidden
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) {
            const reconnectDocKey = pendingReconnectDocKeyRef.current;
            pendingReconnectDocKeyRef.current = null;
            void openPdfFile(file, reconnectDocKey);
          }
          event.target.value = '';
        }}
      />
      <input
        ref={backupInputRef}
        type="file"
        accept="application/json,.json"
        hidden
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) {
            void importBackupFile(file);
          }
          event.target.value = '';
        }}
      />
      <TopBar
        screen={screen}
        theme={stored.settings.theme}
        statusText={statusText}
        onOpenSettings={() => setSettingsOpen(true)}
        onToggleTheme={toggleTheme}
      />
      {screen === 'library' ? (
        <LibraryScreen
          docs={visibleDocs}
          allDocs={getAllStoredDocuments(stored)}
          subjects={Object.values(stored.subjects).sort((a, b) => a.name.localeCompare(b.name))}
          runtimeDocKey={runtimePdf?.key ?? null}
          selectedDocKey={stored.settings.selectedDocKey}
          selectedSubjectId={stored.settings.selectedSubjectId}
          sortMode={stored.settings.sortMode}
          query={query}
          onQueryChange={setQuery}
          onSelectSubject={setSelectedSubject}
          onCreateSubject={createSubject}
          onRenameSubject={renameSubject}
          onDeleteSubject={deleteSubject}
          onSortModeChange={setSortMode}
          onAddPdf={requestPdf}
          onAddDrivePdf={addDrivePdf}
          driveBusy={driveBusy}
          driveConfigured={driveConfigStatus.configured}
          onSelectDoc={selectStoredDoc}
          onAssignDocSubject={assignDocSubject}
          onRenameDoc={renameDoc}
          onReconnectDoc={reconnectDoc}
          onRemoveDoc={removeDoc}
          bookmarks={stored.studyData.bookmarks}
          comments={stored.studyData.comments}
        />
      ) : (
        <ReaderScreen
          doc={selectedDoc}
          runtimePdf={currentRuntimePdf}
          pageIndex={pageIndex}
          zoomMode={currentZoomMode}
          manualZoom={currentManualZoom}
          zoomMetrics={zoomMetrics}
          renderScalePercent={renderScalePercent}
          thumbnailsOpen={thumbnailsOpen}
          bookmarkedOnly={bookmarkedOnly}
          studyOpen={studyOpen}
          comments={currentComments}
          bookmarks={docBookmarks}
          currentPageBookmarked={currentPageBookmarked}
          composerOpen={composerOpen}
          commentDraft={commentDraft}
          editingCommentId={editingCommentId}
          onOpenLibrary={() => setScreen('library')}
          onRequestPdf={requestPdf}
          onReconnectDoc={reconnectDoc}
          onMovePage={movePage}
          onZoomModeChange={changeZoomMode}
          onManualZoomStep={changeManualZoom}
          onZoomMetricsChange={handleZoomMetricsChange}
          onToggleThumbnails={() => setThumbnailsOpen((open) => !open)}
          onToggleBookmarkedOnly={() => setBookmarkedOnly((enabled) => !enabled)}
          onToggleStudy={() => setStudyOpen((open) => !open)}
          onToggleBookmark={toggleBookmark}
          onCopyPage={copyPageStudyPacket}
          onOpenNewComment={openNewComment}
          onCommentDraftChange={setCommentDraft}
          onSaveComment={saveComment}
          onCancelComment={() => {
            setComposerOpen(false);
            setEditingCommentId(null);
            setCommentDraft('');
          }}
          onEditComment={editComment}
          onDeleteComment={deleteComment}
        />
      )}
      {settingsOpen && (
        <SettingsDialog
          copySettings={stored.settings.copySettings}
          syncEnabled={stored.settings.driveSync.enabled}
          syncBusy={syncBusy}
          syncStatus={syncStatus}
          syncIssue={syncIssue}
          lastSyncedAt={stored.settings.driveSync.lastSyncedAt}
          onToggleCopySetting={updateCopySetting}
          onConnectSync={() => void runDriveSync({ userInitiated: true, enable: true })}
          onSyncNow={() => void runDriveSync({ userInitiated: true })}
          onExportSyncData={exportSyncData}
          onDisconnectSync={disconnectSync}
          onResetRemoteSync={resetRemoteSync}
          onExportBackup={exportBackup}
          onImportBackup={() => backupInputRef.current?.click()}
          onClose={() => setSettingsOpen(false)}
        />
      )}
      {pendingDriveImport && (
        <DriveImportDialog
          pending={pendingDriveImport}
          subjects={Object.values(stored.subjects).sort((a, b) => a.name.localeCompare(b.name))}
          onConfirm={confirmDriveImport}
          onClose={cancelDriveImport}
        />
      )}
      {dialog && (
        <AppDialog
          dialog={dialog}
          onClose={() => setDialog(null)}
        />
      )}
      </div>
    </div>
  );
}

function TopBar({
  screen,
  theme,
  statusText,
  onOpenSettings,
  onToggleTheme,
}: {
  screen: AppScreen;
  theme: Theme;
  statusText: string;
  onOpenSettings: () => void;
  onToggleTheme: () => void;
}) {
  return (
    <header className="topbar">
      <div className="brand">Slide Study</div>
      <div className="location">{screen === 'reader' ? 'Reader' : 'Library'}</div>
      <div className="status">{statusText}</div>
      <div className="topbar-spacer" />
      <IconButton
        label="Settings"
        onClick={onOpenSettings}
        icon={Settings}
      />
      <IconButton
        label={theme === 'dark' ? 'Light mode' : 'Dark mode'}
        onClick={onToggleTheme}
        icon={theme === 'dark' ? Sun : Moon}
      />
    </header>
  );
}

function LibraryScreen({
  docs,
  allDocs,
  subjects,
  runtimeDocKey,
  selectedDocKey,
  selectedSubjectId,
  sortMode,
  query,
  bookmarks,
  comments,
  onQueryChange,
  onSelectSubject,
  onCreateSubject,
  onRenameSubject,
  onDeleteSubject,
  onSortModeChange,
  onAddPdf,
  onAddDrivePdf,
  driveBusy,
  driveConfigured,
  onSelectDoc,
  onAssignDocSubject,
  onRenameDoc,
  onReconnectDoc,
  onRemoveDoc,
}: {
  docs: StoredDocument[];
  allDocs: StoredDocument[];
  subjects: StoredSubject[];
  runtimeDocKey: string | null;
  selectedDocKey: string | null;
  selectedSubjectId: string | null;
  sortMode: SortMode;
  query: string;
  bookmarks: Record<string, number[]>;
  comments: StoredComment[];
  onQueryChange: (value: string) => void;
  onSelectSubject: (subjectId: string | null) => void;
  onCreateSubject: () => void;
  onRenameSubject: (subject: StoredSubject) => void;
  onDeleteSubject: (subject: StoredSubject) => void;
  onSortModeChange: (mode: SortMode) => void;
  onAddPdf: () => void;
  onAddDrivePdf: () => void;
  driveBusy: boolean;
  driveConfigured: boolean;
  onSelectDoc: (doc: StoredDocument) => void;
  onAssignDocSubject: (docKey: string, subjectId: string | null) => void;
  onRenameDoc: (doc: StoredDocument) => void;
  onReconnectDoc: (doc: StoredDocument) => void;
  onRemoveDoc: (doc: StoredDocument) => void;
}) {
  const unfiledCount = allDocs.filter((doc) => !doc.subjectId).length;
  const selectedSubject = selectedSubjectId ? subjects.find((subject) => subject.id === selectedSubjectId) : null;
  const driveDisabled = !driveConfigured || driveBusy;
  const driveTitle = driveBusy
    ? 'Google Drive is opening'
    : driveConfigured
      ? 'Add from Drive'
      : 'Google Drive config missing';

  return (
    <main className="content">
      <section className="library-layout">
        <aside className="sidebar">
          <div className="sidebar-header">
            <div className="section-title">Library</div>
            <IconButton label="Add subject" icon={FolderPlus} onClick={onCreateSubject} />
          </div>
          <label className="search-box">
            <Search size={16} />
            <input value={query} onChange={(event) => onQueryChange(event.target.value)} placeholder="Search PDFs" />
          </label>
          <div className="sidebar-section-head">
            <div className="section-title">Subjects</div>
            <span className="muted">{subjects.length}</span>
          </div>
          <button className={`subject ${selectedSubjectId === null ? 'active' : ''}`} onClick={() => onSelectSubject(null)}>
            <span>All</span>
            <span>{allDocs.length}</span>
          </button>
          <button className={`subject ${selectedSubjectId === UnfiledSubjectId ? 'active' : ''}`} onClick={() => onSelectSubject(UnfiledSubjectId)}>
            <span>Unfiled</span>
            <span>{unfiledCount}</span>
          </button>
          <div className="subject-list">
            {subjects.map((subject) => (
              <SubjectRow
                key={subject.id}
                subject={subject}
                active={selectedSubjectId === subject.id}
                count={allDocs.filter((doc) => doc.subjectId === subject.id).length}
                onSelect={onSelectSubject}
                onRename={onRenameSubject}
                onDelete={onDeleteSubject}
              />
            ))}
          </div>
        </aside>
        <section className="library-main">
          <div className="library-head">
            <div>
              <h1 className="library-title">{selectedSubject?.name ?? (selectedSubjectId === UnfiledSubjectId ? 'Unfiled' : 'Study Library')}</h1>
              <p>{docs.length} PDF{docs.length === 1 ? '' : 's'} - comments, bookmarks, zoom, and last page stay in this browser.</p>
            </div>
            <div className="library-actions">
              <div className="toolbar-group compact-group">
                <ArrowUpDown size={16} />
                <button className={`ghost-btn ${sortMode === 'recent' ? 'selected' : ''}`} onClick={() => onSortModeChange('recent')}>Recent</button>
                <button className={`ghost-btn ${sortMode === 'title' ? 'selected' : ''}`} onClick={() => onSortModeChange('title')}>Title</button>
                <button className={`ghost-btn ${sortMode === 'added' ? 'selected' : ''}`} onClick={() => onSortModeChange('added')}>Added</button>
              </div>
              <button
                className="ghost-btn"
                disabled={driveDisabled}
                title={driveTitle}
                onClick={onAddDrivePdf}
              >
                <Cloud size={16} />
                {driveBusy ? 'Opening...' : 'Drive'}
              </button>
              <button className="primary-btn" onClick={onAddPdf}>
                <Upload size={16} />
                Add PDF
              </button>
            </div>
          </div>
          {docs.length === 0 ? (
            <div className="library-empty">
              <FileText size={28} />
              <div>
                <strong>No PDFs here yet.</strong>
                <p>Add a PDF or choose another subject.</p>
              </div>
              <div className="library-empty-actions">
                <button
                  className="ghost-btn"
                  disabled={driveDisabled}
                  title={driveTitle}
                  onClick={onAddDrivePdf}
                >
                  <Cloud size={16} />
                  {driveBusy ? 'Opening...' : 'Drive'}
                </button>
                <button className="ghost-btn" onClick={onAddPdf}>
                  <Upload size={16} />
                  Add PDF
                </button>
              </div>
            </div>
          ) : (
            <div className="pdf-grid">
              {docs.map((doc) => (
                <PdfCard
                  key={doc.key}
                  doc={doc}
                  active={doc.key === selectedDocKey}
                  subjects={subjects}
                  comments={comments.filter((comment) => comment.docKey === doc.key).length}
                  bookmarks={(bookmarks[doc.key] ?? []).length}
                  sourceConnection={getSourceConnectionInfo(doc, runtimeDocKey)}
                  onOpen={onSelectDoc}
                  onAssignSubject={onAssignDocSubject}
                  onRename={onRenameDoc}
                  onReconnect={onReconnectDoc}
                  onRemove={onRemoveDoc}
                />
              ))}
            </div>
          )}
        </section>
      </section>
    </main>
  );
}

function SubjectRow({
  subject,
  active,
  count,
  onSelect,
  onRename,
  onDelete,
}: {
  subject: StoredSubject;
  active: boolean;
  count: number;
  onSelect: (subjectId: string) => void;
  onRename: (subject: StoredSubject) => void;
  onDelete: (subject: StoredSubject) => void;
}) {
  const [menuOpen, setMenuOpen] = useState(false);
  const [menuPosition, setMenuPosition] = useState<SubjectMenuPosition>({ top: 0, left: 0 });
  const rowRef = useRef<HTMLDivElement | null>(null);
  const actionRef = useRef<HTMLDivElement | null>(null);
  const menuRef = useRef<HTMLDivElement | null>(null);

  const updateMenuPosition = useCallback(() => {
    const anchor = actionRef.current?.getBoundingClientRect();
    if (!anchor) return;
    const sidebarRight = actionRef.current?.closest('.sidebar')?.getBoundingClientRect().right ?? 0;
    setMenuPosition(getSubjectMenuPosition(anchor, sidebarRight));
  }, []);

  const toggleMenu = useCallback(() => {
    setMenuOpen((open) => {
      if (open) return false;
      const anchor = actionRef.current?.getBoundingClientRect();
      const sidebarRight = actionRef.current?.closest('.sidebar')?.getBoundingClientRect().right ?? 0;
      if (anchor) setMenuPosition(getSubjectMenuPosition(anchor, sidebarRight));
      return true;
    });
  }, []);

  useEffect(() => {
    if (!menuOpen) return undefined;

    const handlePointerDown = (event: PointerEvent) => {
      const target = event.target;
      if (!(target instanceof Node)) return;
      if (rowRef.current?.contains(target) || menuRef.current?.contains(target)) return;
      setMenuOpen(false);
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setMenuOpen(false);
    };

    window.addEventListener('pointerdown', handlePointerDown);
    window.addEventListener('keydown', handleKeyDown);
    window.addEventListener('resize', updateMenuPosition);
    window.addEventListener('scroll', updateMenuPosition, true);

    return () => {
      window.removeEventListener('pointerdown', handlePointerDown);
      window.removeEventListener('keydown', handleKeyDown);
      window.removeEventListener('resize', updateMenuPosition);
      window.removeEventListener('scroll', updateMenuPosition, true);
    };
  }, [menuOpen, updateMenuPosition]);

  return (
    <div ref={rowRef} className={`subject-row ${active ? 'active' : ''} ${menuOpen ? 'menu-open' : ''}`}>
      <button className="subject-main" onClick={() => onSelect(subject.id)}>
        <span>{subject.name}</span>
        <span>{count}</span>
      </button>
      <div ref={actionRef} className="subject-actions menu-wrap">
        <IconButton label={`${subject.name} actions`} icon={MoreHorizontal} active={menuOpen} onClick={toggleMenu} />
        {menuOpen && (
          <div
            ref={menuRef}
            className="overflow-menu subject-menu"
            style={{ top: menuPosition.top, left: menuPosition.left }}
            role="menu"
          >
            <button role="menuitem" onClick={() => { setMenuOpen(false); onRename(subject); }}><Edit3 size={16} /> Rename</button>
            <button role="menuitem" className="danger" onClick={() => { setMenuOpen(false); onDelete(subject); }}><Trash2 size={16} /> Delete subject</button>
          </div>
        )}
      </div>
    </div>
  );
}

function PdfCard({
  doc,
  active,
  subjects,
  comments,
  bookmarks,
  sourceConnection,
  onOpen,
  onAssignSubject,
  onRename,
  onReconnect,
  onRemove,
}: {
  doc: StoredDocument;
  active: boolean;
  subjects: StoredSubject[];
  comments: number;
  bookmarks: number;
  sourceConnection: SourceConnectionInfo;
  onOpen: (doc: StoredDocument) => void;
  onAssignSubject: (docKey: string, subjectId: string | null) => void;
  onRename: (doc: StoredDocument) => void;
  onReconnect: (doc: StoredDocument) => void;
  onRemove: (doc: StoredDocument) => void;
}) {
  const [menuOpen, setMenuOpen] = useState(false);
  const subjectName = subjects.find((subject) => subject.id === doc.subjectId)?.name ?? 'Unfiled';

  return (
    <article className={`pdf-card ${active ? 'active' : ''}`}>
      <button className="pdf-card-main" onClick={() => onOpen(doc)}>
        <div className="pdf-card-top">
          <div>
            <div className="pdf-title">{doc.title}</div>
            <div className="muted">{doc.fileName}</div>
          </div>
          <span className="subject-pill">{subjectName}</span>
        </div>
        <div className="tag-row">
          <span className="tag"><FileText size={14} /> {doc.pageCount} pages</span>
          <span className={`tag source-tag ${sourceConnection.tone}`} title={sourceConnection.description}>
            <Link2 size={14} /> {sourceConnection.label}
          </span>
          <span className="tag">Last {doc.lastPageIndex + 1}</span>
          {comments > 0 && <span className="tag"><MessageSquare size={14} /> {comments}</span>}
          {bookmarks > 0 && <span className="tag"><Bookmark size={14} /> {bookmarks}</span>}
        </div>
      </button>
      <div className="pdf-card-footer">
        <span className="muted">Added {formatDate(doc.createdAt)}</span>
        <div className="menu-wrap">
          <IconButton label="PDF actions" icon={MoreHorizontal} active={menuOpen} onClick={() => setMenuOpen((open) => !open)} />
          {menuOpen && (
            <div className="overflow-menu">
              <button onClick={() => { setMenuOpen(false); onOpen(doc); }}><FileText size={16} /> Open</button>
              <button onClick={() => { setMenuOpen(false); onRename(doc); }}><Edit3 size={16} /> Rename</button>
              <label>
                <Folder size={16} />
                <select
                  value={doc.subjectId ?? ''}
                  onChange={(event) => onAssignSubject(doc.key, event.target.value || null)}
                >
                  <option value="">Unfiled</option>
                  {subjects.map((subject) => (
                    <option key={subject.id} value={subject.id}>{subject.name}</option>
                  ))}
                </select>
              </label>
              <button onClick={() => { setMenuOpen(false); onReconnect(doc); }}>
                <Link2 size={16} /> {doc.sourceKind === 'drive' ? 'Open from Drive' : 'Reconnect file'}
              </button>
              <div className="menu-separator" />
              <button className="danger" onClick={() => { setMenuOpen(false); onRemove(doc); }}><Trash2 size={16} /> Remove</button>
            </div>
          )}
        </div>
      </div>
    </article>
  );
}

function ReaderScreen({
  doc,
  runtimePdf,
  pageIndex,
  zoomMode,
  manualZoom,
  zoomMetrics,
  renderScalePercent,
  thumbnailsOpen,
  bookmarkedOnly,
  studyOpen,
  comments,
  bookmarks,
  currentPageBookmarked,
  composerOpen,
  commentDraft,
  editingCommentId,
  onOpenLibrary,
  onRequestPdf,
  onReconnectDoc,
  onMovePage,
  onZoomModeChange,
  onManualZoomStep,
  onZoomMetricsChange,
  onToggleThumbnails,
  onToggleBookmarkedOnly,
  onToggleStudy,
  onToggleBookmark,
  onCopyPage,
  onOpenNewComment,
  onCommentDraftChange,
  onSaveComment,
  onCancelComment,
  onEditComment,
  onDeleteComment,
}: {
  doc: StoredDocument | null;
  runtimePdf: RuntimePdf | null;
  pageIndex: number;
  zoomMode: ZoomMode;
  manualZoom: number;
  zoomMetrics: ZoomMetrics;
  renderScalePercent: number;
  thumbnailsOpen: boolean;
  bookmarkedOnly: boolean;
  studyOpen: boolean;
  comments: StoredComment[];
  bookmarks: number[];
  currentPageBookmarked: boolean;
  composerOpen: boolean;
  commentDraft: string;
  editingCommentId: string | null;
  onOpenLibrary: () => void;
  onRequestPdf: () => void;
  onReconnectDoc: (doc: StoredDocument) => void;
  onMovePage: (pageIndex: number) => void;
  onZoomModeChange: (mode: ZoomMode) => void;
  onManualZoomStep: (direction: -1 | 1) => void;
  onZoomMetricsChange: (metrics: ZoomMetrics) => void;
  onToggleThumbnails: () => void;
  onToggleBookmarkedOnly: () => void;
  onToggleStudy: () => void;
  onToggleBookmark: () => void;
  onCopyPage: () => void;
  onOpenNewComment: () => void;
  onCommentDraftChange: (value: string) => void;
  onSaveComment: () => void;
  onCancelComment: () => void;
  onEditComment: (comment: StoredComment) => void;
  onDeleteComment: (commentId: string) => void;
}) {
  const layoutClass = [
    'reader-layout',
    thumbnailsOpen ? 'thumbs-open' : '',
    studyOpen ? 'study-open' : '',
    studyOpen ? 'study-normal' : '',
  ].join(' ');
  const canZoomOut = zoomMetrics.effectiveZoom > zoomMetrics.minManualZoom + 0.004;
  const canZoomIn = zoomMetrics.effectiveZoom < zoomMetrics.maxManualZoom - 0.004;

  if (!doc) {
    return (
      <main className="content">
        <div className="reader-empty">
          <strong>No PDF selected.</strong>
          <button className="primary-btn" onClick={onRequestPdf}>Add PDF</button>
        </div>
      </main>
    );
  }

  return (
    <main className="content">
      <section className={layoutClass}>
        {thumbnailsOpen ? (
          <ThumbnailPanel
            doc={doc}
            runtimePdf={runtimePdf}
            pageIndex={pageIndex}
            bookmarks={bookmarks}
            bookmarkedOnly={bookmarkedOnly}
            onToggleBookmarkedOnly={onToggleBookmarkedOnly}
            onMovePage={onMovePage}
          />
        ) : <div />}
        <section className="reader-main">
          <ReaderToolbar
            doc={doc}
            pageIndex={pageIndex}
            zoomMode={zoomMode}
            renderScalePercent={renderScalePercent}
            canZoomOut={canZoomOut}
            canZoomIn={canZoomIn}
            thumbnailsOpen={thumbnailsOpen}
            onOpenLibrary={onOpenLibrary}
            onMovePage={onMovePage}
            onZoomModeChange={onZoomModeChange}
            onManualZoomStep={onManualZoomStep}
            onToggleThumbnails={onToggleThumbnails}
          />
          {runtimePdf ? (
            <PdfPageView
              pdf={runtimePdf.document}
              pageIndex={pageIndex}
              zoomMode={zoomMode}
              manualZoom={manualZoom}
              onZoomMetricsChange={onZoomMetricsChange}
            />
          ) : (
            <div className="pdf-stage">
              <div className="missing-pdf">
                <FileText size={26} />
                <strong>{doc.sourceKind === 'drive' ? 'Open this Drive PDF.' : 'Reconnect this local PDF.'}</strong>
                <p>{getSourceConnectionInfo(doc, null).description}</p>
                <div className="source-detail">
                  <span>{getSourceKindLabel(doc.sourceKind)}</span>
                  <span>{doc.fileName}</span>
                </div>
                <button className="primary-btn" onClick={() => onReconnectDoc(doc)}>
                  {doc.sourceKind === 'drive' ? 'Open from Drive' : 'Choose PDF'}
                </button>
              </div>
            </div>
          )}
        </section>
        {studyOpen ? (
          <StudyPanel
            pageIndex={pageIndex}
            comments={comments}
            composerOpen={composerOpen}
            commentDraft={commentDraft}
            editingCommentId={editingCommentId}
            onClose={onToggleStudy}
            onOpenNewComment={onOpenNewComment}
            onCommentDraftChange={onCommentDraftChange}
            onSaveComment={onSaveComment}
            onCancelComment={onCancelComment}
            onEditComment={onEditComment}
            onDeleteComment={onDeleteComment}
          />
        ) : (
          <div className="study-panel-placeholder" aria-hidden="true" />
        )}
        <StudyRail
          studyOpen={studyOpen}
          hasComments={comments.length > 0}
          bookmarked={currentPageBookmarked}
          onToggleStudy={onToggleStudy}
          onOpenNewComment={onOpenNewComment}
          onToggleBookmark={onToggleBookmark}
          onCopyPage={onCopyPage}
        />
      </section>
    </main>
  );
}

function ReaderToolbar({
  doc,
  pageIndex,
  zoomMode,
  renderScalePercent,
  canZoomOut,
  canZoomIn,
  thumbnailsOpen,
  onOpenLibrary,
  onMovePage,
  onZoomModeChange,
  onManualZoomStep,
  onToggleThumbnails,
}: {
  doc: StoredDocument;
  pageIndex: number;
  zoomMode: ZoomMode;
  renderScalePercent: number;
  canZoomOut: boolean;
  canZoomIn: boolean;
  thumbnailsOpen: boolean;
  onOpenLibrary: () => void;
  onMovePage: (pageIndex: number) => void;
  onZoomModeChange: (mode: ZoomMode) => void;
  onManualZoomStep: (direction: -1 | 1) => void;
  onToggleThumbnails: () => void;
}) {
  return (
    <div className="reader-toolbar">
      <IconButton label="Library" icon={ArrowLeft} onClick={onOpenLibrary} />
      <div className="doc-meta">
        <div className="doc-title">{doc.title}</div>
        <div className="muted">{pageIndex + 1} / {doc.pageCount}</div>
      </div>
      <IconButton label={thumbnailsOpen ? 'Hide thumbnails' : 'Show thumbnails'} icon={Grid2X2} active={thumbnailsOpen} onClick={onToggleThumbnails} />
      <div className="toolbar-group">
        <IconButton label="Previous page" icon={ChevronLeft} disabled={pageIndex <= 0} onClick={() => onMovePage(pageIndex - 1)} />
        <IconButton label="Next page" icon={ChevronRight} disabled={pageIndex >= doc.pageCount - 1} onClick={() => onMovePage(pageIndex + 1)} />
      </div>
      <div className="page-chip"><span>{pageIndex + 1}</span><span className="muted">/ {doc.pageCount}</span></div>
      <div className="toolbar-group">
        <IconButton label="Fit page" icon={Maximize2} active={zoomMode === 'fitPage'} onClick={() => onZoomModeChange('fitPage')} />
        <IconButton label="Fit width" icon={StretchHorizontal} active={zoomMode === 'fitWidth'} onClick={() => onZoomModeChange('fitWidth')} />
        <IconButton label="Zoom out" icon={ZoomOut} disabled={!canZoomOut} onClick={() => onManualZoomStep(-1)} />
        <span className="zoom-label">{renderScalePercent}%</span>
        <IconButton label="Zoom in" icon={ZoomIn} disabled={!canZoomIn} onClick={() => onManualZoomStep(1)} />
      </div>
    </div>
  );
}

function PdfPageView({
  pdf,
  pageIndex,
  zoomMode,
  manualZoom,
  onZoomMetricsChange,
}: {
  pdf: PDFDocumentProxy;
  pageIndex: number;
  zoomMode: ZoomMode;
  manualZoom: number;
  onZoomMetricsChange: (metrics: ZoomMetrics) => void;
}) {
  const stageRef = useRef<HTMLDivElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const size = useElementSize(stageRef);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    let renderTask: RenderTask | null = null;

    async function render() {
      const canvas = canvasRef.current;
      if (!canvas || size.width <= 0 || size.height <= 0) return;
      setError(null);
      try {
        const page = await pdf.getPage(pageIndex + 1);
        if (cancelled) return;
        const baseViewport = page.getViewport({ scale: 1 });
        const shellGap = 12;
        const availableWidth = Math.max(size.width - shellGap, 1);
        const availableHeight = Math.max(size.height - shellGap, 1);
        const fitPageScale = Math.min(availableWidth / baseViewport.width, availableHeight / baseViewport.height);
        const fitWidthScale = availableWidth / baseViewport.width;
        const minManualZoom = Math.min(MinManualZoom, fitPageScale);
        const maxManualZoom = Math.max(minManualZoom, fitWidthScale);
        const scale = zoomMode === 'fitPage'
          ? fitPageScale
          : zoomMode === 'fitWidth'
            ? fitWidthScale
            : clamp(manualZoom, minManualZoom, maxManualZoom);
        const boundedScale = Math.max(scale, 0.1);
        const viewport = page.getViewport({ scale: boundedScale });
        const ratio = window.devicePixelRatio || 1;
        const context = canvas.getContext('2d');
        if (!context) return;
        canvas.width = Math.floor(viewport.width * ratio);
        canvas.height = Math.floor(viewport.height * ratio);
        canvas.style.width = `${viewport.width}px`;
        canvas.style.height = `${viewport.height}px`;
        context.setTransform(1, 0, 0, 1, 0, 0);
        context.clearRect(0, 0, canvas.width, canvas.height);
        renderTask = page.render({
          canvasContext: context,
          viewport,
          transform: ratio === 1 ? undefined : [ratio, 0, 0, ratio, 0, 0],
        });
        await renderTask.promise;
        if (!cancelled) {
          onZoomMetricsChange({
            effectiveZoom: boundedScale,
            minManualZoom,
            maxManualZoom,
          });
        }
      } catch (err) {
        if (!cancelled && (err as Error).name !== 'RenderingCancelledException') {
          setError((err as Error).message || 'Could not render this page.');
        }
      }
    }

    void render();
    return () => {
      cancelled = true;
      renderTask?.cancel();
    };
  }, [manualZoom, onZoomMetricsChange, pageIndex, pdf, size.height, size.width, zoomMode]);

  return (
    <div ref={stageRef} className={`pdf-stage ${zoomMode === 'fitPage' ? 'fit-page' : 'scroll-fit'}`}>
      {error ? (
        <div className="missing-pdf">{error}</div>
      ) : (
        <div className="pdf-page-shell">
          <canvas ref={canvasRef} className="pdf-canvas" aria-label={`PDF page ${pageIndex + 1}`} />
        </div>
      )}
    </div>
  );
}

function ThumbnailPanel({
  doc,
  runtimePdf,
  pageIndex,
  bookmarks,
  bookmarkedOnly,
  onToggleBookmarkedOnly,
  onMovePage,
}: {
  doc: StoredDocument;
  runtimePdf: RuntimePdf | null;
  pageIndex: number;
  bookmarks: number[];
  bookmarkedOnly: boolean;
  onToggleBookmarkedOnly: () => void;
  onMovePage: (pageIndex: number) => void;
}) {
  const pages = useMemo(() => {
    const all = Array.from({ length: doc.pageCount }, (_, index) => index);
    return bookmarkedOnly ? all.filter((page) => bookmarks.includes(page)) : all;
  }, [bookmarkedOnly, bookmarks, doc.pageCount]);

  return (
    <aside className="thumbnail-panel">
      <div className="section-header">
        <div>
          <div className="section-title">Pages</div>
          <div className="muted">{pageIndex + 1} / {doc.pageCount}</div>
        </div>
        <IconButton label={bookmarkedOnly ? 'Show all pages' : 'Bookmarked only'} icon={Bookmark} active={bookmarkedOnly} onClick={onToggleBookmarkedOnly} />
      </div>
      <div className="thumb-grid">
        {pages.length === 0 ? (
          <div className="empty-state">Bookmarked pages will appear here.</div>
        ) : pages.map((page) => (
          <button
            key={page}
            className={`thumb-card ${page === pageIndex ? 'active' : ''}`}
            onClick={() => onMovePage(page)}
          >
            <div className="thumb-page">
              {runtimePdf ? (
                <ThumbnailCanvas pdf={runtimePdf.document} pageIndex={page} />
              ) : (
                <div className="thumb-placeholder"><FileText size={18} /></div>
              )}
              {bookmarks.includes(page) && <span className="thumb-bookmark"><Bookmark size={14} /></span>}
            </div>
            <span>{page + 1}</span>
          </button>
        ))}
      </div>
    </aside>
  );
}

function ThumbnailCanvas({ pdf, pageIndex }: { pdf: PDFDocumentProxy; pageIndex: number }) {
  const boxRef = useRef<HTMLDivElement | null>(null);
  const canvasRef = useRef<HTMLCanvasElement | null>(null);
  const size = useElementSize(boxRef);

  useEffect(() => {
    let cancelled = false;
    let renderTask: RenderTask | null = null;

    async function render() {
      const canvas = canvasRef.current;
      if (!canvas || size.width <= 0) return;
      const page = await pdf.getPage(pageIndex + 1);
      if (cancelled) return;
      const baseViewport = page.getViewport({ scale: 1 });
      const scale = Math.max(size.width / baseViewport.width, 0.08);
      const viewport = page.getViewport({ scale });
      const ratio = window.devicePixelRatio || 1;
      const context = canvas.getContext('2d');
      if (!context) return;
      canvas.width = Math.floor(viewport.width * ratio);
      canvas.height = Math.floor(viewport.height * ratio);
      canvas.style.width = `${viewport.width}px`;
      canvas.style.height = `${viewport.height}px`;
      context.clearRect(0, 0, canvas.width, canvas.height);
      renderTask = page.render({
        canvasContext: context,
        viewport,
        transform: ratio === 1 ? undefined : [ratio, 0, 0, ratio, 0, 0],
      });
      await renderTask.promise.catch(() => undefined);
    }

    void render();
    return () => {
      cancelled = true;
      renderTask?.cancel();
    };
  }, [pageIndex, pdf, size.width]);

  return (
    <div ref={boxRef} className="thumb-canvas-box">
      <canvas ref={canvasRef} />
    </div>
  );
}

function StudyRail({
  studyOpen,
  hasComments,
  bookmarked,
  onToggleStudy,
  onOpenNewComment,
  onToggleBookmark,
  onCopyPage,
}: {
  studyOpen: boolean;
  hasComments: boolean;
  bookmarked: boolean;
  onToggleStudy: () => void;
  onOpenNewComment: () => void;
  onToggleBookmark: () => void;
  onCopyPage: () => void;
}) {
  return (
    <aside className="study-rail">
      <IconButton label={studyOpen ? 'Hide Study Panel' : 'Study Panel'} icon={studyOpen ? PanelRightClose : PanelRightOpen} active={studyOpen} onClick={onToggleStudy} />
      <div className="rail-separator" />
      <IconButton label="Add comment" icon={MessageSquare} active={hasComments} onClick={onOpenNewComment} />
      <IconButton label={bookmarked ? 'Remove bookmark' : 'Bookmark page'} icon={Bookmark} active={bookmarked} onClick={onToggleBookmark} />
      <IconButton label="Copy text and page image" icon={Copy} onClick={onCopyPage} />
    </aside>
  );
}

function StudyPanel({
  pageIndex,
  comments,
  composerOpen,
  commentDraft,
  editingCommentId,
  onClose,
  onOpenNewComment,
  onCommentDraftChange,
  onSaveComment,
  onCancelComment,
  onEditComment,
  onDeleteComment,
}: {
  pageIndex: number;
  comments: StoredComment[];
  composerOpen: boolean;
  commentDraft: string;
  editingCommentId: string | null;
  onClose: () => void;
  onOpenNewComment: () => void;
  onCommentDraftChange: (value: string) => void;
  onSaveComment: () => void;
  onCancelComment: () => void;
  onEditComment: (comment: StoredComment) => void;
  onDeleteComment: (commentId: string) => void;
}) {
  return (
    <aside className="study-panel">
      <div className="panel-header">
        <div>
          <div className="section-title">Comments</div>
          <div className="muted">Page {pageIndex + 1}</div>
        </div>
        <div className="panel-actions">
          <IconButton label="Close Study Panel" icon={PanelRightClose} onClick={onClose} />
        </div>
      </div>
      <CommentsPanel
        pageIndex={pageIndex}
        comments={comments}
        composerOpen={composerOpen}
        commentDraft={commentDraft}
        editingCommentId={editingCommentId}
        onOpenNewComment={onOpenNewComment}
        onCommentDraftChange={onCommentDraftChange}
        onSaveComment={onSaveComment}
        onCancelComment={onCancelComment}
        onEditComment={onEditComment}
        onDeleteComment={onDeleteComment}
      />
    </aside>
  );
}

function CommentsPanel({
  pageIndex,
  comments,
  composerOpen,
  commentDraft,
  editingCommentId,
  onOpenNewComment,
  onCommentDraftChange,
  onSaveComment,
  onCancelComment,
  onEditComment,
  onDeleteComment,
}: {
  pageIndex: number;
  comments: StoredComment[];
  composerOpen: boolean;
  commentDraft: string;
  editingCommentId: string | null;
  onOpenNewComment: () => void;
  onCommentDraftChange: (value: string) => void;
  onSaveComment: () => void;
  onCancelComment: () => void;
  onEditComment: (comment: StoredComment) => void;
  onDeleteComment: (commentId: string) => void;
}) {
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);
  useEffect(() => {
    if (composerOpen) textareaRef.current?.focus();
  }, [composerOpen]);

  return (
    <div className="panel-body">
      <div className="mini-head">
        <div>
          <div className="section-title">Page {pageIndex + 1}</div>
          <div className="muted">{comments.length} comment{comments.length === 1 ? '' : 's'}</div>
        </div>
        {!composerOpen && (
          <IconButton label="Add comment" icon={Plus} onClick={onOpenNewComment} />
        )}
      </div>
      <div className="comment-list">
        {comments.length === 0 ? (
          <div className="empty-state">No comments yet.</div>
        ) : comments.map((comment) => (
          <CommentCard
            key={comment.id}
            comment={comment}
            selected={editingCommentId === comment.id}
            onEdit={onEditComment}
            onDelete={onDeleteComment}
          />
        ))}
      </div>
      {composerOpen && (
        <div className="composer">
          <textarea
            ref={textareaRef}
            value={commentDraft}
            onChange={(event) => onCommentDraftChange(event.target.value)}
            placeholder="Comment"
          />
          <div className="composer-actions">
            <button className="ghost-btn" onClick={onCancelComment}>Cancel</button>
            <button className="primary-btn compact" disabled={!commentDraft.trim()} onClick={onSaveComment}>
              {editingCommentId ? 'Save' : 'Add'}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function CommentCard({
  comment,
  selected,
  onEdit,
  onDelete,
}: {
  comment: StoredComment;
  selected: boolean;
  onEdit: (comment: StoredComment) => void;
  onDelete: (commentId: string) => void;
}) {
  return (
    <article className={`comment-card ${selected ? 'active' : ''}`}>
      <p>{comment.body}</p>
      <div className="comment-actions">
        <button className="ghost-btn" onClick={() => onEdit(comment)}>Edit</button>
        <button className="ghost-btn danger" onClick={() => onDelete(comment.id)}>Delete</button>
      </div>
    </article>
  );
}

function DriveImportDialog({
  pending,
  subjects,
  onConfirm,
  onClose,
}: {
  pending: PendingDriveImport;
  subjects: StoredSubject[];
  onConfirm: (choice: DriveImportChoice) => void;
  onClose: () => void;
}) {
  const [title, setTitle] = useState(pending.title);
  const [subjectId, setSubjectId] = useState(pending.subjectId ?? '');

  useEffect(() => {
    setTitle(pending.title);
    setSubjectId(pending.subjectId ?? '');
  }, [pending]);

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [onClose]);

  const submit = (event: FormEvent) => {
    event.preventDefault();
    const normalizedTitle = title.trim();
    if (!normalizedTitle) return;
    onConfirm({
      title: normalizedTitle,
      subjectId: subjectId || null,
    });
  };

  return (
    <div className="dialog-backdrop" role="presentation" onMouseDown={onClose}>
      <form
        className="app-dialog drive-import-dialog"
        onSubmit={submit}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <div>
          <h2>Add Drive PDF</h2>
          <p>Choose where this PDF belongs before adding it to Library.</p>
        </div>
        <div className="drive-import-summary">
          <div>
            <span>File</span>
            <strong>{pending.source.fileName}</strong>
          </div>
          <div>
            <span>Pages</span>
            <strong>{pending.source.pageCount}</strong>
          </div>
        </div>
        <label className="dialog-field">
          <span>PDF title</span>
          <input value={title} onChange={(event) => setTitle(event.target.value)} />
        </label>
        <label className="dialog-field">
          <span>Subject</span>
          <select value={subjectId} onChange={(event) => setSubjectId(event.target.value)}>
            <option value="">Unfiled</option>
            {subjects.map((subject) => (
              <option key={subject.id} value={subject.id}>{subject.name}</option>
            ))}
          </select>
        </label>
        <div className="dialog-actions">
          <button type="button" className="ghost-btn" onClick={onClose}>Cancel</button>
          <button type="submit" className="primary-btn compact" disabled={!title.trim()}>
            Add to Library
          </button>
        </div>
      </form>
    </div>
  );
}

function SettingsDialog({
  copySettings,
  syncEnabled,
  syncBusy,
  syncStatus,
  syncIssue,
  lastSyncedAt,
  onToggleCopySetting,
  onConnectSync,
  onSyncNow,
  onExportSyncData,
  onDisconnectSync,
  onResetRemoteSync,
  onExportBackup,
  onImportBackup,
  onClose,
}: {
  copySettings: CopyPacketOptions;
  syncEnabled: boolean;
  syncBusy: boolean;
  syncStatus: SyncStatusState;
  syncIssue: string | null;
  lastSyncedAt: number | null;
  onToggleCopySetting: (key: VisibleCopySettingKey, value: boolean) => void;
  onConnectSync: () => void;
  onSyncNow: () => void;
  onExportSyncData: () => void;
  onDisconnectSync: () => void;
  onResetRemoteSync: () => void;
  onExportBackup: () => void;
  onImportBackup: () => void;
  onClose: () => void;
}) {
  const [dataManagementOpen, setDataManagementOpen] = useState(false);

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [onClose]);

  return (
    <div className="dialog-backdrop" role="presentation" onMouseDown={onClose}>
      <section
        className="app-dialog settings-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="settings-title"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <div>
          <h2 id="settings-title">Settings</h2>
        </div>
        <div className="settings-section">
          <div className="settings-section-title">Copy</div>
          <div className="settings-list">
            <button
              type="button"
              className="setting-row setting-row-button"
              aria-pressed={copySettings.includePageImage}
              onClick={() => onToggleCopySetting('includePageImage', !copySettings.includePageImage)}
            >
              <span>
                <strong>Page image</strong>
                <span>PDF page screenshot</span>
              </span>
              <span className={`setting-toggle ${copySettings.includePageImage ? 'on' : ''}`} aria-hidden="true">
                <span />
              </span>
            </button>
            <button
              type="button"
              className="setting-row setting-row-button"
              aria-pressed={copySettings.includeComments}
              onClick={() => onToggleCopySetting('includeComments', !copySettings.includeComments)}
            >
              <span>
                <strong>Comments</strong>
                <span>Current page comments</span>
              </span>
              <span className={`setting-toggle ${copySettings.includeComments ? 'on' : ''}`} aria-hidden="true">
                <span />
              </span>
            </button>
          </div>
        </div>
        <div className="settings-section">
          <div className="settings-section-title">Drive sync</div>
          <div className={`sync-status-card ${syncStatus.kind}`}>
            <span>{syncStatus.label}</span>
            <strong>{syncStatusDetail(syncStatus, lastSyncedAt)}</strong>
            {syncIssue && (syncStatus.kind === 'failed' || syncStatus.kind === 'paused') && (
              <em>{syncIssue}</em>
            )}
          </div>
          <div className="settings-actions">
            <button
              type="button"
              className="ghost-btn"
              onClick={syncEnabled ? onSyncNow : onConnectSync}
              disabled={syncBusy}
            >
              <Cloud size={16} />
              {syncBusy ? 'Syncing...' : syncEnabled ? 'Sync now' : 'Connect'}
            </button>
            {syncEnabled && (
              <button type="button" className="ghost-btn" onClick={onDisconnectSync} disabled={syncBusy}>
                <Power size={16} />
                Disconnect
              </button>
            )}
          </div>
        </div>
        <div className="settings-section">
          <button
            type="button"
            className="settings-disclosure"
            aria-expanded={dataManagementOpen}
            onClick={() => setDataManagementOpen((open) => !open)}
          >
            <span>
              <strong>Data management</strong>
              <span>Backup and recovery tools</span>
            </span>
            <ChevronRight className={dataManagementOpen ? 'open' : ''} size={18} aria-hidden="true" />
          </button>
          {dataManagementOpen && (
            <div className="data-management-list">
              <div className="data-management-group">
                <div className="data-management-copy">
                  <strong>Sync data</strong>
                  <span>Drive-linked study data used for device sync.</span>
                </div>
                <div className="settings-actions">
                  <button type="button" className="ghost-btn" onClick={onExportSyncData}>
                    <Download size={16} />
                    Export sync data
                  </button>
                  {syncEnabled && (
                    <button type="button" className="ghost-btn danger" onClick={onResetRemoteSync} disabled={syncBusy}>
                      <RefreshCcw size={16} />
                      Reset remote
                    </button>
                  )}
                </div>
              </div>
              <div className="data-management-group">
                <div className="data-management-copy">
                  <strong>Local backup</strong>
                  <span>A manual snapshot of this browser's app data.</span>
                </div>
                <div className="settings-actions">
                  <button type="button" className="ghost-btn" onClick={onExportBackup}>
                    <Download size={16} />
                    Export local backup
                  </button>
                  <button type="button" className="ghost-btn" onClick={onImportBackup}>
                    <Upload size={16} />
                    Import local backup
                  </button>
                </div>
              </div>
            </div>
          )}
        </div>
        <div className="dialog-actions">
          <button type="button" className="ghost-btn" onClick={onClose}>Close</button>
        </div>
      </section>
    </div>
  );
}

function AppDialog({ dialog, onClose }: { dialog: DialogState; onClose: () => void }) {
  const [value, setValue] = useState(dialog.type === 'text' ? dialog.initialValue : '');
  const inputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => {
    if (dialog.type === 'text') {
      setValue(dialog.initialValue);
      requestAnimationFrame(() => inputRef.current?.focus());
    }
  }, [dialog]);

  const submit = (event?: FormEvent) => {
    event?.preventDefault();
    if (dialog.type === 'text') {
      const normalized = value.trim();
      if (!normalized) return;
      dialog.onConfirm(normalized);
      onClose();
      return;
    }
    dialog.onConfirm();
    onClose();
  };

  return (
    <div className="dialog-backdrop" role="presentation" onMouseDown={onClose}>
      <form className="app-dialog" onSubmit={submit} onMouseDown={(event) => event.stopPropagation()}>
        <div>
          <h2>{dialog.title}</h2>
          <p>{dialog.description}</p>
        </div>
        {dialog.type === 'text' && (
          <label className="dialog-field">
            <span>{dialog.label}</span>
            <input
              ref={inputRef}
              value={value}
              onChange={(event) => setValue(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Escape') onClose();
              }}
            />
          </label>
        )}
        <div className="dialog-actions">
          <button type="button" className="ghost-btn" onClick={onClose}>Cancel</button>
          <button
            type="submit"
            className={`primary-btn compact ${dialog.type === 'confirm' && dialog.danger ? 'danger-primary' : ''}`}
            disabled={dialog.type === 'text' && !value.trim()}
          >
            {dialog.confirmLabel}
          </button>
        </div>
      </form>
    </div>
  );
}

function IconButton({
  label,
  icon: Icon,
  active = false,
  disabled = false,
  onClick,
}: {
  label: string;
  icon: typeof ArrowLeft;
  active?: boolean;
  disabled?: boolean;
  onClick: () => void;
}) {
  return (
    <button
      className={`icon-btn ${active ? 'active' : ''}`}
      aria-label={label}
      title={label}
      disabled={disabled}
      onClick={onClick}
    >
      <Icon size={18} />
    </button>
  );
}

function useElementSize<T extends HTMLElement>(ref: React.RefObject<T>) {
  const [size, setSize] = useState({ width: 0, height: 0 });
  useEffect(() => {
    const element = ref.current;
    if (!element) return;
    const observer = new ResizeObserver((entries) => {
      const entry = entries[0];
      if (!entry) return;
      const { width, height } = entry.contentRect;
      setSize({ width, height });
    });
    observer.observe(element);
    return () => observer.disconnect();
  }, [ref]);
  return size;
}

export default App;
