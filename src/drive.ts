export type DriveConfig = {
  clientId: string;
  appId: string;
  apiKey: string;
};

export type DriveConfigStatus = {
  configured: boolean;
  missing: Array<keyof DriveConfig>;
};

export type DrivePdfFile = {
  id: string;
  name: string;
  mimeType: string;
  size: number | null;
  modifiedTime: string | null;
};

export type DriveAuthOptions = {
  hasGrantedFileAccess?: boolean;
  hasGrantedAppDataAccess?: boolean;
  forceAccountSelection?: boolean;
  forceConsent?: boolean;
  onTokenGranted?: () => void;
};

export type DriveAppDataJsonFile = {
  id: string;
  name: string;
  modifiedTime: string | null;
  size: number | null;
  data: unknown;
};

export type DriveAppDataSaveResult = {
  id: string;
  name: string;
  modifiedTime: string | null;
  size: number | null;
};

type GoogleTokenResponse = {
  access_token?: string;
  expires_in?: number;
  error?: string;
  error_description?: string;
};

type GoogleTokenClient = {
  requestAccessToken: (options?: { prompt?: '' | 'consent' | 'select_account' }) => void;
};

type GoogleOAuth = {
  initTokenClient: (config: {
    client_id: string;
    scope: string;
    callback: (response: GoogleTokenResponse) => void;
    error_callback?: (error: unknown) => void;
  }) => GoogleTokenClient;
};

type GooglePickerDocument = Record<string, unknown>;

type GooglePickerDocsView = {
  setMimeTypes: (mimeTypes: string) => GooglePickerDocsView;
  setIncludeFolders: (includeFolders: boolean) => GooglePickerDocsView;
  setParent: (parent: string) => GooglePickerDocsView;
  setSelectFolderEnabled: (enabled: boolean) => GooglePickerDocsView;
};

type GooglePickerBuilder = {
  setAppId: (appId: string) => GooglePickerBuilder;
  setDeveloperKey: (developerKey: string) => GooglePickerBuilder;
  setOAuthToken: (token: string) => GooglePickerBuilder;
  addView: (view: unknown) => GooglePickerBuilder;
  setCallback: (callback: (data: Record<string, unknown>) => void) => GooglePickerBuilder;
  setSize: (width: number, height: number) => GooglePickerBuilder;
  build: () => {
    setVisible: (visible: boolean) => void;
  };
};

type GooglePickerNamespace = {
  Action: {
    PICKED: string;
    CANCEL: string;
  };
  Document: {
    ID: string;
    NAME: string;
    MIME_TYPE: string;
  };
  DocsView: new (viewId?: string) => GooglePickerDocsView;
  PickerBuilder: new () => GooglePickerBuilder;
  Response: {
    ACTION: string;
    DOCUMENTS: string;
  };
  ViewId: {
    DOCS: string;
  };
};

type GoogleGlobal = {
  accounts?: {
    oauth2?: GoogleOAuth;
  };
  picker?: GooglePickerNamespace;
};

type GapiGlobal = {
  load: (api: string, callback: () => void) => void;
};

declare global {
  interface Window {
    google?: GoogleGlobal;
    gapi?: GapiGlobal;
  }
}

const DriveFileScope = 'https://www.googleapis.com/auth/drive.file';
export const DriveAppDataScope = 'https://www.googleapis.com/auth/drive.appdata';
export const DriveScopes = [DriveFileScope, DriveAppDataScope] as const;

const GoogleIdentityScriptUrl = 'https://accounts.google.com/gsi/client';
const GoogleApiScriptUrl = 'https://apis.google.com/js/api.js';
const DriveApiBaseUrl = 'https://www.googleapis.com/drive/v3/files';
const DriveUploadBaseUrl = 'https://www.googleapis.com/upload/drive/v3/files';
const PdfMimeType = 'application/pdf';
const JsonMimeType = 'application/json';
const TokenRefreshSkewMs = 60_000;
const PickerMaxWidth = 980;
const PickerMinWidth = 720;
const PickerMaxHeight = 560;
const PickerMinHeight = 460;

let identityScriptPromise: Promise<void> | null = null;
let pickerScriptPromise: Promise<void> | null = null;
let pickerLoadPromise: Promise<void> | null = null;
let accessTokenState: { token: string; expiresAt: number } | null = null;

export function getDriveConfig(): DriveConfig {
  return {
    clientId: readEnv('VITE_GOOGLE_CLIENT_ID'),
    appId: readEnv('VITE_GOOGLE_APP_ID'),
    apiKey: readEnv('VITE_GOOGLE_API_KEY'),
  };
}

export function getDriveConfigStatus(config = getDriveConfig()): DriveConfigStatus {
  const missing = (Object.keys(config) as Array<keyof DriveConfig>)
    .filter((key) => !config[key]);
  return {
    configured: missing.length === 0,
    missing,
  };
}

export function preloadDriveApis(): void {
  if (!getDriveConfigStatus().configured) return;
  void loadIdentityApi().catch(() => undefined);
  void loadPickerApi().catch(() => undefined);
}

export async function pickDrivePdf(authOptions: DriveAuthOptions = {}): Promise<DrivePdfFile | null> {
  const config = requireDriveConfig();
  const token = await getDriveAccessToken(authOptions, config);
  await loadPickerApi();
  const picker = window.google?.picker;
  if (!picker) throw new Error('Google Drive Picker could not be loaded.');
  const pickerSize = getPickerSize();

  return new Promise((resolve, reject) => {
    const view = new picker.DocsView(picker.ViewId.DOCS)
      .setMimeTypes(PdfMimeType)
      .setIncludeFolders(true)
      .setParent('root')
      .setSelectFolderEnabled(false);

    const pickerInstance = new picker.PickerBuilder()
      .setAppId(config.appId)
      .setDeveloperKey(config.apiKey)
      .setOAuthToken(token)
      .addView(view)
      .setSize(pickerSize.width, pickerSize.height)
      .setCallback((data) => {
        const action = data[picker.Response.ACTION];
        if (action === picker.Action.CANCEL) {
          resolve(null);
          return;
        }
        if (action !== picker.Action.PICKED) return;

        const documents = data[picker.Response.DOCUMENTS];
        const firstDocument = Array.isArray(documents) ? documents[0] : null;
        const picked = isRecord(firstDocument) ? readPickerDocument(firstDocument, picker) : null;
        if (!picked) {
          reject(new Error('Could not read the selected Drive PDF.'));
          return;
        }

        void fetchDrivePdfMetadata(picked.id, authOptions)
          .then((metadata) => resolve(metadata ?? picked))
          .catch(() => resolve(picked));
      })
      .build();

    pickerInstance.setVisible(true);
  });
}

export async function downloadDrivePdf(fileId: string, authOptions: DriveAuthOptions = {}): Promise<Blob> {
  const token = await getDriveAccessToken(authOptions);
  const url = `${DriveApiBaseUrl}/${encodeURIComponent(fileId)}?alt=media&supportsAllDrives=true`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: PdfMimeType,
    },
  });

  if (!response.ok) {
    if (response.status === 401 || response.status === 403) {
      accessTokenState = null;
    }
    throw new Error(`Could not download the Drive PDF (${response.status}).`);
  }

  return response.blob();
}

export async function fetchDrivePdfMetadata(fileId: string, authOptions: DriveAuthOptions = {}): Promise<DrivePdfFile | null> {
  const token = await getDriveAccessToken(authOptions);
  const params = new URLSearchParams({
    fields: 'id,name,mimeType,size,modifiedTime',
    supportsAllDrives: 'true',
  });
  const response = await fetch(`${DriveApiBaseUrl}/${encodeURIComponent(fileId)}?${params.toString()}`, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });
  if (!response.ok) return null;
  const raw = await response.json() as unknown;
  if (!isRecord(raw)) return null;
  const id = textOrNull(raw.id);
  const name = textOrNull(raw.name);
  if (!id || !name) return null;
  return {
    id,
    name,
    mimeType: textOrNull(raw.mimeType) ?? PdfMimeType,
    size: numberFromText(raw.size),
    modifiedTime: textOrNull(raw.modifiedTime),
  };
}

export async function readDriveAppDataJsonFile(
  fileName: string,
  authOptions: DriveAuthOptions = {},
): Promise<DriveAppDataJsonFile | null> {
  const token = await getDriveAccessToken(authOptions);
  const params = new URLSearchParams({
    spaces: 'appDataFolder',
    q: `name = '${escapeDriveQueryString(fileName)}' and trashed = false`,
    fields: 'files(id,name,modifiedTime,size)',
    pageSize: '1',
  });
  const listResponse = await fetch(`${DriveApiBaseUrl}?${params.toString()}`, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });
  if (!listResponse.ok) {
    handleAuthFailure(listResponse.status);
    throw new Error(`Could not read Drive sync file (${listResponse.status}).`);
  }

  const listRaw = await listResponse.json() as unknown;
  const files = isRecord(listRaw) && Array.isArray(listRaw.files) ? listRaw.files : [];
  const first = files.find(isRecord) ?? null;
  if (!first) return null;

  const id = textOrNull(first.id);
  const name = textOrNull(first.name) ?? fileName;
  if (!id) return null;

  const mediaResponse = await fetch(`${DriveApiBaseUrl}/${encodeURIComponent(id)}?alt=media`, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: JsonMimeType,
    },
  });
  if (!mediaResponse.ok) {
    handleAuthFailure(mediaResponse.status);
    throw new Error(`Could not download Drive sync file (${mediaResponse.status}).`);
  }

  return {
    id,
    name,
    modifiedTime: textOrNull(first.modifiedTime),
    size: numberFromText(first.size),
    data: await mediaResponse.json() as unknown,
  };
}

export async function saveDriveAppDataJsonFile(
  fileName: string,
  data: unknown,
  authOptions: DriveAuthOptions = {},
  existingFileId?: string | null,
): Promise<DriveAppDataSaveResult> {
  const token = await getDriveAccessToken(authOptions);
  const body = JSON.stringify(data, null, 2);
  const response = existingFileId
    ? await updateDriveAppDataJsonFile(existingFileId, body, token)
    : await createDriveAppDataJsonFile(fileName, body, token);
  if (!response.ok) {
    handleAuthFailure(response.status);
    throw new Error(`Could not save Drive sync file (${response.status}).`);
  }

  const raw = await response.json() as unknown;
  if (!isRecord(raw)) throw new Error('Drive did not return sync file metadata.');
  const id = textOrNull(raw.id);
  const name = textOrNull(raw.name) ?? fileName;
  if (!id) throw new Error('Drive did not return a sync file ID.');

  return {
    id,
    name,
    modifiedTime: textOrNull(raw.modifiedTime),
    size: numberFromText(raw.size),
  };
}

async function getDriveAccessToken(
  authOptions: DriveAuthOptions = {},
  config = requireDriveConfig(),
): Promise<string> {
  if (accessTokenState && accessTokenState.expiresAt - TokenRefreshSkewMs > Date.now()) {
    return accessTokenState.token;
  }

  await loadIdentityApi();
  const oauth = window.google?.accounts?.oauth2;
  if (!oauth) throw new Error('Google Identity Services could not be loaded.');

  return new Promise((resolve, reject) => {
    const tokenClient = oauth.initTokenClient({
      client_id: config.clientId,
      scope: DriveScopes.join(' '),
      callback: (response) => {
        if (response.error) {
          reject(new Error(response.error_description || response.error));
          return;
        }
        if (!response.access_token) {
          reject(new Error('Google did not return an access token.'));
          return;
        }
        accessTokenState = {
          token: response.access_token,
          expiresAt: Date.now() + Math.max((response.expires_in ?? 3600) - 30, 60) * 1000,
        };
        authOptions.onTokenGranted?.();
        resolve(response.access_token);
      },
      error_callback: () => {
        reject(new Error('Google sign-in was closed or blocked.'));
      },
    });

    const prompt = authOptions.forceAccountSelection
      ? 'select_account'
      : authOptions.forceConsent || !authOptions.hasGrantedFileAccess || !authOptions.hasGrantedAppDataAccess
        ? 'consent'
        : authOptions.hasGrantedFileAccess && authOptions.hasGrantedAppDataAccess
        ? ''
        : 'consent';
    tokenClient.requestAccessToken({ prompt });
  });
}

function createDriveAppDataJsonFile(fileName: string, body: string, token: string): Promise<Response> {
  const boundary = `slide-study-${Date.now().toString(36)}`;
  const metadata = {
    name: fileName,
    parents: ['appDataFolder'],
    mimeType: JsonMimeType,
  };
  const multipartBody = [
    `--${boundary}`,
    `Content-Type: ${JsonMimeType}; charset=UTF-8`,
    '',
    JSON.stringify(metadata),
    `--${boundary}`,
    `Content-Type: ${JsonMimeType}; charset=UTF-8`,
    '',
    body,
    `--${boundary}--`,
    '',
  ].join('\r\n');
  const params = new URLSearchParams({
    uploadType: 'multipart',
    fields: 'id,name,modifiedTime,size',
  });
  return fetch(`${DriveUploadBaseUrl}?${params.toString()}`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': `multipart/related; boundary=${boundary}`,
    },
    body: multipartBody,
  });
}

function updateDriveAppDataJsonFile(fileId: string, body: string, token: string): Promise<Response> {
  const params = new URLSearchParams({
    uploadType: 'media',
    fields: 'id,name,modifiedTime,size',
  });
  return fetch(`${DriveUploadBaseUrl}/${encodeURIComponent(fileId)}?${params.toString()}`, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': `${JsonMimeType}; charset=UTF-8`,
    },
    body,
  });
}

function handleAuthFailure(status: number): void {
  if (status === 401 || status === 403) {
    accessTokenState = null;
  }
}

function escapeDriveQueryString(value: string): string {
  return value.replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function requireDriveConfig(): DriveConfig {
  const config = getDriveConfig();
  const status = getDriveConfigStatus(config);
  if (!status.configured) {
    throw new Error(`Google Drive is missing config: ${status.missing.join(', ')}.`);
  }
  return config;
}

function loadIdentityApi(): Promise<void> {
  identityScriptPromise ??= loadScript(GoogleIdentityScriptUrl);
  return identityScriptPromise;
}

async function loadPickerApi(): Promise<void> {
  pickerScriptPromise ??= loadScript(GoogleApiScriptUrl);
  await pickerScriptPromise;
  if (window.google?.picker) return;
  pickerLoadPromise ??= new Promise((resolve, reject) => {
    if (!window.gapi) {
      reject(new Error('Google API client could not be loaded.'));
      return;
    }
    window.gapi.load('picker', resolve);
  });
  await pickerLoadPromise;
}

function loadScript(src: string): Promise<void> {
  const existing = document.querySelector<HTMLScriptElement>(`script[src="${src}"]`);
  if (existing?.dataset.loaded === 'true') return Promise.resolve();

  return new Promise((resolve, reject) => {
    const script = existing ?? document.createElement('script');
    script.src = src;
    script.async = true;
    script.defer = true;
    script.onload = () => {
      script.dataset.loaded = 'true';
      resolve();
    };
    script.onerror = () => reject(new Error(`Could not load ${src}.`));
    if (!existing) document.head.appendChild(script);
  });
}

function getPickerSize(): { width: number; height: number } {
  const viewportWidth = Math.max(document.documentElement.clientWidth, window.innerWidth || 0);
  const viewportHeight = Math.max(document.documentElement.clientHeight, window.innerHeight || 0);
  const maxWidth = Math.max(320, Math.min(PickerMaxWidth, viewportWidth - 180));
  const maxHeight = Math.max(360, Math.min(PickerMaxHeight, viewportHeight - 260));
  const minWidth = Math.min(PickerMinWidth, maxWidth);
  const minHeight = Math.min(PickerMinHeight, maxHeight);
  return {
    width: clampNumber(Math.floor(viewportWidth - 240), minWidth, maxWidth),
    height: clampNumber(Math.floor(viewportHeight - 300), minHeight, maxHeight),
  };
}

function readPickerDocument(document: GooglePickerDocument, picker: GooglePickerNamespace): DrivePdfFile | null {
  const id = textOrNull(document[picker.Document.ID]) ?? textOrNull(document.id);
  const name = textOrNull(document[picker.Document.NAME]) ?? textOrNull(document.name);
  if (!id || !name) return null;
  return {
    id,
    name,
    mimeType: textOrNull(document[picker.Document.MIME_TYPE]) ?? PdfMimeType,
    size: numberFromText(document.sizeBytes ?? document.size),
    modifiedTime: textOrNull(document.modifiedTime),
  };
}

function readEnv(key: keyof ImportMetaEnv): string {
  const value = import.meta.env[key];
  return typeof value === 'string' ? value.trim() : '';
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function textOrNull(value: unknown): string | null {
  return typeof value === 'string' && value.trim() ? value.trim() : null;
}

function numberFromText(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return Math.max(0, Math.floor(value));
  if (typeof value !== 'string' || !value.trim()) return null;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? Math.max(0, Math.floor(parsed)) : null;
}

function clampNumber(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}
