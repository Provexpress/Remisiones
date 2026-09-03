import {
  PublicClientApplication,
  type AccountInfo,
  type AuthenticationResult,
  type Configuration,
} from '@azure/msal-browser';
import type { FileMetadata, UserProfile } from '../types';

const clientId = import.meta.env.VITE_AZURE_CLIENT_ID || '';
const tenantId = import.meta.env.VITE_AZURE_TENANT_ID || '';
const sharepointUrl = import.meta.env.VITE_SHAREPOINT_FILE_URL || '';
const graphDriveId = import.meta.env.VITE_GRAPH_DRIVE_ID || '';
const graphItemId = import.meta.env.VITE_GRAPH_ITEM_ID || '';

const scopes = ['User.Read', 'Files.Read.All'];

const config: Configuration = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    redirectUri: `${window.location.origin}/redirect.html`,
    postLogoutRedirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: 'sessionStorage',
  },
};

const app = new PublicClientApplication(config);
let initialization: Promise<void> | null = null;

function initialize(): Promise<void> {
  if (!clientId || !tenantId) {
    return Promise.reject(new Error('Falta configurar VITE_AZURE_CLIENT_ID o VITE_AZURE_TENANT_ID.'));
  }
  if (!initialization) {
    initialization = app.initialize().then(async () => {
      const result = await app.handleRedirectPromise();
      if (result?.account) app.setActiveAccount(result.account);
      const account = app.getActiveAccount() || app.getAllAccounts()[0];
      if (account) app.setActiveAccount(account);
    });
  }
  return initialization;
}

async function accountFromLogin(): Promise<AccountInfo> {
  await initialize();
  const existing = app.getActiveAccount() || app.getAllAccounts()[0];
  if (existing) {
    app.setActiveAccount(existing);
    return existing;
  }
  const result = await app.loginPopup({ scopes, prompt: 'select_account' });
  if (!result.account) throw new Error('Microsoft 365 no devolvió una cuenta válida.');
  app.setActiveAccount(result.account);
  return result.account;
}

async function acquireToken(account: AccountInfo): Promise<AuthenticationResult> {
  try {
    return await app.acquireTokenSilent({ scopes, account });
  } catch {
    return app.acquireTokenPopup({ scopes, account });
  }
}

async function graphFetch(path: string, token: string, init: RequestInit = {}): Promise<Response> {
  const response = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...init,
    cache: 'no-store',
    headers: {
      Authorization: `Bearer ${token}`,
      ...(init.headers || {}),
    },
  });
  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    const message = body?.error?.message || `Microsoft Graph respondió ${response.status}.`;
    throw new Error(message);
  }
  return response;
}

function encodeSharingUrl(url: string): string {
  const bytes = new TextEncoder().encode(url);
  let binary = '';
  bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
  return `u!${btoa(binary).replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-')}`;
}

export async function getExistingProfile(): Promise<UserProfile | null> {
  await initialize();
  const account = app.getActiveAccount() || app.getAllAccounts()[0];
  if (!account) return null;
  app.setActiveAccount(account);
  return {
    name: account.name || account.username.split('@')[0],
    email: account.username,
  };
}

export async function signIn(): Promise<UserProfile> {
  const account = await accountFromLogin();
  const token = await acquireToken(account);
  const response = await graphFetch('/me?$select=displayName,mail,userPrincipalName', token.accessToken);
  const profile = await response.json();
  return {
    name: profile.displayName || account.name || 'Usuario',
    email: profile.mail || profile.userPrincipalName || account.username,
  };
}

export async function loadSharePointWorkbook(): Promise<{ buffer: ArrayBuffer; metadata: FileMetadata }> {
  const account = await accountFromLogin();
  const token = await acquireToken(account);
  let itemPath: string;
  let metadata: FileMetadata;
  let download: Response;

  if (graphDriveId && graphItemId) {
    itemPath = `/drives/${encodeURIComponent(graphDriveId)}/items/${encodeURIComponent(graphItemId)}`;
    const response = await graphFetch(`${itemPath}?$select=id,name,size,lastModifiedDateTime,webUrl`, token.accessToken);
    metadata = await response.json();
    download = await graphFetch(`${itemPath}/content`, token.accessToken);
  } else {
    if (!sharepointUrl) throw new Error('Falta configurar VITE_SHAREPOINT_FILE_URL.');
    const shareId = encodeSharingUrl(sharepointUrl);
    itemPath = `/shares/${shareId}/driveItem`;
    download = await graphFetch(`${itemPath}/content`, token.accessToken);
    metadata = {
      name: 'Remisiones.xlsx',
      webUrl: sharepointUrl,
      lastModifiedDateTime: download.headers.get('last-modified') || undefined,
    };
  }

  return { buffer: await download.arrayBuffer(), metadata };
}

export async function signOut(): Promise<void> {
  await initialize();
  const account = app.getActiveAccount() || app.getAllAccounts()[0];
  await app.logoutPopup({ account, postLogoutRedirectUri: window.location.origin });
}

export const microsoftConfig = { clientId, tenantId, sharepointUrl, scopes };
