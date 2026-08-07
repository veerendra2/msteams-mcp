/**
 * New Teams (teams.cloud.microsoft) stores many API tokens under
 * `tmp.auth.v1.*.Token.*` with AES-256-CBC encryption instead of classic
 * MSAL AccessToken entries. Decrypt using GLOBAL ExportedEncryptionKey.
 */

import { createDecipheriv } from 'node:crypto';
import { readSessionState, getTeamsOrigin, type SessionState } from './session-store.js';

/**
 * State New Teams writes to UserAuthState when a silent (`prompt=none`) token
 * request is refused because there is no Entra ID session in the browser.
 * Every `.Token.` slot then holds a `dummy-token` placeholder.
 */
export const INTERACTION_REQUIRED = 'InteractionRequired';

/** Decrypted tmp.auth resource token with expiry. */
export interface TmpAuthTokenInfo {
  token: string;
  expiry: Date;
  /** Resource URL key fragment, e.g. HTTPS://CHATSVCAGG.TEAMS.MICROSOFT.COM */
  resource: string;
}

/**
 * Summary of the New Teams auth-service state.
 *
 * The Teams UI renders from its IndexedDB cache even when it holds no live
 * tokens, so "the app looks signed in" proves nothing. Token slots are the
 * only trustworthy signal.
 */
export interface TmpAuthState {
  /** UserAuthState.state, or null when the tmp.auth store is absent. */
  state: string | null;
  /** Token slots holding a usable token. */
  usableTokens: number;
  /** Token slots holding a placeholder because interaction is required. */
  placeholderTokens: number;
}

/** Nested item shape inside a tmp.auth Token localStorage entry. */
interface TmpAuthTokenItem {
  token?: string;
  expires?: number;
  encryptedToken?: string;
  iv?: string;
  error?: { errorCode?: string };
}

/** Wrapper stored in localStorage for tmp.auth entries. */
interface TmpAuthStorageValue {
  item?: TmpAuthTokenItem;
}

/** ExportedEncryptionKey localStorage payload. */
interface ExportedEncryptionKeyValue {
  item?: { exportedKey?: string };
}

/**
 * Reads the AES-256 key Teams exports for decrypting tmp.auth tokens.
 */
export function extractExportedEncryptionKey(
  localStorage: Array<{ name: string; value: string }>
): Buffer | null {
  for (const item of localStorage) {
    if (!item.name.includes('ExportedEncryptionKey')) continue;
    try {
      const parsed = JSON.parse(item.value) as ExportedEncryptionKeyValue;
      const keyB64 = parsed.item?.exportedKey;
      if (!keyB64) continue;
      const key = Buffer.from(keyB64, 'base64');
      if (key.length === 32) return key;
    } catch {
      continue;
    }
  }
  return null;
}

/**
 * Decrypts a New Teams tmp.auth encryptedToken (AES-256-CBC + base64 IV).
 */
export function decryptTmpAuthToken(
  encryptedTokenB64: string,
  ivB64: string,
  key: Buffer
): string | null {
  try {
    const encrypted = Buffer.from(encryptedTokenB64, 'base64');
    const iv = Buffer.from(ivB64, 'base64');
    if (key.length !== 32 || iv.length < 16 || encrypted.length === 0) {
      return null;
    }
    const decipher = createDecipheriv('aes-256-cbc', key, iv.subarray(0, 16));
    const plaintext = Buffer.concat([decipher.update(encrypted), decipher.final()]);
    const token = plaintext.toString('utf8');
    return token.startsWith('ey') ? token : null;
  } catch {
    return null;
  }
}

/**
 * Resolves a usable JWT from a tmp.auth Token item (plaintext or encrypted).
 */
export function resolveTmpAuthTokenItem(
  item: TmpAuthTokenItem,
  key: Buffer | null
): { token: string; expiry: Date } | null {
  const expiresSec = typeof item.expires === 'number' ? item.expires : 0;
  const expiry = new Date(expiresSec * 1000);
  if (expiresSec > 0 && expiry.getTime() <= Date.now()) {
    return null;
  }

  if (typeof item.token === 'string' && item.token.startsWith('ey')) {
    return { token: item.token, expiry };
  }

  // InteractionRequired placeholders are not usable.
  if (item.token === 'dummy-token' || item.error?.errorCode === INTERACTION_REQUIRED) {
    return null;
  }

  if (key && item.encryptedToken && item.iv) {
    const token = decryptTmpAuthToken(item.encryptedToken, item.iv, key);
    if (token) {
      // Prefer JWT exp when present; fall back to storage expires.
      try {
        const payload = JSON.parse(
          Buffer.from(token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString()
        ) as { exp?: number };
        if (typeof payload.exp === 'number') {
          const jwtExpiry = new Date(payload.exp * 1000);
          if (jwtExpiry.getTime() <= Date.now()) return null;
          return { token, expiry: jwtExpiry };
        }
      } catch {
        // use storage expiry
      }
      return expiresSec > 0 ? { token, expiry } : null;
    }
  }

  return null;
}

/**
 * Reads the New Teams auth-service state from saved session state.
 *
 * Use this to tell "signed in" from "needs an interactive sign-in": when
 * `usableTokens` is 0 and `state` is InteractionRequired, no amount of
 * refreshing will help — the browser has no Entra ID session left.
 */
export function readTmpAuthState(state?: SessionState): TmpAuthState {
  const sessionState = state ?? readSessionState();
  const localStorage = sessionState ? getTeamsOrigin(sessionState)?.localStorage : null;
  if (!localStorage) {
    return { state: null, usableTokens: 0, placeholderTokens: 0 };
  }

  const key = extractExportedEncryptionKey(localStorage);
  let authState: string | null = null;
  let usableTokens = 0;
  let placeholderTokens = 0;

  for (const entry of localStorage) {
    if (entry.name.includes('UserAuthState')) {
      try {
        const parsed = JSON.parse(entry.value) as { item?: { state?: string } };
        authState = parsed.item?.state ?? null;
      } catch {
        // Leave as null; a malformed entry is the same as an absent one.
      }
      continue;
    }

    if (!entry.name.includes('.Token.')) continue;

    try {
      const parsed = JSON.parse(entry.value) as TmpAuthStorageValue;
      if (!parsed.item) continue;
      if (resolveTmpAuthTokenItem(parsed.item, key)) {
        usableTokens++;
      } else {
        placeholderTokens++;
      }
    } catch {
      continue;
    }
  }

  return { state: authState, usableTokens, placeholderTokens };
}

/**
 * Finds and decrypts a tmp.auth Token for a resource URL substring
 * (case-insensitive), e.g. "chatsvcagg.teams.microsoft.com".
 */
export function extractTmpAuthResourceToken(
  resourceSubstring: string,
  state?: SessionState
): TmpAuthTokenInfo | null {
  const sessionState = state ?? readSessionState();
  if (!sessionState) return null;
  const localStorage = getTeamsOrigin(sessionState)?.localStorage;
  if (!localStorage) return null;

  const needle = resourceSubstring.toUpperCase();
  const key = extractExportedEncryptionKey(localStorage);
  let best: TmpAuthTokenInfo | null = null;

  for (const entry of localStorage) {
    if (!entry.name.includes('.Token.')) continue;
    const resourcePart = entry.name.split('.Token.')[1] ?? '';
    if (!resourcePart.toUpperCase().includes(needle)) continue;

    try {
      const parsed = JSON.parse(entry.value) as TmpAuthStorageValue;
      if (!parsed.item) continue;
      const resolved = resolveTmpAuthTokenItem(parsed.item, key);
      if (!resolved) continue;
      if (!best || resolved.expiry.getTime() > best.expiry.getTime()) {
        best = {
          token: resolved.token,
          expiry: resolved.expiry,
          resource: resourcePart,
        };
      }
    } catch {
      continue;
    }
  }

  return best;
}
