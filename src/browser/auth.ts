/**
 * Authentication handling for Microsoft Teams.
 * Manages login detection and manual authentication flows.
 */

import type { Page, BrowserContext } from 'playwright';
import { saveSessionState } from './context.js';
import {
  OVERLAY_STEP_PAUSE_MS,
  OVERLAY_COMPLETE_PAUSE_MS,
} from '../constants.js';
import { extractSubstrateToken, extractMessageAuth } from '../auth/token-extractor.js';
import { INTERACTION_REQUIRED } from '../auth/tmp-auth-tokens.js';
import { ensureMessagingCookies } from '../auth/token-refresh-http.js';
import * as logger from '../utils/logger.js';

/**
 * Default Teams URL for initial login (New Teams).
 *
 * Commercial tenants use teams.cloud.microsoft. GCC/GCC-High/DoD tenants are
 * redirected by Microsoft's login flow to the appropriate host
 * (teams.microsoft.us, etc.). API base URLs still come from DISCOVER-REGION-GTM.
 */
const TEAMS_URL = 'https://teams.cloud.microsoft';

/** Legacy commercial host — kept as a secondary clear/fallback target. */
const TEAMS_LEGACY_URL = 'https://teams.microsoft.com';

// ─────────────────────────────────────────────────────────────────────────────
// Progress Overlay UI
// ─────────────────────────────────────────────────────────────────────────────

const PROGRESS_OVERLAY_ID = 'mcp-login-progress-overlay';

/** Phases for the login progress overlay. */
type OverlayPhase =
  | 'signed-in'
  | 'acquiring-tokens'
  | 'saving'
  | 'complete'
  | 'needs-reauth'
  | 'error';

/** Content for each overlay phase. */
const OVERLAY_CONTENT: Record<OverlayPhase, { message: string; detail: string }> = {
  'signed-in': {
    message: "You're signed in!",
    detail: 'Setting up your connection to Teams...',
  },
  'acquiring-tokens': {
    message: 'Finishing sign-in...',
    detail: 'Waiting for Teams API tokens. Keep this window open — it can take up to a minute.',
  },
  'saving': {
    message: 'Saving your session...',
    detail: "So you won't need to log in again.",
  },
  'complete': {
    message: 'All done!',
    detail: 'This window will close automatically.',
  },
  'needs-reauth': {
    message: 'One more sign-in step',
    detail: 'Please sign in again so we can finish connecting to Teams.',
  },
  'error': {
    message: 'Something went wrong',
    detail: 'Please try again or check the console for details.',
  },
};

/**
 * Shows a progress overlay for a specific phase.
 * Handles injection, content, and optional pause.
 * Failures are silently ignored - the overlay is purely cosmetic.
 */
async function showLoginProgress(
  page: Page,
  phase: OverlayPhase,
  options: { pause?: boolean } = {}
): Promise<void> {
  const content = OVERLAY_CONTENT[phase];
  const isComplete = phase === 'complete';
  const isError = phase === 'error';
  const isNeedsReauth = phase === 'needs-reauth';

  try {
    await page.evaluate(({ id, message, detail, complete, error, needsReauth }) => {
      // Remove existing overlay if present
      const existing = document.getElementById(id);
      if (existing) existing.remove();

      // Create overlay container
      const overlay = document.createElement('div');
      overlay.id = id;
      Object.assign(overlay.style, {
        position: 'fixed',
        top: '0',
        left: '0',
        right: '0',
        bottom: '0',
        background: 'rgba(0, 0, 0, 0.7)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: '999999',
        fontFamily: "'Segoe UI', system-ui, sans-serif",
      });

      // Create modal card
      const modal = document.createElement('div');
      Object.assign(modal.style, {
        background: 'white',
        borderRadius: '12px',
        padding: '40px 48px',
        maxWidth: '420px',
        textAlign: 'center',
        boxShadow: '0 8px 32px rgba(0, 0, 0, 0.3)',
      });

      // Create icon
      const icon = document.createElement('div');
      const iconBg = error ? '#c42b1c' : complete ? '#107c10' : needsReauth ? '#f7630c' : '#5b5fc7';
      Object.assign(icon.style, {
        width: '64px',
        height: '64px',
        borderRadius: '50%',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        fontSize: '32px',
        background: iconBg,
        color: 'white',
        margin: '0 auto 24px',
      });
      icon.textContent = error ? '✕' : complete ? '✓' : needsReauth ? '!' : '⋯';

      // Create title
      const title = document.createElement('h2');
      Object.assign(title.style, {
        margin: '0 0 12px',
        fontSize: '20px',
        fontWeight: '600',
        color: '#242424',
      });
      title.textContent = message;

      // Create detail text
      const detailEl = document.createElement('p');
      Object.assign(detailEl.style, {
        margin: '0',
        fontSize: '14px',
        color: '#616161',
        lineHeight: '1.5',
      });
      detailEl.textContent = detail;

      // Assemble and append
      modal.appendChild(icon);
      modal.appendChild(title);
      modal.appendChild(detailEl);
      overlay.appendChild(modal);
      document.body.appendChild(overlay);
    }, {
      id: PROGRESS_OVERLAY_ID,
      message: content.message,
      detail: content.detail,
      complete: isComplete,
      error: isError,
      needsReauth: isNeedsReauth,
    });

    // Pause if requested (for steps that need user to see the message)
    if (options.pause) {
      const pauseMs = (isComplete || isError || isNeedsReauth)
        ? OVERLAY_COMPLETE_PAUSE_MS
        : OVERLAY_STEP_PAUSE_MS;
      await page.waitForTimeout(pauseMs);
    }
  } catch {
    // Overlay is cosmetic - don't fail login if it can't be shown
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Authentication Detection
// ─────────────────────────────────────────────────────────────────────────────

// URLs that indicate we're in a login flow
const LOGIN_URL_PATTERNS = [
  'login.microsoftonline.com',
  'login.live.com',
  'login.microsoft.com',
];

export interface AuthStatus {
  isAuthenticated: boolean;
  isOnLoginPage: boolean;
  currentUrl: string;
  /**
   * Teams reported that silent token acquisition failed and only an
   * interactive sign-in can fix it. Distinguishes "needs credentials" from
   * "still booting".
   */
  interactionRequired?: boolean;
  /** The page holds no auth store at all (cleared or first-run profile). */
  authStoreEmpty?: boolean;
}

/**
 * Checks if the current page URL indicates a login flow.
 */
function isLoginUrl(url: string): boolean {
  return LOGIN_URL_PATTERNS.some(pattern => url.includes(pattern));
}

/** How long to wait for the SPA to acquire tokens before giving up (ms). */
const AUTH_EVIDENCE_TIMEOUT_MS = 25000;

/** Interval between auth evidence checks (ms). */
const AUTH_EVIDENCE_POLL_MS = 500;

/**
 * Minimum time on the page before a persisted InteractionRequired state is
 * believed. The value survives in localStorage from the previous run, so
 * acting on it immediately would abort a login that was about to succeed.
 */
const INTERACTION_REQUIRED_DWELL_MS = 9000;

/** URL patterns that indicate we're on a Teams page (not redirected elsewhere). */
const TEAMS_URL_PATTERNS = [
  'teams.cloud.microsoft',   // New Teams (default)
  'teams.microsoft.com',     // Legacy commercial
  'teams.microsoft.us',      // GCC-High
  'dod.teams.microsoft.us',  // DoD
  'gov.teams.microsoft.us',  // GCC
];

/**
 * Checks if a URL is a Teams domain.
 */
function isTeamsUrl(url: string): boolean {
  return TEAMS_URL_PATTERNS.some(pattern => url.includes(pattern));
}

/** Auth evidence read straight out of the Teams page's own token stores. */
export interface PageAuthEvidence {
  /** New Teams UserAuthState.state, null when the tmp.auth store is absent. */
  authServiceState: string | null;
  /** Live tokens: New Teams tmp.auth slots plus classic MSAL access tokens. */
  usableTokens: number;
  /** New Teams token slots holding an interaction-required placeholder. */
  placeholderTokens: number;
  /** Minutes left on the best Substrate search token, -1 when there is none. */
  substrateMinutesRemaining: number;
}

/**
 * Reads every auth signal the page holds in one pass.
 *
 * Deliberately ignores the DOM. Teams renders its full chat list from an
 * IndexedDB cache while holding nothing but `dummy-token` placeholders, so
 * screen scraping reports success on a session that cannot call a single API.
 *
 * Returns null when the page could not be read at all — during boot the SPA
 * navigates several times and `evaluate` throws "execution context destroyed".
 * That must not be confused with "read fine, found nothing", which is a real
 * and very different conclusion.
 */
export async function readPageAuthEvidence(page: Page): Promise<PageAuthEvidence | null> {
  try {
    if (page.isClosed()) return null;

    // NOTE: declare no named functions inside this callback. Under tsx/esbuild
    // a named inner function is rewritten to `__name(fn, "...")`, and that
    // helper does not exist in the page — every call then dies with
    // "ReferenceError: __name is not defined". Keep the logic inline.
    return await page.evaluate(() => {
      const now = Date.now();
      let authServiceState: string | null = null;
      let usableTokens = 0;
      let placeholderTokens = 0;
      let substrateMinutesRemaining = -1;

      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (!key) continue;
        const raw = localStorage.getItem(key);
        if (!raw?.startsWith('{')) continue;

        let entry: Record<string, unknown>;
        try {
          entry = JSON.parse(raw) as Record<string, unknown>;
        } catch {
          continue;
        }

        // New Teams: auth-service state.
        if (key.includes('UserAuthState')) {
          const item = entry.item as { state?: string } | undefined;
          authServiceState = item?.state ?? null;
          continue;
        }

        // New Teams: per-resource token slots.
        if (key.includes('.Token.')) {
          const item = entry.item as
            | { token?: string; encryptedToken?: string; expires?: number; error?: { errorCode?: string } }
            | undefined;
          if (!item) continue;

          const expired = typeof item.expires === 'number' && item.expires > 0 && item.expires * 1000 <= now;
          const placeholder = item.token === 'dummy-token' || item.error?.errorCode !== undefined;
          const hasToken = item.token?.startsWith('ey') === true || typeof item.encryptedToken === 'string';

          if (!placeholder && hasToken && !expired) usableTokens++;
          else placeholderTokens++;
          continue;
        }

        // Classic MSAL cache (legacy teams.microsoft.com sessions).
        const credentialType = entry.credentialType as string | undefined;
        const secret = entry.secret as string | undefined;
        if (!secret?.startsWith('ey')) continue;

        let secretExpiryMs: number | null = null;
        try {
          const b64 = secret.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
          const payload = JSON.parse(atob(b64)) as { exp?: number };
          if (typeof payload.exp === 'number') secretExpiryMs = payload.exp * 1000;
        } catch {
          // Not a decodable JWT; leave the expiry unknown.
        }

        if (secretExpiryMs === null || secretExpiryMs <= now) continue;

        if (credentialType === 'AccessToken') usableTokens++;

        const target = entry.target as string | undefined;
        if (target?.includes('SubstrateSearch')) {
          const mins = Math.round((secretExpiryMs - now) / 60000);
          if (mins > substrateMinutesRemaining) substrateMinutesRemaining = mins;
        }
      }

      return {
        authServiceState,
        usableTokens,
        placeholderTokens,
        substrateMinutesRemaining,
      };
    });
  } catch (error) {
    logger.warn('auth', `Auth store read failed: ${error instanceof Error ? error.message.split('\n')[0] : String(error)}`);
    return null;
  }
}

/** True when the page holds at least one token an API call could actually use. */
function hasUsableAuth(evidence: PageAuthEvidence | null): boolean {
  return evidence !== null && (evidence.usableTokens > 0 || evidence.substrateMinutesRemaining > 0);
}

/**
 * Gets the current authentication status.
 *
 * Accepts every Teams host (including teams.cloud.microsoft) and decides based
 * on live tokens rather than page chrome or redirect behaviour.
 */
export async function getAuthStatus(page: Page): Promise<AuthStatus> {
  const currentUrl = page.url();

  // If on login page, definitely not authenticated
  if (isLoginUrl(currentUrl)) {
    return {
      isAuthenticated: false,
      isOnLoginPage: true,
      currentUrl,
    };
  }

  if (isTeamsUrl(currentUrl)) {
    const evidence = await readPageAuthEvidence(page);
    const authenticated = hasUsableAuth(evidence);

    return {
      isAuthenticated: authenticated,
      isOnLoginPage: false,
      currentUrl,
      interactionRequired: !authenticated && evidence?.authServiceState === INTERACTION_REQUIRED,
      // Only meaningful when the page actually read back: a cleared or
      // first-run profile holds no credential to refresh.
      authStoreEmpty:
        evidence !== null &&
        evidence.authServiceState === null &&
        evidence.usableTokens === 0 &&
        evidence.placeholderTokens === 0,
    };
  }

  // Unknown state
  return {
    isAuthenticated: false,
    isOnLoginPage: false,
    currentUrl,
  };
}

/**
 * Navigates to Teams and waits until the page proves whether it has tokens.
 *
 * "No redirect to a login page" must never be read as success. When its Entra
 * ID session is gone, New Teams keeps the top frame on teams.cloud.microsoft
 * and fails `prompt=none` requests inside a hidden iframe, so a redirect-based
 * check reports authenticated 100% of the time — including on a session that
 * holds nothing but placeholders.
 */
export async function navigateToTeams(
  page: Page,
  maxWaitMs: number = AUTH_EVIDENCE_TIMEOUT_MS,
): Promise<AuthStatus> {
  const startTime = Date.now();

  await page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded' });

  let status = await getAuthStatus(page);

  while (!status.isAuthenticated && !status.isOnLoginPage) {
    const elapsed = Date.now() - startTime;

    // A persisted InteractionRequired only counts once the SPA has had time to
    // retry; the stored value is left over from the previous run.
    if (status.interactionRequired && elapsed >= INTERACTION_REQUIRED_DWELL_MS) break;
    if (elapsed >= maxWaitMs) break;
    if (page.isClosed()) break;

    try {
      await page.waitForTimeout(AUTH_EVIDENCE_POLL_MS);
    } catch {
      break;
    }
    status = await getAuthStatus(page);
  }

  return status;
}

/**
 * Waits for the user to complete manual authentication.
 * Returns when authenticated or throws after timeout.
 *
 * @param page - The page to monitor
 * @param context - Browser context for saving session
 * @param timeoutMs - Maximum time to wait (default: 5 minutes)
 * @param onProgress - Callback for progress updates
 * @param showOverlay - Whether to show progress overlay (default: true for visible browsers)
 */
export async function waitForManualLogin(
  page: Page,
  context: BrowserContext,
  timeoutMs: number = 5 * 60 * 1000,
  onProgress?: (message: string) => void,
  showOverlay: boolean = true
): Promise<void> {
  const startTime = Date.now();
  const log = onProgress ?? console.log;
  let announcedSignIn = false;
  let lastVerifyAt = 0;
  let verifyIntervalMs = SESSION_VERIFY_INTERVAL_MS;

  log('Waiting for manual login...');

  while (Date.now() - startTime < timeoutMs) {
    if (page.isClosed()) {
      throw new Error('Authentication failed: the browser window was closed before sign-in completed.');
    }

    const status = await getAuthStatus(page);

    // Verifying means writing ~600KB of encrypted session state and possibly
    // calling authsvc, so it is rate limited rather than run on every poll.
    if (status.isAuthenticated && Date.now() - lastVerifyAt >= verifyIntervalMs) {
      lastVerifyAt = Date.now();

      if (!announcedSignIn) {
        announcedSignIn = true;
        log('Tokens acquired — finishing sign-in.');
        if (showOverlay) {
          await showLoginProgress(page, 'signed-in', { pause: true });
          await showLoginProgress(page, 'acquiring-tokens');
        }
      }

      await saveSessionState(context);

      if (await ensureUsableSession(log)) {
        log('Session state saved.');
        if (showOverlay) {
          await showLoginProgress(page, 'saving', { pause: true });
          await showLoginProgress(page, 'complete', { pause: true });
        }
        return;
      }

      // Cookies can lag the tokens by a few seconds, but if they never arrive
      // back off rather than retrying authsvc every few seconds for minutes.
      verifyIntervalMs = Math.min(verifyIntervalMs * 2, MAX_SESSION_VERIFY_INTERVAL_MS);
    }

    await page.waitForTimeout(LOGIN_POLL_INTERVAL_MS);
  }

  if (showOverlay) {
    await showLoginProgress(page, 'error', { pause: true });
  }

  throw new Error('Authentication timeout: user did not complete login within the allowed time');
}

/** How often to re-check auth while waiting for the user to sign in (ms). */
const LOGIN_POLL_INTERVAL_MS = 1500;

/** Minimum gap between full session save + verify passes (ms). */
const SESSION_VERIFY_INTERVAL_MS = 3000;

/** Ceiling for the verify interval once it starts backing off (ms). */
const MAX_SESSION_VERIFY_INTERVAL_MS = 30000;

/** True when the saved session already has credentials a tool could use. */
function sessionIsUsable(): boolean {
  // Substrate-only sessions still power search and people tools.
  return extractMessageAuth() !== null || extractSubstrateToken() !== null;
}

/**
 * Checks the saved session and, if messaging auth is missing, mints the
 * messaging cookies from the Skype Spaces token before giving up.
 *
 * Performs at most one authsvc round trip per call.
 */
async function ensureUsableSession(log: (message: string) => void): Promise<boolean> {
  if (extractMessageAuth()) return true;

  const derived = await ensureMessagingCookies();
  if (derived.ok && derived.value) {
    log('Derived messaging cookies from the Skype Spaces token.');
  }

  return sessionIsUsable();
}

/**
 * Polls for the saved session to become usable after tokens have landed.
 *
 * Teams writes its messaging cookies a moment after the tokens appear, so a
 * single check straight after acquisition can race and fail.
 */
async function waitForUsableSession(
  page: Page,
  context: BrowserContext,
  log: (message: string) => void,
  timeoutMs: number,
): Promise<boolean> {
  const deadline = Date.now() + timeoutMs;

  while (Date.now() < deadline) {
    if (await ensureUsableSession(log)) return true;
    if (page.isClosed()) return false;

    try {
      await page.waitForTimeout(SESSION_VERIFY_INTERVAL_MS);
    } catch {
      return false;
    }
    await saveSessionState(context);
  }

  return sessionIsUsable();
}

// ─────────────────────────────────────────────────────────────────────────────
// Token Refresh Polling (for headless browser fallback)
// ─────────────────────────────────────────────────────────────────────────────

/** Absolute max wait for tokens to appear during a silent refresh (ms). */
const TOKEN_REFRESH_MAX_WAIT_MS = 45000;

/**
 * Shorter cap for headless refresh, which runs inside an MCP tool call. Left
 * unbounded it can stack with navigateToTeams and blow past client timeouts.
 */
const HEADLESS_TOKEN_WAIT_MS = 15000;

/** How long to let the messaging cookies catch up once tokens exist (ms). */
const SESSION_SETTLE_WAIT_MS = 9000;

/** Interval for checking if tokens have been refreshed (ms). */
const TOKEN_REFRESH_POLL_INTERVAL_MS = 1000;

/** How often to reload the page during token wait to kick silent acquisition. */
const TOKEN_REFRESH_RELOAD_INTERVAL_MS = 12000;

/** Options for waitForTokenRefresh. */
interface TokenRefreshWaitOptions {
  /** Absolute max wait for usable tokens. */
  maxWaitMs?: number;
}

/**
 * Clears the page's auth stores so the next navigation must perform a real sign-in.
 *
 * Covers both storage layouts: classic MSAL cache keys and the New Teams
 * `tmp.auth.*` store (tokens, encryption key, and the cached auth state). Only
 * clearing the MSAL keys is a no-op on New Teams, which is why a "fresh" login
 * used to boot straight back into the same dead session.
 */
async function clearAuthLocalStorage(page: Page): Promise<number> {
  try {
    return await page.evaluate(() => {
      const keys: string[] = [];
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (!key) continue;
        if (
          key.startsWith('msal.') ||
          key.startsWith('tmp.auth.') ||
          /-accesstoken-/i.test(key) ||
          /-refreshtoken-/i.test(key) ||
          /-idtoken-/i.test(key)
        ) {
          keys.push(key);
        }
      }
      for (const key of keys) {
        localStorage.removeItem(key);
      }
      return keys.length;
    });
  } catch {
    return 0;
  }
}

/**
 * Clears cookies and token caches so the next navigation must perform a real sign-in.
 * Prefers the New Teams host; also clears the legacy commercial host if needed.
 */
export async function clearAuthStateForFreshLogin(
  page: Page,
  context: BrowserContext,
  onProgress?: (message: string) => void,
): Promise<void> {
  const log = onProgress ?? ((msg: string) => logger.debug('auth', msg));

  await context.clearCookies();

  for (const url of [TEAMS_URL, TEAMS_LEGACY_URL]) {
    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 12000 });
      const removed = await clearAuthLocalStorage(page);
      if (removed > 0) {
        log(`Cleared ${removed} cached auth entries on ${url}`);
      }
      // Cleared on primary host — skip the second visit to start login faster
      if (url === TEAMS_URL) {
        break;
      }
    } catch (error) {
      log(`Could not clear cached auth on ${url}: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
}

/**
 * Waits for the page to acquire usable tokens, reloading periodically to prod
 * silent acquisition along.
 *
 * Gating this on a classic MSAL RefreshToken (as it used to) makes it a no-op
 * against New Teams, which keeps no refresh token in web storage at all and
 * relies purely on the Entra ID session cookie.
 *
 * @returns true if usable tokens appeared, false on timeout or a definitive
 *          InteractionRequired
 */
async function waitForTokenRefresh(
  page: Page,
  context: BrowserContext,
  onProgress?: (message: string) => void,
  options: TokenRefreshWaitOptions = {},
): Promise<boolean> {
  const log = onProgress ?? ((msg: string) => logger.debug('auth', msg));
  const maxWaitMs = options.maxWaitMs ?? TOKEN_REFRESH_MAX_WAIT_MS;

  log('Waiting for Teams to acquire tokens...');
  const startTime = Date.now();
  let lastReloadAt = startTime;

  while (Date.now() - startTime < maxWaitMs) {
    if (page.isClosed()) {
      log('Browser closed during token wait.');
      return false;
    }

    // Never touch a page that is in the middle of a sign-in. Reloading here
    // navigates the user off the Microsoft login form and restarts the flow.
    if (isLoginUrl(page.url())) {
      log('Sign-in page reached — handing over to interactive login.');
      return false;
    }

    const evidence = await readPageAuthEvidence(page);
    if (evidence && hasUsableAuth(evidence)) {
      await saveSessionState(context);
      log(
        `Tokens acquired (${evidence.usableTokens} live token(s)` +
        (evidence.substrateMinutesRemaining > 0
          ? `, Substrate valid ${evidence.substrateMinutesRemaining} mins`
          : '') +
        ').'
      );
      return true;
    }

    const elapsed = Date.now() - startTime;

    log(
      evidence === null
        ? `Auth store unreadable (page busy) at ${page.url()}`
        : `Auth store: state=${evidence.authServiceState ?? 'none'} live=${evidence.usableTokens} ` +
          `placeholders=${evidence.placeholderTokens} substrateMins=${evidence.substrateMinutesRemaining} url=${page.url()}`
    );

    // A readable but empty auth store means a cleared profile with nothing to
    // refresh. Requires a successful read — an unreadable page is just early.
    if (
      evidence !== null &&
      evidence.authServiceState === null &&
      evidence.placeholderTokens === 0 &&
      elapsed >= INTERACTION_REQUIRED_DWELL_MS
    ) {
      log('No auth store in the profile — interactive sign-in needed.');
      return false;
    }

    // Teams has told us silent acquisition is refused; more waiting is pointless.
    if (evidence?.authServiceState === INTERACTION_REQUIRED && elapsed >= INTERACTION_REQUIRED_DWELL_MS) {
      log('Teams reports InteractionRequired — no Entra ID session, interactive sign-in needed.');
      return false;
    }

    if (Date.now() - lastReloadAt >= TOKEN_REFRESH_RELOAD_INTERVAL_MS) {
      lastReloadAt = Date.now();
      try {
        // Only ever reload a Teams page. Anything else may be part of an
        // in-flight auth redirect chain that a navigation would break.
        if (isTeamsUrl(page.url())) {
          log('Reloading Teams to trigger token acquisition...');
          await page.reload({ waitUntil: 'domcontentloaded', timeout: 20000 });
        } else {
          log('Not on Teams — leaving the current page alone.');
          return false;
        }
      } catch (error) {
        log(`Reload during token wait failed: ${error instanceof Error ? error.message : String(error)}`);
        if (page.isClosed()) return false;
      }
    }

    try {
      await page.waitForTimeout(TOKEN_REFRESH_POLL_INTERVAL_MS);
    } catch {
      return false;
    }
  }

  log('Token acquisition timed out.');
  return false;
}

/**
 * Performs a full authentication flow:
 * 1. Navigate to Teams
 * 2. Check if already authenticated
 * 3. If not, wait for manual login (or throw if headless)
 *
 * @param page - The page to use
 * @param context - Browser context for session management
 * @param onProgress - Callback for progress updates
 * @param showOverlay - Whether to show progress overlay (default: true for visible browsers)
 * @param headless - If true, throw immediately if user interaction is required (default: false)
 */
export async function ensureAuthenticated(
  page: Page,
  context: BrowserContext,
  onProgress?: (message: string) => void,
  showOverlay: boolean = true,
  headless: boolean = false
): Promise<void> {
  const log = onProgress ?? console.log;

  // Headless callers run inside an MCP request, so every wait has to stay well
  // under the client's timeout. Visible logins can afford to be patient.
  const tokenWaitMs = headless ? HEADLESS_TOKEN_WAIT_MS : TOKEN_REFRESH_MAX_WAIT_MS;

  log('Navigating to Teams...');
  const status = await navigateToTeams(page, tokenWaitMs);

  if (status.isAuthenticated) {
    log('Tokens present — saving session state.');
    await saveSessionState(context);

    // Tokens are live, but the messaging cookies may still be a beat behind.
    if (await waitForUsableSession(page, context, log, SESSION_SETTLE_WAIT_MS)) {
      log('Session state saved.');
      return;
    }

    if (headless) {
      throw new Error('Headless SSO failed: tokens did not yield a usable session');
    }

    log('Please complete sign-in in the browser window.');
    await waitForManualLogin(page, context, undefined, onProgress, showOverlay);
    return;
  }

  // No usable tokens yet. Silent acquisition may still be in flight, so give it
  // the token wait unless Teams has explicitly refused. An empty auth store is
  // deliberately not treated as fatal here: waitForTokenRefresh reloads the page
  // and can recover, and a boot-time read can look empty when it is merely early.
  if (!status.isOnLoginPage && !status.interactionRequired) {
    const acquired = await waitForTokenRefresh(page, context, onProgress, { maxWaitMs: tokenWaitMs });
    if (acquired && await waitForUsableSession(page, context, log, SESSION_SETTLE_WAIT_MS)) {
      log('Session state saved.');
      return;
    }
  }

  // User interaction required - fail fast if headless
  if (headless) {
    const reason = status.isOnLoginPage
      ? 'Login page detected - user credentials required'
      : status.interactionRequired
        ? 'Teams reports InteractionRequired - no Entra ID session in the browser profile'
        : `No usable tokens at ${status.currentUrl}`;
    throw new Error(`Headless SSO failed: ${reason}`);
  }

  // Teams keeps rendering its cached UI when its session is dead, so it offers
  // no sign-in prompt to click. Drop the dead state to force the login redirect.
  if (status.interactionRequired) {
    log('No Entra ID session — clearing dead auth state to force a sign-in prompt.');
    await clearAuthStateForFreshLogin(page, context, onProgress);
    await page.goto(TEAMS_URL, { waitUntil: 'domcontentloaded' });
  }

  log('Login required. Please complete authentication in the browser window.');
  await waitForManualLogin(page, context, undefined, onProgress, showOverlay);
}

/**
 * Forces a new login by clearing cookies + MSAL cache and navigating to Teams.
 * Cookie-only clears leave expired access tokens that make the UI look signed-in
 * while Substrate search tokens never refresh (no refresh token).
 */
export async function forceNewLogin(
  page: Page,
  context: BrowserContext,
  onProgress?: (message: string) => void
): Promise<void> {
  const log = onProgress ?? console.log;

  log('Starting fresh login...');

  await clearAuthStateForFreshLogin(page, context, onProgress);

  // Navigate and wait for login
  await navigateToTeams(page);
  await waitForManualLogin(page, context, undefined, onProgress);
}
