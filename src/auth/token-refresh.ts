/**
 * Token refresh orchestrator.
 * 
 * Tries HTTP-based refresh first (direct OAuth2 token endpoint call, ~100ms),
 * then falls back to headless browser refresh (~8s) if HTTP fails.
 * 
 * HTTP refresh works by extracting the MSAL refresh token from session state
 * and exchanging it for new access tokens via Azure AD's token endpoint.
 * This works identically for standard Microsoft login and corporate SSO
 * (ADFS/Okta federation) — refresh tokens are standard Azure AD tokens
 * regardless of how the user originally authenticated.
 * 
 * Browser fallback covers cases where:
 * - The refresh token has expired (typically after days/weeks of inactivity)
 * - Conditional Access policies require interactive auth
 * - The MSAL cache format has changed unexpectedly
 * 
 * First login always requires a browser — there's no refresh token to use yet.
 */

import { TOKEN_REFRESH_THRESHOLD_MS } from '../constants.js';
import { ErrorCode, createError } from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';
import {
  extractSubstrateToken,
  extractMessageAuth,
  getMessageAuthStatus,
  clearTokenCache,
} from './token-extractor.js';
import { hasSessionState } from './session-store.js';
import { refreshTokensViaHttp, ensureMessagingCookies } from './token-refresh-http.js';
import * as log from '../utils/logger.js';

/** Nominal lifetime assumed for a messaging token whose expiry cannot be read. */
const ASSUMED_MESSAGING_LIFETIME_MS = 60 * 60 * 1000;

/**
 * Expiry of whatever credential currently makes the session usable: the
 * Substrate token when present, otherwise the messaging skypetoken.
 *
 * Substrate is routinely InteractionRequired on New Teams, so treating it as
 * the only proof of a successful refresh reports failure for sessions whose
 * messaging APIs are perfectly healthy — and then burns a second browser
 * launch trying to "fix" them.
 */
function usableSessionExpiry(): Date | null {
  const substrate = extractSubstrateToken();
  if (substrate && substrate.expiry.getTime() > Date.now()) {
    return substrate.expiry;
  }

  if (!extractMessageAuth()) return null;

  const messaging = getMessageAuthStatus();
  if (!messaging.hasToken) return null;

  return messaging.expiresAt
    ? new Date(messaging.expiresAt)
    : new Date(Date.now() + ASSUMED_MESSAGING_LIFETIME_MS);
}

/** Result of a successful token refresh. */
export interface TokenRefreshResult {
  /** New token expiry time. */
  newExpiry: Date;
  /** Previous expiry time (for comparison). */
  previousExpiry: Date;
  /** Minutes gained by refresh. */
  minutesGained: number;
  /** Whether a refresh was actually needed (token was close to expiry). */
  refreshNeeded: boolean;
  /** Which method was used: 'http' or 'browser'. */
  method: 'http' | 'browser';
}

/** Module-level flag to prevent concurrent refresh attempts. */
let refreshInProgress = false;

/**
 * Builds a TokenRefreshResult after a successful refresh.
 * When no prior valid Substrate token existed (expired/missing), treat refresh as required.
 */
function buildRefreshResult(
  afterToken: { expiry: Date },
  previousExpiry: Date,
  hadValidTokenBefore: boolean,
  method: 'http' | 'browser',
): TokenRefreshResult {
  const minutesGained = Math.round(
    (afterToken.expiry.getTime() - previousExpiry.getTime()) / 1000 / 60
  );
  const refreshNeeded = !hadValidTokenBefore
    || previousExpiry.getTime() - Date.now() < TOKEN_REFRESH_THRESHOLD_MS;

  return {
    newExpiry: afterToken.expiry,
    previousExpiry,
    minutesGained,
    refreshNeeded,
    method,
  };
}

/**
 * Refreshes tokens using HTTP-first strategy with browser fallback.
 * 
 * 1. Try direct HTTP refresh via OAuth2 token endpoint (~100ms)
 * 2. If HTTP fails, fall back to headless browser refresh (~8s)
 * 3. If both fail, return error directing to teams_login
 *
 * Expired Substrate access tokens are not a blocker: extractSubstrateToken()
 * only returns non-expired tokens, but MSAL refresh tokens in session state
 * can still mint new access tokens via HTTP (or browser SSO).
 */
export async function refreshTokensViaBrowser(): Promise<Result<TokenRefreshResult>> {
  // Prevent concurrent refresh attempts
  if (refreshInProgress) {
    return err(createError(
      ErrorCode.UNKNOWN,
      'Token refresh already in progress. Please wait and try again.',
      { retryable: true, suggestions: ['Wait a moment and retry the request'] }
    ));
  }

  // Need a saved session (cookies / MSAL cache). Access-token presence is optional —
  // expired Substrate tokens are filtered out by extractSubstrateToken(), but the
  // refresh token in session state is enough for HTTP refresh.
  if (!hasSessionState()) {
    log.warn('token-refresh', 'No session state found - cannot refresh, browser login required');
    return err(createError(
      ErrorCode.AUTH_REQUIRED,
      'ACTION REQUIRED: No session found. You MUST call teams_login to authenticate.',
    ));
  }

  const beforeToken = extractSubstrateToken();
  const hadValidTokenBefore = beforeToken !== null;
  // If access token is already gone/expired, treat previous expiry as just-expired
  // so minutesGained reflects the new token lifetime rather than epoch math.
  const previousExpiry = beforeToken?.expiry ?? new Date(Date.now() - 1);

  if (beforeToken) {
    log.debug(
      'token-refresh',
      `Current token expires at ${beforeToken.expiry.toISOString()} (${Math.round((beforeToken.expiry.getTime() - Date.now()) / 60000)} mins remaining)`
    );
  } else {
    log.info(
      'token-refresh',
      'No valid Substrate access token in session (expired or missing); attempting refresh via MSAL refresh token / browser SSO'
    );
  }

  refreshInProgress = true;

  try {
    // ── Strategy 1: HTTP refresh (fast, no browser needed) ──────────────
    const httpResult = await refreshTokensViaHttp();

    if (httpResult.ok) {
      log.info('token-refresh', `HTTP refresh succeeded: ${httpResult.value.tokensRefreshed} tokens refreshed` +
        (httpResult.value.skypeTokenRefreshed ? ', skype token refreshed' : '') +
        (httpResult.value.refreshTokenRotated ? ', refresh token rotated' : ''));

      const expiry = usableSessionExpiry();
      if (expiry) {
        return ok(buildRefreshResult({ expiry }, previousExpiry, hadValidTokenBefore, 'http'));
      }

      // HTTP refresh reported success but we can't extract a valid token — fall through
      log.warn('token-refresh', 'HTTP refresh reported success but no usable token found, falling back to browser');
    } else {
      log.warn('token-refresh', `HTTP refresh failed: ${httpResult.error.message}, falling back to browser`);

      // If the error is definitively an auth error (expired refresh token),
      // don't bother with browser fallback — it won't help either
      if (httpResult.error.code === ErrorCode.AUTH_EXPIRED) {
        // Still try browser — the persistent profile's session cookies may work
        log.info('token-refresh', 'Auth expired, but trying browser fallback (session cookies may still be valid)');
      }
    }

    // ── Strategy 2: Re-mint messaging cookies from an existing token ─────
    // New Teams keeps its tokens in the encrypted tmp.auth store, so the hourly
    // Skype Spaces token often outlives the messaging cookies. Exchanging it is
    // one HTTP call and skips a ~8s browser launch entirely.
    const cookieResult = await ensureMessagingCookies();
    if (cookieResult.ok && cookieResult.value) {
      const expiry = usableSessionExpiry();
      if (expiry) {
        log.info('token-refresh', 'Recovered messaging auth from the stored Skype Spaces token');
        return ok(buildRefreshResult({ expiry }, previousExpiry, hadValidTokenBefore, 'http'));
      }
    }

    // ── Strategy 3: Browser refresh (fallback) ──────────────────────────
    return await refreshTokensViaBrowserImpl(previousExpiry, hadValidTokenBefore);

  } finally {
    refreshInProgress = false;
  }
}

/**
 * Browser-based token refresh implementation.
 * Opens a headless browser with the persistent profile, lets MSAL
 * silently refresh tokens using session cookies.
 */
async function refreshTokensViaBrowserImpl(
  previousExpiry: Date,
  hadValidTokenBefore: boolean,
): Promise<Result<TokenRefreshResult>> {
  // Import browser functions dynamically to avoid circular dependencies
  const { createBrowserContext, closeBrowser } = await import('../browser/context.js');

  let manager: Awaited<ReturnType<typeof createBrowserContext>> | null = null;

  try {
    // Open headless browser with persistent profile
    manager = await createBrowserContext({ headless: true });

    // Import auth functions
    const { ensureAuthenticated } = await import('../browser/auth.js');

    // Use the same auth flow that works for login - this triggers token acquisition
    // showOverlay: false since headless browser has no visible window
    // headless: true to fail fast if user interaction is required
    await ensureAuthenticated(manager.page, manager.context, (msg) => {
      // Silent logging for headless refresh
      log.debug('token-refresh', msg);
    }, false, true);

    // Close browser (ensureAuthenticated already saved the session)
    await closeBrowser(manager, false);
    manager = null;

    // Clear our token cache to force re-extraction from the new session
    clearTokenCache();

    // Accept messaging-only recovery: Substrate is often unavailable on New
    // Teams, and messaging is what most tools actually need.
    const expiry = usableSessionExpiry();
    if (!expiry) {
      return err(createError(
        ErrorCode.AUTH_EXPIRED,
        'ACTION REQUIRED: Token refresh failed. You MUST call teams_login to re-authenticate.',
      ));
    }

    const result = buildRefreshResult({ expiry }, previousExpiry, hadValidTokenBefore, 'browser');

    // If we needed a refresh but didn't get a newer expiry, that's an error
    if (result.refreshNeeded && expiry.getTime() <= previousExpiry.getTime()) {
      return err(createError(
        ErrorCode.AUTH_EXPIRED,
        'ACTION REQUIRED: Token was not refreshed despite being close to expiry. You MUST call teams_login to re-authenticate.',
      ));
    }

    return ok(result);

  } catch (error) {
    // Clean up browser if still open
    if (manager) {
      try {
        await closeBrowser(manager, false);
      } catch {
        // Ignore cleanup errors
      }
    }

    const message = error instanceof Error ? error.message : 'Unknown error';
    return err(createError(
      ErrorCode.UNKNOWN,
      `Token refresh via browser failed: ${message}. Call teams_login to re-authenticate.`,
    ));
  }
}
