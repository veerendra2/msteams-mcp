/**
 * Authentication-related tool handlers.
 */

import { createRequire } from 'module';
import { z } from 'zod';
import type { Tool } from '@modelcontextprotocol/sdk/types.js';
import type { RegisteredTool, ToolContext, ToolResult } from './index.js';

const require = createRequire(import.meta.url);
const pkg = require('../../package.json') as { version: string };
import {
  hasSessionState,
  isSessionLikelyExpired,
  clearSessionState,
  readSessionState,
} from '../auth/session-store.js';
import {
  getSubstrateTokenStatus,
  getMessageAuthStatus,
  extractMessageAuth,
  extractCsaToken,
  clearTokenCache,
  hasValidSubstrateToken,
} from '../auth/token-extractor.js';
import { readTmpAuthState, INTERACTION_REQUIRED } from '../auth/tmp-auth-tokens.js';
import { createBrowserContext, closeBrowser, ensureBrowserProfileUnlocked, hasBrowserProfile } from '../browser/context.js';
import * as log from '../utils/logger.js';
import { ensureAuthenticated, forceNewLogin, getAuthStatus } from '../browser/auth.js';
import { ErrorCode, createError } from '../types/errors.js';

// ─────────────────────────────────────────────────────────────────────────────
// Schemas
// ─────────────────────────────────────────────────────────────────────────────

export const LoginInputSchema = z.object({
  forceNew: z.boolean().optional().default(false),
});

// ─────────────────────────────────────────────────────────────────────────────
// Tool Definitions
// ─────────────────────────────────────────────────────────────────────────────

const loginToolDefinition: Tool = {
  name: 'teams_login',
  description: 'Trigger manual login flow for Microsoft Teams. Use this if the session has expired or you need to switch accounts.',
  inputSchema: {
    type: 'object',
    properties: {
      forceNew: {
        type: 'boolean',
        description: 'Force a new login even if a session exists (default: false)',
      },
    },
  },
};

const statusToolDefinition: Tool = {
  name: 'teams_status',
  description: 'Check the current authentication status and session state.',
  inputSchema: {
    type: 'object',
    properties: {},
  },
};

// ─────────────────────────────────────────────────────────────────────────────
// Handlers
// ─────────────────────────────────────────────────────────────────────────────

/** Minimum minutes remaining on token to consider it valid (skip browser). */
const TOKEN_VALID_THRESHOLD_MINUTES = 10;

/**
 * Returns true when the saved session could plausibly be refreshed without the
 * user typing credentials.
 *
 * Classic sessions need an MSAL RefreshToken. New Teams never stores one — it
 * re-acquires against the Entra ID session cookie — so a live token in the
 * `tmp.auth` store is the equivalent signal. Requiring a refresh token alone
 * meant headless SSO was skipped unconditionally on New Teams.
 */
function canAttemptSilentRefresh(): boolean {
  const state = readSessionState();
  if (!state) return false;

  for (const origin of state.origins ?? []) {
    for (const item of origin.localStorage ?? []) {
      if (!item.value?.startsWith('{')) continue;
      try {
        const entry = JSON.parse(item.value) as { credentialType?: string; secret?: string };
        if (
          entry.credentialType === 'RefreshToken' &&
          typeof entry.secret === 'string' &&
          entry.secret.length > 0
        ) {
          return true;
        }
      } catch {
        continue;
      }
    }
  }

  return readTmpAuthState(state).usableTokens > 0;
}

/**
 * Ensures login left usable messaging auth (skypetoken + authtoken).
 * Substrate search is optional on New Teams (often InteractionRequired).
 */
function requireMessagingAfterLogin(): ToolResult | null {
  if (extractMessageAuth()) {
    return null;
  }

  // Distinguish "sign-in never happened" from "signed in but cookies missing";
  // telling the caller to retry forceNew when no session exists just loops.
  const authState = readTmpAuthState();
  const needsInteractive = authState.usableTokens === 0 && authState.state === INTERACTION_REQUIRED;

  return {
    success: false,
    error: createError(
      ErrorCode.AUTH_REQUIRED,
      needsInteractive
        ? 'Sign-in did not complete: Teams has no Entra ID session in its browser profile, so it cannot mint tokens. ' +
          'Complete the Microsoft sign-in in the browser window that opens (including MFA) and tick "Stay signed in" ' +
          'so the session survives browser restarts.'
        : 'Login finished but the messaging cookies were not issued. Retry the request; if it persists, ' +
          'run teams_login with forceNew: true and complete sign-in again.',
      {
        suggestions: needsInteractive
          ? [
              'Tell the user to complete the sign-in prompt in the visible browser window',
              'Ensure "Stay signed in" is ticked so the session persists',
              'Do not silently retry forceNew — it cannot succeed without the user signing in',
            ]
          : [
              'Retry the original request once',
              'If it still fails, call teams_login with { "forceNew": true }',
            ],
      }
    ),
  };
}

/** Builds the login success payload, noting Substrate availability. */
function loginSuccessData(message: string): Record<string, unknown> {
  const messaging = extractMessageAuth() !== null;
  const substrate = hasValidSubstrateToken();
  return {
    message: substrate
      ? message
      : `${message} Messaging APIs are ready; Substrate search is unavailable (channel find uses membership list).`,
    capabilities: {
      messaging,
      substrateSearch: substrate,
      favorites: messaging && extractCsaToken() !== null,
    },
  };
}

async function handleLogin(
  input: z.infer<typeof LoginInputSchema>,
  ctx: ToolContext
): Promise<ToolResult> {
  // Close existing browser if any
  const existingManager = ctx.server.getBrowserManager();
  if (existingManager) {
    await closeBrowser(existingManager, !input.forceNew);
    ctx.server.resetBrowserState();
  }

  if (input.forceNew) {
    clearSessionState();
    clearTokenCache();
  }

  // Fast path: skip browser when messaging session is still usable.
  // Substrate alone is not enough — messaging cookies are the ground truth
  // (same rule as requireMessagingAfterLogin after browser flows).
  if (!input.forceNew) {
    const tokenStatus = getSubstrateTokenStatus();
    if (tokenStatus.hasToken &&
        tokenStatus.minutesRemaining !== undefined &&
        tokenStatus.minutesRemaining >= TOKEN_VALID_THRESHOLD_MINUTES) {
      if (extractMessageAuth()) {
        ctx.server.markInitialised();
        return {
          success: true,
          data: {
            ...loginSuccessData(
              `Already authenticated. Substrate token valid for ${tokenStatus.minutesRemaining} more minutes.`
            ),
            tokenStatus: {
              expiresAt: tokenStatus.expiresAt,
              minutesRemaining: tokenStatus.minutesRemaining,
            },
          },
        };
      }
      log.info('login', 'Substrate valid but messaging auth missing/expired; continuing refresh');
    }

    const messageAuth = extractMessageAuth();
    if (messageAuth) {
      ctx.server.markInitialised();
      return {
        success: true,
        data: loginSuccessData('Already authenticated (messaging session).'),
      };
    }
  }

  // Try headless first whenever there is anything to try it with. The browser
  // profile carries the Entra ID session, so a stale session file alone is no
  // reason to interrupt the user with a visible window.
  const tryHeadless = !input.forceNew && (canAttemptSilentRefresh() || hasBrowserProfile());
  if (!tryHeadless) {
    log.info(
      'login',
      input.forceNew
        ? 'Skipping headless SSO for forceNew — opening visible browser for interactive login'
        : 'Skipping headless SSO (no browser profile yet) — opening visible browser'
    );
  }

  if (tryHeadless) {
    const headlessManager = await createBrowserContext({ headless: true });
    ctx.server.setBrowserManager(headlessManager);

    try {
      await ensureAuthenticated(
        headlessManager.page,
        headlessManager.context,
        (msg) => {
          log.info('login:headless', msg);
        },
        false, // No overlay in headless
        true   // Headless mode - throw immediately if user interaction required
      );

      await closeBrowser(headlessManager, true);
      ctx.server.resetBrowserState();

      const tokenFailure = requireMessagingAfterLogin();
      if (tokenFailure) {
        return tokenFailure;
      }

      ctx.server.markInitialised();

      return {
        success: true,
        data: loginSuccessData('Login completed silently via SSO. Session has been saved.'),
      };
    } catch (error) {
      // Headless attempt failed - fall through to visible browser
      log.warn('login:headless', `Headless SSO failed, falling back to visible browser: ${error instanceof Error ? error.message : String(error)}`);
      try {
        await closeBrowser(headlessManager, false);
      } catch {
        // Ignore cleanup errors
      }
      ctx.server.resetBrowserState();
      // Release profile lock before opening a visible Chrome (SingletonLock races)
      await ensureBrowserProfileUnlocked();
    }
  }

  // Open visible browser for user interaction
  await ensureBrowserProfileUnlocked();
  const browserManager = await createBrowserContext({ headless: false });
  ctx.server.setBrowserManager(browserManager);

  try {
    // forceNew always clears; otherwise prefer soft ensure (messaging is enough).
    if (input.forceNew) {
      await forceNewLogin(
        browserManager.page,
        browserManager.context,
        (msg) => {
          log.info('login', msg);
        }
      );
    } else {
      await ensureAuthenticated(
        browserManager.page,
        browserManager.context,
        (msg) => {
          log.info('login', msg);
        }
      );
    }
  } finally {
    // Close browser after login - we only need the saved session/tokens
    await closeBrowser(browserManager, true);
    ctx.server.resetBrowserState();
  }

  const tokenFailure = requireMessagingAfterLogin();
  if (tokenFailure) {
    return tokenFailure;
  }

  ctx.server.markInitialised();

  return {
    success: true,
    data: loginSuccessData('Login completed successfully. Session has been saved.'),
  };
}

async function handleStatus(
  _input: Record<string, never>,
  ctx: ToolContext
): Promise<ToolResult> {
  const sessionExists = hasSessionState();
  const sessionExpired = isSessionLikelyExpired();
  const tokenStatus = getSubstrateTokenStatus();
  const messageAuthStatus = getMessageAuthStatus();
  // Use the same check messaging APIs require (skypetoken_asm + authtoken),
  // not the weaker skypetoken-only diagnostic.
  const messageAuth = extractMessageAuth();
  const csaToken = extractCsaToken();

  let authStatus = null;
  const browserManager = ctx.server.getBrowserManager();
  if (browserManager && ctx.server.isInitialisedState()) {
    authStatus = await getAuthStatus(browserManager.page);
  }

  const tmpAuth = readTmpAuthState();

  return {
    success: true,
    data: {
      version: pkg.version,
      directApi: {
        available: tokenStatus.hasToken,
        expiresAt: tokenStatus.expiresAt,
        minutesRemaining: tokenStatus.minutesRemaining,
      },
      // Ground truth for "is Teams actually signed in": live token slots vs
      // interaction-required placeholders.
      teamsAuthService: {
        state: tmpAuth.state,
        usableTokens: tmpAuth.usableTokens,
        placeholderTokens: tmpAuth.placeholderTokens,
        // True only when nothing usable exists AND nothing is left to refresh
        // from. Live tokens with missing cookies are recoverable silently, so
        // they must not be reported as needing the user.
        needsInteractiveLogin:
          !tokenStatus.hasToken &&
          messageAuth === null &&
          (tmpAuth.usableTokens === 0 || tmpAuth.state === INTERACTION_REQUIRED),
      },
      messaging: {
        available: messageAuth !== null,
        // Only attach skype expiry when messaging auth is actually usable;
        // otherwise minutesRemaining can look healthy while available is false
        // (e.g. live skypetoken_asm + expired authtoken).
        ...(messageAuth
          ? {
              expiresAt: messageAuthStatus.expiresAt,
              minutesRemaining: messageAuthStatus.minutesRemaining,
            }
          : {}),
      },
      favorites: {
        available: messageAuth !== null && csaToken !== null,
      },
      session: {
        exists: sessionExists,
        likelyExpired: sessionExpired,
      },
      browser: {
        running: browserManager !== null,
        initialised: ctx.server.isInitialisedState(),
      },
      authentication: authStatus,
    },
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// Exports
// ─────────────────────────────────────────────────────────────────────────────

export const loginTool: RegisteredTool<typeof LoginInputSchema> = {
  definition: loginToolDefinition,
  schema: LoginInputSchema,
  handler: handleLogin,
};

export const statusTool: RegisteredTool<z.ZodObject<Record<string, never>>> = {
  definition: statusToolDefinition,
  schema: z.object({}),
  handler: handleStatus,
};

/** All auth-related tools. */
export const authTools = [loginTool, statusTool];
