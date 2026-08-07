/**
 * Playwright browser context management.
 * Creates and manages browser contexts with session persistence.
 * 
 * Uses the system's installed Chrome or Edge browser rather than downloading
 * Playwright's bundled Chromium. This significantly reduces install size.
 * 
 * All modes (headless and visible) use a persistent browser profile stored at
 * ~/.teams-mcp-server/browser-profile/. This means:
 * - Microsoft session cookies persist across launches (longer-lived than MSAL tokens)
 * - Headless token refresh can silently re-authenticate using the profile's session
 * - Visible login retains extensions (e.g. Bitwarden) and form autofill data
 * - No need for storageState temp files or encrypted session restoration for browser use
 */

import { execFileSync } from 'child_process';
import { chromium, type BrowserContext, type Page } from 'playwright';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  ensureUserDataDir,
  CONFIG_DIR,
  writeSessionState,
} from '../auth/session-store.js';
import { clearRegionCache } from '../utils/auth-guards.js';
import * as log from '../utils/logger.js';

export interface BrowserManager {
  /** Always null — persistent contexts have no separate Browser object. */
  browser: null;
  context: BrowserContext;
  page: Page;
  isNewSession: boolean;
  /** Always true — all contexts use the persistent browser profile. */
  persistent: true;
}

export interface CreateBrowserOptions {
  headless?: boolean;
  viewport?: { width: number; height: number };
}

const DEFAULT_OPTIONS: Required<CreateBrowserOptions> = {
  headless: true,
  viewport: { width: 1280, height: 800 },
};

/**
 * Directory for the persistent browser profile.
 * This is a dedicated Chrome/Edge profile within the config dir, so extensions
 * (e.g. Bitwarden) and form autofill data persist across login sessions
 * without conflicting with the user's running browser instance.
 * 
 * Both headless and visible modes share this profile, so Microsoft's long-lived
 * session cookies enable silent headless re-authentication without user interaction.
 */
const BROWSER_PROFILE_DIR = path.join(CONFIG_DIR, 'browser-profile');

/**
 * Returns true if the persistent browser profile has already been created.
 *
 * The profile — not the saved session file — holds the Entra ID session, so it
 * is worth a headless attempt even when the saved session looks dead.
 */
export function hasBrowserProfile(): boolean {
  return fs.existsSync(path.join(BROWSER_PROFILE_DIR, 'Default'));
}

/**
 * Path to the Chromium SingletonLock file.
 * This file prevents multiple processes from using the same profile.
 */
const SINGLETON_LOCK_PATH = path.join(BROWSER_PROFILE_DIR, 'SingletonLock');

/** Other Chromium singleton files that can block relaunch after a crash. */
const SINGLETON_AUX_PATHS = [
  path.join(BROWSER_PROFILE_DIR, 'SingletonCookie'),
  path.join(BROWSER_PROFILE_DIR, 'SingletonSocket'),
];

/**
 * Checks whether a path exists without following symlinks.
 *
 * Chromium's singleton files are symlinks whose targets never exist — the
 * target string encodes `<hostname>-<pid>`. `fs.existsSync` follows the link
 * and so reports `false` for a lock that is very much present, which silently
 * disables every cleanup path below and leaves the profile unlaunchable.
 */
function lstatExists(filePath: string): boolean {
  try {
    fs.lstatSync(filePath);
    return true;
  } catch {
    return false;
  }
}

/**
 * Checks if a process with the given PID is running.
 *
 * EPERM means the process exists but belongs to another user, so it must count
 * as running — treating it as dead would remove a live Chrome's lock and risk
 * profile corruption.
 */
function isProcessRunning(pid: number): boolean {
  try {
    // kill with signal 0 checks if process exists without killing it
    process.kill(pid, 0);
    return true;
  } catch (error) {
    return (error as NodeJS.ErrnoException).code === 'EPERM';
  }
}

/**
 * Removes Chromium singleton lock files for the MCP browser profile.
 */
function removeSingletonFiles(): void {
  for (const lockPath of [SINGLETON_LOCK_PATH, ...SINGLETON_AUX_PATHS]) {
    try {
      fs.unlinkSync(lockPath);
    } catch {
      // Already gone, or not removable — best-effort cleanup.
    }
  }
}

/** How long to wait for orphaned profile browsers to exit before escalating (ms). */
const ORPHAN_EXIT_TIMEOUT_MS = 5000;

/** Poll interval while waiting for orphaned browsers to exit (ms). */
const ORPHAN_POLL_INTERVAL_MS = 250;

/**
 * Finds Chrome/Edge processes still using our persistent profile.
 *
 * The `user-data-dir=` argument only ever appears in a browser's argv, so this
 * cannot match unrelated processes.
 */
function findProfileBrowserPids(): number[] {
  if (process.platform === 'win32') {
    return [];
  }

  try {
    const output = execFileSync(
      'pgrep',
      ['-f', `user-data-dir=${BROWSER_PROFILE_DIR}`],
      { encoding: 'utf8' }
    );
    return output
      .split('\n')
      .map(line => parseInt(line.trim(), 10))
      .filter(pid => Number.isFinite(pid) && pid > 0 && pid !== process.pid);
  } catch {
    // pgrep exit 1 = no matches
    return [];
  }
}

/**
 * Terminates browsers holding our profile and waits for them to actually exit.
 *
 * Signalling and moving straight on is not enough: a Chrome process family
 * takes seconds to unwind, and launching a new instance while the old one is
 * still attached to the profile yields a context whose every `evaluate` call
 * fails — which reads downstream as "no tokens" rather than as contention.
 */
async function terminateProfileBrowsers(): Promise<void> {
  let pids = findProfileBrowserPids();
  if (pids.length === 0) return;

  for (const pid of pids) {
    try {
      process.kill(pid, 'SIGTERM');
      log.info('browser', `Sent SIGTERM to orphan profile browser PID ${pid}`);
    } catch {
      // Already gone or not permitted
    }
  }

  const deadline = Date.now() + ORPHAN_EXIT_TIMEOUT_MS;
  while (Date.now() < deadline) {
    await new Promise(resolve => setTimeout(resolve, ORPHAN_POLL_INTERVAL_MS));
    pids = findProfileBrowserPids();
    if (pids.length === 0) return;
  }

  // Playwright sometimes cannot kill Chrome (EPERM) and SIGTERM can be ignored
  // during shutdown; escalate rather than launching into a contended profile.
  for (const pid of pids) {
    try {
      process.kill(pid, 'SIGKILL');
      log.warn('browser', `Force killed unresponsive profile browser PID ${pid}`);
    } catch {
      // Already gone or not permitted
    }
  }
  await new Promise(resolve => setTimeout(resolve, ORPHAN_POLL_INTERVAL_MS * 2));
}

/**
 * Cleans up stale SingletonLock files left by crashed browser sessions.
 * 
 * On Unix systems, SingletonLock is a symlink whose target contains the PID.
 * On Windows, it's a regular file. If the owning process is no longer running,
 * we can safely remove the lock.
 * 
 * @returns true if a stale lock was removed, false otherwise
 */
function cleanupStaleSingletonLock(): boolean {
  if (!lstatExists(SINGLETON_LOCK_PATH)) {
    return false;
  }

  try {
    // On Unix, SingletonLock is a symlink like "hostname-12345"
    // The number after the last dash is the PID
    const linkTarget = fs.readlinkSync(SINGLETON_LOCK_PATH);
    const pidMatch = linkTarget.match(/^(.*)-(\d+)$/);

    if (pidMatch) {
      const lockHostname = pidMatch[1];
      const pid = parseInt(pidMatch[2], 10);

      // A lock written under a different hostname cannot be validated by PID on
      // this machine. Hostnames here are DHCP-derived and change between
      // networks, so an old one means the lock is stale, not foreign.
      if (lockHostname !== os.hostname()) {
        log.info('browser', `Removing SingletonLock from another hostname (${lockHostname})`);
        removeSingletonFiles();
        return true;
      }

      if (isProcessRunning(pid)) {
        // Process is still running - lock is valid
        return false;
      }

      // Process is dead - remove stale lock (+ aux singleton files)
      log.info('browser', `Removing stale SingletonLock (PID ${pid} not running)`);
      removeSingletonFiles();
      return true;
    }
  } catch {
    // On Windows or if readlink fails, try checking file age
    // If the lock file is old and we can't verify the process, remove it
    try {
      const stats = fs.lstatSync(SINGLETON_LOCK_PATH);
      const ageMs = Date.now() - stats.mtimeMs;
      const oneHourMs = 60 * 60 * 1000;
      
      if (ageMs > oneHourMs) {
        log.info('browser', `Removing old SingletonLock (${Math.round(ageMs / 60000)} minutes old)`);
        removeSingletonFiles();
        return true;
      }
    } catch {
      // Can't read stats, leave it alone
    }
  }

  return false;
}

/**
 * Ensures the persistent browser profile is not locked by a dead/orphan Chrome.
 * Call between headless and visible login, and before retrying a failed launch.
 */
export async function ensureBrowserProfileUnlocked(): Promise<void> {
  cleanupStaleSingletonLock();

  // A live browser on the profile is contention even without a lock file, and a
  // lock without a live browser is a leftover. Both block a clean launch.
  if (!lstatExists(SINGLETON_LOCK_PATH) && findProfileBrowserPids().length === 0) {
    return;
  }

  await terminateProfileBrowsers();
  removeSingletonFiles();
}

/**
 * Determines the browser channel to use based on the platform.
 * - Windows: Use Microsoft Edge (always installed on Windows 10+)
 * - macOS/Linux: Use Chrome
 * 
 * @returns The browser channel name for Playwright
 */
function getBrowserChannel(): 'msedge' | 'chrome' {
  return process.platform === 'win32' ? 'msedge' : 'chrome';
}

/**
 * Creates a browser context using a persistent profile.
 *
 * Uses the system's installed Chrome or Edge browser rather than downloading
 * Playwright's bundled Chromium (~180MB savings).
 *
 * The persistent profile at ~/.teams-mcp-server/browser-profile/ is shared
 * between headless and visible modes. This provides:
 * - Silent headless re-auth via long-lived Microsoft session cookies
 * - Extensions (e.g. Bitwarden) and form autofill for visible login
 * - No storageState temp file management needed
 *
 * Note: Only one process can use the profile at a time (Chromium profile lock).
 * The MCP server serialises tool calls, and token-refresh checks for an active
 * browser before attempting refresh to avoid lock contention.
 *
 * @param options - Browser configuration options
 * @returns Browser manager with context and page
 * @throws Error if system browser is not found (with helpful suggestions)
 */
export async function createBrowserContext(
  options: CreateBrowserOptions = {}
): Promise<BrowserManager> {
  const opts = { ...DEFAULT_OPTIONS, ...options };

  ensureUserDataDir();

  const channel = getBrowserChannel();

  // Clean up any stale lock files before attempting to launch
  await ensureBrowserProfileUnlocked();

  const launchBrowser = async (): Promise<BrowserManager> => {
    const context = await chromium.launchPersistentContext(BROWSER_PROFILE_DIR, {
      headless: opts.headless,
      channel,
      viewport: opts.viewport,
      acceptDownloads: false,
    });

    // Persistent contexts start with one page; use it or create one
    const page = context.pages()[0] ?? await context.newPage();

    return {
      browser: null,
      context,
      page,
      isNewSession: true,
      persistent: true,
    };
  };

  try {
    return await launchBrowser();
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    
    // Check if this is a profile lock error and try to recover
    if (errorMessage.includes('ProcessSingleton') || errorMessage.includes('SingletonLock')) {
      log.warn('browser', 'Profile lock detected, attempting to clean up and retry...');
      try {
        await ensureBrowserProfileUnlocked();
        log.info('browser', 'Cleared profile lock, retrying browser launch...');
        return await launchBrowser();
      } catch (cleanupError) {
        log.error('browser', `Failed to recover from SingletonLock: ${cleanupError instanceof Error ? cleanupError.message : String(cleanupError)}`);
      }
    }

    const browserName = channel === 'msedge' ? 'Microsoft Edge' : 'Google Chrome';
    const installHint = channel === 'msedge'
      ? 'Edge should be pre-installed on Windows. Try updating Windows or reinstalling Edge.'
      : 'Install Chrome from https://www.google.com/chrome/ or run: npx playwright install chromium';

    throw new Error(
      `Could not launch ${browserName}. ${installHint}\n\n` +
      `Original error: ${errorMessage}`
    );
  }
}

/**
 * Saves the current browser context's session state.
 */
export async function saveSessionState(context: BrowserContext): Promise<void> {
  const state = await context.storageState();
  writeSessionState(state);
  // Clear cached region config so new session values are picked up
  clearRegionCache();
}

/**
 * Closes the browser context and optionally saves session state.
 */
export async function closeBrowser(
  manager: BrowserManager,
  saveSession: boolean = true
): Promise<void> {
  if (saveSession) {
    await saveSessionState(manager.context);
  }
  try {
    await manager.context.close();
  } finally {
    // Playwright sometimes fails to kill Chrome (EPERM), leaving SingletonLock held.
    await ensureBrowserProfileUnlocked();
  }
}
