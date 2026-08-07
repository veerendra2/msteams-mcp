/**
 * Lightweight logger for consistent log output.
 *
 * All production logging goes through this module so that:
 * - Output is consistently prefixed with `[module]` context
 * - Log levels can be controlled via LOG_LEVEL environment variable
 * - Future structured logging can be added in one place
 *
 * MCP note: stdio transport uses stdout for JSON-RPC only. Every log line
 * must go to stderr or Cursor will treat it as a protocol parse error.
 *
 * Levels: error | warn | info (default) | debug
 * Set LOG_LEVEL=debug for verbose output during development.
 */

type LogLevel = 'error' | 'warn' | 'info' | 'debug';

const LEVEL_PRIORITY: Record<LogLevel, number> = {
  error: 0,
  warn: 1,
  info: 2,
  debug: 3,
};

function getConfiguredLevel(): LogLevel {
  const env = process.env.LOG_LEVEL?.toLowerCase();
  if (env && env in LEVEL_PRIORITY) return env as LogLevel;
  return 'info';
}

let currentLevel = getConfiguredLevel();

function shouldLog(level: LogLevel): boolean {
  return LEVEL_PRIORITY[level] <= LEVEL_PRIORITY[currentLevel];
}

/** Write a log line to stderr (never stdout — MCP JSON-RPC lives there). */
function write(context: string, message: string): void {
  // Use stderr.write — not console.error. Cursor's MCP log capture treats
  // console.* as Console API messages and can append a stray "undefined"
  // (the void return) on the next line. Raw stderr is also one clean line.
  process.stderr.write(`[${context}] ${message}\n`);
}

/** Log an error message. Always shown unless level is somehow above error. */
export function error(context: string, message: string): void {
  if (shouldLog('error')) {
    write(context, message);
  }
}

/** Log a warning message. */
export function warn(context: string, message: string): void {
  if (shouldLog('warn')) {
    write(context, message);
  }
}

/** Log an informational message. */
export function info(context: string, message: string): void {
  if (shouldLog('info')) {
    write(context, message);
  }
}

/** Log a debug message. Only shown when LOG_LEVEL=debug. */
export function debug(context: string, message: string): void {
  if (shouldLog('debug')) {
    write(context, message);
  }
}
