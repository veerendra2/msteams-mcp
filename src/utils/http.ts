/**
 * HTTP utilities with retry, timeout, and error handling.
 */

import { 
  ErrorCode, 
  createError, 
  classifyHttpError, 
  extractRetryAfter,
  type McpError,
} from '../types/errors.js';
import { type Result, ok, err } from '../types/result.js';

/** Options for HTTP requests. */
export interface HttpOptions extends Omit<RequestInit, 'signal'> {
  /** Timeout in milliseconds (default: 30000). */
  timeoutMs?: number;
  /** Maximum retry attempts (default: 3). */
  maxRetries?: number;
  /** Base delay for exponential backoff in ms (default: 1000). */
  retryBaseDelayMs?: number;
  /** Maximum delay between retries in ms (default: 10000). */
  retryMaxDelayMs?: number;
}

/** Successful HTTP response. */
export interface HttpResponse<T = unknown> {
  status: number;
  headers: Headers;
  data: T;
}

/** Per-host rate limit deadlines (epoch ms). */
const rateLimitedUntilByHost = new Map<string, number>();

/** Resolves the host key used for rate-limit tracking. */
function rateLimitHost(url: string): string {
  try {
    return new URL(url).host;
  } catch {
    return url;
  }
}

/**
 * Makes an HTTP request with timeout, retry, and error handling.
 */
export async function httpRequest<T = unknown>(
  url: string,
  options: HttpOptions = {}
): Promise<Result<HttpResponse<T>>> {
  const {
    timeoutMs = 30000,
    maxRetries = 3,
    retryBaseDelayMs = 1000,
    retryMaxDelayMs = 10000,
    ...fetchOptions
  } = options;

  const host = rateLimitHost(url);

  // Check rate limit state for this host only
  const hostLimitedUntil = rateLimitedUntilByHost.get(host);
  if (hostLimitedUntil && Date.now() < hostLimitedUntil) {
    const waitMs = hostLimitedUntil - Date.now();
    return err(createError(
      ErrorCode.RATE_LIMITED,
      `Rate limited. Retry after ${Math.ceil(waitMs / 1000)}s`,
      { retryable: true, retryAfterMs: waitMs }
    ));
  }

  let lastError: McpError | null = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const result = await fetchWithTimeout<T>(url, fetchOptions, timeoutMs);
      
      if (result.ok) {
        return result;
      }

      lastError = result.error;

      // Handle rate limiting — scoped to this host
      if (result.error.code === ErrorCode.RATE_LIMITED && result.error.retryAfterMs) {
        rateLimitedUntilByHost.set(host, Date.now() + result.error.retryAfterMs);
      }

      // Don't retry non-retryable errors
      if (!result.error.retryable) {
        return result;
      }

      // Don't retry on last attempt
      if (attempt === maxRetries) {
        return result;
      }

      // Calculate backoff delay with jitter; honor Retry-After when larger
      const baseDelay = Math.min(
        retryBaseDelayMs * Math.pow(2, attempt - 1),
        retryMaxDelayMs
      );
      const backoff = baseDelay * (0.5 + Math.random() * 0.5);
      const delay = Math.max(backoff, result.error.retryAfterMs ?? 0);
      
      await sleep(delay);
      
    } catch (error) {
      // Unexpected errors
      lastError = createError(
        ErrorCode.UNKNOWN,
        error instanceof Error ? error.message : String(error),
        { retryable: false }
      );
      
      if (attempt === maxRetries) {
        return err(lastError);
      }
    }
  }

  return err(lastError ?? createError(ErrorCode.UNKNOWN, 'Request failed'));
}

/**
 * Makes a single fetch request with timeout.
 */
async function fetchWithTimeout<T>(
  url: string,
  options: RequestInit,
  timeoutMs: number
): Promise<Result<HttpResponse<T>>> {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const response = await fetch(url, {
      ...options,
      signal: controller.signal,
    });

    clearTimeout(timeoutId);

    // Handle error responses
    if (!response.ok) {
      const retryAfterMs = extractRetryAfter(response.headers);
      const errorText = await response.text().catch(() => '');
      
      return err(createError(
        classifyHttpError(response.status, errorText),
        `HTTP ${response.status}: ${errorText || response.statusText}`,
        { retryAfterMs }
      ));
    }

    // Parse JSON response
    const contentType = response.headers.get('content-type') || '';
    let data: T;
    
    if (contentType.includes('application/json')) {
      const text = await response.text();
      data = text ? JSON.parse(text) as T : {} as T;
    } else {
      data = await response.text() as unknown as T;
    }

    return ok({
      status: response.status,
      headers: response.headers,
      data,
    });

  } catch (error) {
    clearTimeout(timeoutId);

    if (error instanceof Error) {
      if (error.name === 'AbortError') {
        return err(createError(
          ErrorCode.TIMEOUT,
          `Request timed out after ${timeoutMs}ms`,
          { retryable: true }
        ));
      }

      if (error.message.includes('ECONNRESET') || 
          error.message.includes('ETIMEDOUT') ||
          error.message.includes('ENOTFOUND')) {
        return err(createError(
          ErrorCode.NETWORK_ERROR,
          error.message,
          { retryable: true }
        ));
      }
    }

    return err(createError(
      ErrorCode.UNKNOWN,
      error instanceof Error ? error.message : String(error),
      { retryable: false }
    ));
  }
}

/**
 * Sleep for a specified duration.
 */
function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}
