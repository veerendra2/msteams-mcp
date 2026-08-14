/**
 * Adaptive Card decoding for Teams SWIFT.1 / RichText/Media_Card messages.
 *
 * Workflows and similar bots send the card as base64 inside `<Swift b64="…">`.
 * `stripHtml()` would otherwise leave only the "cards.unsupported" placeholder.
 */

import type { ExtractedLink } from '../types/teams.js';

/** Nested Adaptive Card node (body items, columns, facts, actions). */
interface AdaptiveCardNode {
  type?: string;
  text?: string;
  title?: string;
  value?: string;
  url?: string;
  facts?: Array<{ title?: string; value?: string }>;
  items?: AdaptiveCardNode[];
  columns?: Array<{ items?: AdaptiveCardNode[] }>;
  body?: AdaptiveCardNode[];
  actions?: Array<{ title?: string; url?: string; type?: string }>;
}

/** SWIFT payload wrapping one or more Adaptive Card attachments. */
interface SwiftPayload {
  attachments?: Array<{
    contentType?: string;
    content?: AdaptiveCardNode;
  }>;
}

/** Placeholder link Teams inserts when the client cannot render the card. */
const UNSUPPORTED_CARD_URL = 'go.skype.com/cards.unsupported';

/**
 * Pulls the base64 SWIFT payload out of chatsvc HTML, if present.
 */
function extractSwiftB64(html: string): string | undefined {
  const match = /<Swift\b[^>]*\bb64="([^"]+)"/i.exec(html);
  const value = match?.[1];
  return value && value.length > 0 ? value : undefined;
}

/**
 * Recursively flattens Adaptive Card nodes into readable lines and action links.
 */
function flattenCard(
  node: AdaptiveCardNode | undefined,
  lines: string[],
  links: ExtractedLink[]
): void {
  if (!node) return;

  if (typeof node.text === 'string') {
    const text = node.text.trim();
    if (text) lines.push(text);
  }

  if (Array.isArray(node.facts)) {
    for (const fact of node.facts) {
      const title = fact.title?.trim() ?? '';
      const value = fact.value?.trim() ?? '';
      if (title || value) {
        lines.push(`${title}: ${value}`.trim());
      }
    }
  }

  for (const item of node.items ?? []) {
    flattenCard(item, lines, links);
  }
  for (const column of node.columns ?? []) {
    for (const item of column.items ?? []) {
      flattenCard(item, lines, links);
    }
  }
  for (const child of node.body ?? []) {
    flattenCard(child, lines, links);
  }

  if (Array.isArray(node.actions)) {
    for (const action of node.actions) {
      const title = action.title?.trim() ?? '';
      const url = action.url?.trim() ?? '';
      if (!title && !url) continue;
      if (url) {
        links.push({ url, text: title || url });
        lines.push(title ? `${title}: ${url}` : url);
      } else if (title) {
        lines.push(title);
      }
    }
  }
}

/**
 * Decodes a SWIFT Adaptive Card embedded in chatsvc HTML.
 *
 * @returns Flattened text and action links, or null when the HTML is not a SWIFT card
 *          or the payload cannot be parsed (caller should fall back to HTML stripping).
 */
export function decodeSwiftAdaptiveCard(
  html: string
): { text: string; links: ExtractedLink[] } | null {
  const b64 = extractSwiftB64(html);
  if (!b64) return null;

  let payload: SwiftPayload;
  try {
    payload = JSON.parse(Buffer.from(b64, 'base64').toString('utf8')) as SwiftPayload;
  } catch {
    return null;
  }

  const lines: string[] = [];
  const links: ExtractedLink[] = [];
  for (const attachment of payload.attachments ?? []) {
    flattenCard(attachment.content, lines, links);
  }

  if (lines.length === 0 && links.length === 0) return null;
  return { text: lines.join('\n'), links };
}

/**
 * True when a URL is the Teams "card not supported" placeholder, not a real action.
 */
export function isUnsupportedCardPlaceholderUrl(url: string): boolean {
  return url.includes(UNSUPPORTED_CARD_URL);
}
