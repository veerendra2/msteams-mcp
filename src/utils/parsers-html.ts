/**
 * HTML parsing and escaping utilities.
 */

import type { ExtractedLink } from '../types/teams.js';

// Re-export ExtractedLink so existing imports from parsers.ts continue to work
export type { ExtractedLink };

/**
 * Extracts links from HTML content before stripping.
 * Returns an array of { url, text, contentType? } objects.
 *
 * Parses both:
 * - <a href="...">text</a> - standard HTML links
 * - <item type="..." uri="..."> - Teams recording/transcript URIs (amsTranscript, onedriveForBusinessVideo, etc.)
 *   Handles type and uri in either order.
 */
export function extractLinks(html: string): ExtractedLink[] {
  const links: ExtractedLink[] = [];

  // <a href="...">text</a> - standard HTML links
  for (const m of html.matchAll(/<a\s+[^>]*href=["'](?<url>[^"']+)["'][^>]*>(?<text>[\s\S]*?)<\/a>/gi)) {
    const { url, text } = m.groups!;
    if (url && !url.startsWith('javascript:')) {
      links.push({ url, text: stripHtmlInline(text) || url });
    }
  }

  // <item type="..." uri="..."> - Teams recording/transcript URIs (attribute order may vary)
  const itemTypeFirst = /<item\s+[^>]*type=["'](?<type>[^"']+)["'][^>]*uri=["'](?<uri>[^"']+)["'][^>]*>/gi;
  const itemUriFirst = /<item\s+[^>]*uri=["'](?<uri>[^"']+)["'][^>]*type=["'](?<type>[^"']+)["'][^>]*>/gi;
  const seen = new Set<string>();
  for (const re of [itemTypeFirst, itemUriFirst]) {
    for (const m of html.matchAll(re)) {
      const { type, uri } = m.groups!;
      if (uri && !seen.has(uri)) {
        seen.add(uri);
        links.push({ url: uri, text: type, contentType: type });
      }
    }
  }

  return links;
}

/**
 * Block-level tag boundaries mapped to the text break they represent.
 *
 * Teams sends message bodies as `RichText/Html`, so a multi-paragraph message
 * arrives as `<p>…</p>\r\n<p>…</p>` and a line break arrives as `<br>`. Those
 * tags are the only record of the author's line structure, so they are turned
 * into real newlines before the remaining tags are dropped. Empty spacer
 * paragraphs (`<p>&nbsp;</p>`, which the Teams client inserts between
 * paragraphs) collapse away with the rest of the blank-line runs.
 *
 * Order matters only in that closing tags are listed before their openers for
 * readability; the patterns are mutually exclusive.
 */
const BLOCK_BREAKS: [RegExp, string][] = [
  // Teams emits `<br>\r\n`, so the source's own newline is swallowed here —
  // otherwise a soft line break would read back as a blank line.
  [/<br\s*\/?>[ \t]*\n?/gi, '\n'],
  // Paragraph-like blocks get a blank line so paragraphs stay distinguishable
  // from the single line breaks a <br> produces.
  [/<\/(?:p|h[1-6]|blockquote|pre|table|ul|ol)\s*>/gi, '\n\n'],
  [/<(?:p|h[1-6]|blockquote|pre|table|ul|ol)(?:\s[^>]*)?>/gi, '\n'],
  [/<\/(?:div|tr)\s*>/gi, '\n'],
  [/<li(?:\s[^>]*)?>/gi, '\n- '],
  // Cell boundaries would otherwise run adjacent cells together.
  [/<\/(?:td|th)\s*>/gi, ' | '],
];

/** HTML entities, `&amp;` last so `&amp;lt;` decodes to `&lt;` and not `<`. */
const ENTITIES: [RegExp, string][] = [
  [/&nbsp;/gi, ' '],
  [/&lt;/gi, '<'],
  [/&gt;/gi, '>'],
  [/&quot;/gi, '"'],
  [/&#0*39;/g, "'"],
  [/&apos;/gi, "'"],
  [/&amp;/gi, '&'],
];

/**
 * Converts message HTML to plain text, preserving the author's line structure.
 *
 * Paragraphs, `<br>` breaks, list items and table rows survive as newlines;
 * inline tags are removed without inserting a space, so `un<b>break</b>able`
 * stays one word. Horizontal whitespace is collapsed and blank-line runs are
 * capped at one, but leading indentation is kept so fenced code blocks remain
 * readable.
 *
 * Use {@link stripHtmlInline} where the result has to occupy a single line.
 */
export function stripHtml(html: string): string {
  let text = html.replace(/\r\n?/g, '\n');

  for (const [pattern, replacement] of BLOCK_BREAKS) {
    text = text.replace(pattern, replacement);
  }

  // A row's final cell leaves a separator with nothing after it. Matching the
  // trailing space keeps a pipe the author typed themselves (`| a | b |`).
  text = text.replace(/ \| (?=\n|$)/g, '');

  // Remaining tags are inline (b, i, a, code, span, …) and carry no structure.
  text = text.replace(/<[^>]*>/g, '');

  for (const [pattern, replacement] of ENTITIES) {
    text = text.replace(pattern, replacement);
  }

  const lines = text.split('\n').map(line => {
    const indent = /^[ \t]*/.exec(line)![0];
    const body = line
      .slice(indent.length)
      .replace(/[^\S\n]+/g, ' ')
      .trimEnd();
    return body ? indent + body : '';
  });

  return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}

/**
 * Converts message HTML to a single line of plain text.
 *
 * For previews and labels — chat-list summaries, link text — where a newline
 * would break the surrounding layout.
 */
export function stripHtmlInline(html: string): string {
  return stripHtml(html).replace(/\s+/g, ' ').trim();
}

/**
 * Escapes HTML special characters in text.
 */
export function escapeHtmlChars(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Extracts the lowercased URI scheme from a URL, or undefined if the URL is
 * relative (RFC 3986 §3.1).
 *
 * A scheme is the run of letters/digits/`+`/`-`/`.` that starts with a letter
 * and precedes the first `:` — but only when that `:` comes before any `/`, `?`
 * or `#`, so `path/to:thing`, `foo?a=b:c` and `#a:b` are correctly treated as
 * scheme-less relative URLs.
 */
function urlScheme(url: string): string | undefined {
  const colon = url.indexOf(':');
  if (colon <= 0) return undefined;
  const scheme = url.slice(0, colon);
  if (!/^[a-zA-Z][a-zA-Z0-9+.-]*$/.test(scheme)) return undefined;
  // A `/`, `?` or `#` before the colon means the colon is inside a relative URL.
  if (/[/?#]/.test(url.slice(0, colon))) return undefined;
  return scheme.toLowerCase();
}

/**
 * Neutralises a link destination whose URL scheme is not on the allowlist.
 *
 * Only `http`, `https` and `mailto` (case-insensitive) are kept; any other
 * scheme (`javascript:`, `data:`, `vbscript:`, `file:`, …) is rewritten to `#`
 * so a link can never carry an executable scheme into a Teams-rendered message.
 * Scheme-less URLs — relative paths and fragment-only links — carry no
 * executable scheme and pass through unchanged.
 */
export function sanitizeLinkUrl(url: string): string {
  const scheme = urlScheme(url);
  if (scheme !== undefined && !['http', 'https', 'mailto'].includes(scheme)) {
    return '#';
  }
  return url;
}
