/**
 * Approximate text measurement for layout decisions (row heights, pagination).
 * PowerPoint does its own final layout, so these estimates err on the wide side.
 * Widths are in em, tuned for geometric sans fonts such as Poppins / Inter.
 */
const NARROW = new Set("il.,:;'|!()[]{}/\\`ijtfr");
const WIDE = new Set("mwMW@%&");

export function charWidthEm(ch: string, widthScale = 1): number {
  let w: number;
  if (ch === " ") w = 0.27;
  else if (NARROW.has(ch)) w = 0.3;
  else if (WIDE.has(ch)) w = 0.88;
  else if (/[A-Z0-9]/.test(ch)) w = 0.66;
  else w = 0.56;
  return w * widthScale;
}

export function textWidth(text: string, fontSize: number, widthScale = 1): number {
  let w = 0;
  for (const ch of text) w += charWidthEm(ch, widthScale);
  return w * fontSize;
}

/** Greedy word-wrap; returns the wrapped lines. */
export function wrapText(text: string, fontSize: number, maxWidth: number, widthScale = 1): string[] {
  const lines: string[] = [];
  for (const para of text.split("\n")) {
    let line = "";
    for (const word of para.split(/\s+/).filter(Boolean)) {
      const candidate = line ? `${line} ${word}` : word;
      if (line && textWidth(candidate, fontSize, widthScale) > maxWidth) {
        lines.push(line);
        line = word;
      } else line = candidate;
    }
    lines.push(line);
  }
  return lines;
}

/** Height in px of `text` wrapped into `maxWidth`. */
export function textHeight(
  text: string,
  opts: { size: number; width: number; lineHeight?: number; widthScale?: number },
): number {
  const lines = wrapText(text, opts.size, opts.width, opts.widthScale).length;
  return lines * opts.size * (opts.lineHeight ?? 1.2);
}

/** Truncate to fit `maxLines`, appending an ellipsis. */
export function clampText(
  text: string,
  opts: { size: number; width: number; maxLines: number; widthScale?: number },
): string {
  const lines = wrapText(text, opts.size, opts.width, opts.widthScale);
  if (lines.length <= opts.maxLines) return text;
  const kept = lines.slice(0, opts.maxLines);
  let last = kept[kept.length - 1];
  while (last && textWidth(`${last}…`, opts.size, opts.widthScale) > opts.width) last = last.slice(0, -1);
  kept[kept.length - 1] = `${last.trimEnd()}…`;
  return kept.join(" ");
}

/** Split items into pages so that the summed heights fit `available`. */
export function paginate<T>(items: T[], heightOf: (item: T) => number, available: number): T[][] {
  const pages: T[][] = [];
  let page: T[] = [];
  let used = 0;
  for (const item of items) {
    const h = heightOf(item);
    if (page.length && used + h > available) {
      pages.push(page);
      page = [];
      used = 0;
    }
    page.push(item);
    used += h;
  }
  if (page.length || !pages.length) pages.push(page);
  return pages;
}
