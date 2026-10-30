import type { Child } from "./jsx-runtime.js";
import { flatten, hex } from "./render.js";
import type {
  ChartProps,
  Fill,
  ImageProps,
  LineProps,
  PathProps,
  PresentationProps,
  RingProps,
  ShapeProps,
  Shadow,
  SlideNumberProps,
  SlideProps,
  SpanProps,
  Stroke,
  TextProps,
  TextStyle,
} from "./types.js";

type HostNode = ReturnType<typeof flatten>[number];

export interface SvgOptions {
  /**
   * Map PowerPoint font names to CSS family + weight, e.g.
   * `{ "Poppins Medium": { family: "Poppins", weight: 500 } }`.
   */
  fontMap?: Record<string, { family: string; weight?: number }>;
}

export interface SvgSlide {
  svg: string;
  notes?: string;
}

/**
 * Render a <presentation> to one SVG string per slide, for previews in the browser
 * (Storybook, a React app, a docs page). Geometry matches the .pptx output; text is laid
 * out by the browser, so wrapping may differ slightly from PowerPoint.
 */
export function renderToSvg(tree: Child, options: SvgOptions = {}): SvgSlide[] {
  const roots = flatten(tree);
  const root = roots.length === 1 && roots[0].type === "presentation" ? roots[0] : null;
  if (!root) throw new Error("renderToSvg() expects a single <presentation> root");
  const pres = root.props as PresentationProps;
  const W = pres.width ?? 1280;
  const H = pres.height ?? 720;
  const slides = flatten(root.props.children).filter((n) => n.type === "slide");
  return slides.map((node, i) => {
    const s = new SvgSlideWriter(W, H, i + 1, pres, options);
    return { svg: s.slide(node.props as SlideProps), notes: node.props.notes };
  });
}

let uid = 0;

class SvgSlideWriter {
  private defs: string[] = [];
  private id = `pjx${(++uid).toString(36)}`;
  private n = 0;

  constructor(
    private W: number,
    private H: number,
    private page: number,
    private pres: PresentationProps,
    private options: SvgOptions,
  ) {}

  private family(name: string) {
    const m = this.options.fontMap?.[name];
    return m ? `font-family:'${m.family}',sans-serif;${m.weight ? `font-weight:${m.weight};` : ""}` : `font-family:'${name}',sans-serif;`;
  }

  slide(p: SlideProps) {
    const body: string[] = [];
    const bg = p.background;
    if (bg && typeof bg === "object" && "image" in bg) {
      body.push(`<image href="${attr(bg.image)}" x="0" y="0" width="${this.W}" height="${this.H}" preserveAspectRatio="xMidYMid slice"/>`);
    } else {
      body.push(`<rect x="0" y="0" width="${this.W}" height="${this.H}" fill="${this.paint(bg ?? "FFFFFF")}"/>`);
    }
    for (const c of flatten(p.children)) body.push(this.el(c));
    const font = this.pres.font ? this.family(this.pres.font) : "font-family:sans-serif;";
    return (
      `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${this.W} ${this.H}" style="${font}display:block;width:100%;height:auto">` +
      `<defs>${this.defs.join("")}</defs>${body.join("")}</svg>`
    );
  }

  private el(node: HostNode): string {
    const p = node.props;
    switch (node.type) {
      case "text": return this.text(p as TextProps);
      case "shape": return this.shape(p as ShapeProps);
      case "path": return this.path(p as PathProps);
      case "line": return this.line(p as LineProps);
      case "image": return this.image(p as ImageProps);
      case "chart": return this.chart(p as ChartProps);
      case "ring": return this.ring(p as RingProps);
      case "slideNumber": return this.slideNumber(p as SlideNumberProps);
      case "group": return flatten(p.children).map((c) => this.el(c)).join("");
      default: return "";
    }
  }

  // ---- paint -------------------------------------------------------------

  private paint(fill: Fill | undefined): string {
    if (fill == null) return "none";
    if (typeof fill === "string") return `#${hex(fill)}`;
    if ("color" in fill) return `#${hex(fill.color)}`;
    const gid = `${this.id}g${++this.n}`;
    const stops = fill.stops
      .map((s) => `<stop offset="${s.pos}%" stop-color="#${hex(s.color)}" stop-opacity="${1 - (s.transparency ?? 0) / 100}"/>`)
      .join("");
    if (fill.type === "radial") {
      this.defs.push(`<radialGradient id="${gid}">${stops}</radialGradient>`);
    } else {
      const a = (fill.angle * Math.PI) / 180;
      const dx = Math.sin(a) / 2, dy = -Math.cos(a) / 2;
      this.defs.push(
        `<linearGradient id="${gid}" x1="${0.5 - dx}" y1="${0.5 - dy}" x2="${0.5 + dx}" y2="${0.5 + dy}">${stops}</linearGradient>`,
      );
    }
    return `url(#${gid})`;
  }

  private opacity(fill: Fill | undefined) {
    return fill && typeof fill === "object" && "color" in fill && fill.transparency
      ? ` fill-opacity="${1 - fill.transparency / 100}"`
      : "";
  }

  private stroke(s: Stroke | undefined) {
    if (!s) return "";
    const w = s.width ?? 1;
    const dash = s.dash && s.dash !== "solid" ? ` stroke-dasharray="${w * 4} ${w * 3}"` : "";
    return ` stroke="#${hex(s.color)}" stroke-width="${w}"${dash}`;
  }

  private shadow(s: Shadow | undefined) {
    if (!s) return "";
    const off = s.offset ?? 4;
    const a = ((s.angle ?? 90) * Math.PI) / 180;
    const c = hex(s.color ?? "000000");
    const o = s.opacity ?? 0.18;
    const rgba = `rgba(${parseInt(c.slice(0, 2), 16)},${parseInt(c.slice(2, 4), 16)},${parseInt(c.slice(4, 6), 16)},${o})`;
    return ` style="filter:drop-shadow(${(Math.cos(a) * off).toFixed(1)}px ${(Math.sin(a) * off).toFixed(1)}px ${(s.blur ?? 12) / 2}px ${rgba})"`;
  }

  // ---- elements ----------------------------------------------------------

  private css(s: TextStyle) {
    let c = "";
    if (s.font) c += this.family(s.font);
    if (s.size != null) c += `font-size:${s.size}px;`;
    if (s.color) c += `color:#${hex(s.color)};`;
    if (s.bold) c += "font-weight:700;";
    if (s.italic) c += "font-style:italic;";
    if (s.underline) c += "text-decoration:underline;";
    if (s.tracking != null) c += `letter-spacing:${s.tracking}px;`;
    return c;
  }

  private text(p: TextProps) {
    const parts = flatten(p.children).map((c) => {
      if (c.type === "#text") return esc(c.props.value);
      const sp = c.props as SpanProps;
      const t = flatten(sp.children).map((n) => (n.type === "#text" ? n.props.value : "")).join("");
      return `<span style="${this.css(sp)}">${esc(t)}</span>${sp.br ? "<br/>" : ""}`;
    });
    const justify = { top: "flex-start", middle: "center", bottom: "flex-end" }[p.valign ?? "top"];
    const base = this.css({ color: this.pres.color, ...p, font: p.font });
    const bg = p.fill ? `<rect x="${p.x}" y="${p.y}" width="${p.w}" height="${p.h}" fill="${this.paint(p.fill)}"/>` : "";
    const rot = p.rotate ? ` transform="rotate(${p.rotate} ${p.x + p.w / 2} ${p.y + p.h / 2})"` : "";
    return (
      `${bg}<foreignObject x="${p.x}" y="${p.y}" width="${p.w}" height="${p.h}"${rot}${p.wrap === false ? ' overflow="visible"' : ""}>` +
      `<div xmlns="http://www.w3.org/1999/xhtml" style="box-sizing:border-box;width:100%;height:100%;display:flex;flex-direction:column;justify-content:${justify};` +
      `padding:${p.inset ?? 0}px;overflow:${p.wrap === false ? "visible" : "hidden"};font-size:18px;line-height:${p.lineHeight ?? 1.2};text-align:${p.align ?? "left"};` +
      `white-space:${p.wrap === false ? "nowrap" : "pre-wrap"};overflow-wrap:break-word;${base}"><div>${parts.join("")}</div></div></foreignObject>`
    );
  }

  private shape(p: ShapeProps) {
    const kind = p.kind ?? (p.radius ? "roundRect" : "rect");
    const { x, y, w, h } = p;
    const paint = ` fill="${this.paint(p.fill)}"${this.opacity(p.fill)}${this.stroke(p.stroke)}${this.shadow(p.shadow)}`;
    const t: string[] = [];
    if (p.rotate) t.push(`rotate(${p.rotate} ${x + w / 2} ${y + h / 2})`);
    if (p.flipH || p.flipV) {
      t.push(`translate(${x + w / 2} ${y + h / 2}) scale(${p.flipH ? -1 : 1} ${p.flipV ? -1 : 1}) translate(${-(x + w / 2)} ${-(y + h / 2)})`);
    }
    const tr = t.length ? ` transform="${t.join(" ")}"` : "";
    const poly = (pts: [number, number][]) =>
      `<polygon points="${pts.map(([px, py]) => `${x + px * w},${y + py * h}`).join(" ")}"${paint}${tr}/>`;
    switch (kind) {
      case "ellipse":
        return `<ellipse cx="${x + w / 2}" cy="${y + h / 2}" rx="${w / 2}" ry="${h / 2}"${paint}${tr}/>`;
      case "roundRect": {
        const r = Math.min(p.radius ?? Math.min(w, h) * 0.1667, w / 2, h / 2);
        return `<rect x="${x}" y="${y}" width="${w}" height="${h}" rx="${r}"${paint}${tr}/>`;
      }
      case "diamond": return poly([[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]]);
      case "triangle": return poly([[0.5, 0], [1, 1], [0, 1]]);
      case "chevron": return poly([[0, 0], [0.75, 0], [1, 0.5], [0.75, 1], [0, 1], [0.25, 0.5]]);
      case "pentagon": return poly([[0, 0], [0.75, 0], [1, 0.5], [0.75, 1], [0, 1]]);
      case "hexagon": return poly([[0.25, 0], [0.75, 0], [1, 0.5], [0.75, 1], [0.25, 1], [0, 0.5]]);
      case "rightArrow": return poly([[0, 0.25], [0.6, 0.25], [0.6, 0], [1, 0.5], [0.6, 1], [0.6, 0.75], [0, 0.75]]);
      case "star5": {
        const pts: [number, number][] = [];
        for (let i = 0; i < 10; i++) {
          const r = i % 2 ? 0.2 : 0.5;
          const a = -Math.PI / 2 + (i * Math.PI) / 5;
          pts.push([0.5 + r * Math.cos(a), 0.5 + r * Math.sin(a)]);
        }
        return poly(pts);
      }
      default:
        return `<rect x="${x}" y="${y}" width="${w}" height="${h}"${paint}${tr}/>`;
    }
  }

  private path(p: PathProps) {
    return `<path d="${attr(p.d)}" fill="${this.paint(p.fill)}"${this.opacity(p.fill)}${this.stroke(p.stroke)}${this.shadow(p.shadow)}/>`;
  }

  private line(p: LineProps) {
    return `<line x1="${p.x1}" y1="${p.y1}" x2="${p.x2}" y2="${p.y2}"${this.stroke({ color: p.color ?? "000000", width: p.width, dash: p.dash })}/>`;
  }

  private image(p: ImageProps) {
    const href = p.data ?? p.src ?? "";
    const par = p.fit === "cover" ? "xMidYMid slice" : p.fit === "contain" ? "xMidYMid meet" : "none";
    return `<image href="${attr(href)}" x="${p.x}" y="${p.y}" width="${p.w}" height="${p.h}" preserveAspectRatio="${par}"${p.transparency ? ` opacity="${1 - p.transparency / 100}"` : ""}/>`;
  }

  private ring(p: RingProps) {
    const t = p.thickness ?? Math.max(2, p.r * 0.22);
    const rm = p.r - t / 2;
    const v = Math.max(0, Math.min(1, p.value));
    const circ = 2 * Math.PI * rm;
    const track = v < 1 ? `<circle cx="${p.cx}" cy="${p.cy}" r="${rm}" fill="none" stroke="#${hex(p.track ?? "E5E7EB")}" stroke-width="${t}"/>` : "";
    const val = v > 0
      ? `<circle cx="${p.cx}" cy="${p.cy}" r="${rm}" fill="none" stroke="#${hex(p.color)}" stroke-width="${t}" stroke-dasharray="${circ * v} ${circ}" transform="rotate(-90 ${p.cx} ${p.cy})"/>`
      : "";
    return track + val;
  }

  private slideNumber(p: SlideNumberProps) {
    return this.text({ ...p, align: p.align ?? "center", valign: p.valign ?? "middle", children: String(this.page) });
  }

  private chart(p: ChartProps) {
    const colors = (p.colors ?? ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47"]).map((c) => `#${hex(c)}`);
    const series = p.data[0];
    if (!series) return "";
    if (p.type === "doughnut" || p.type === "pie") {
      const total = series.values.reduce((a, b) => a + b, 0) || 1;
      const cx = p.x + p.w / 2, cy = p.y + p.h / 2;
      const R = Math.min(p.w, p.h) / 2;
      const r = p.type === "doughnut" ? (R * (p.hole ?? 50)) / 100 : 0;
      let a0 = -Math.PI / 2;
      return series.values
        .map((v, i) => {
          const a1 = a0 + (v / total) * Math.PI * 2;
          const d = arcPath(cx, cy, R, r, a0, Math.min(a1, a0 + Math.PI * 2 - 1e-4));
          a0 = a1;
          return v > 0 ? `<path d="${d}" fill="${colors[i % colors.length]}" stroke="#FFFFFF" stroke-width="1"/>` : "";
        })
        .join("");
    }
    // Simple bar / line / area preview (first series drives the categories)
    const pad = 28;
    const max = Math.max(...p.data.flatMap((s) => s.values), 0) || 1;
    const n = series.labels.length;
    const plotX = p.x + pad, plotY = p.y + 8, plotW = p.w - pad - 8, plotH = p.h - pad - 8;
    const out: string[] = [`<line x1="${plotX}" y1="${plotY + plotH}" x2="${plotX + plotW}" y2="${plotY + plotH}" stroke="#D1D5DB"/>`];
    series.labels.forEach((l, i) => {
      out.push(`<text x="${plotX + (i + 0.5) * (plotW / n)}" y="${plotY + plotH + 16}" font-size="11" fill="#6B7280" text-anchor="middle">${esc(l)}</text>`);
    });
    p.data.forEach((s, si) => {
      const color = colors[si % colors.length];
      if (p.type === "bar") {
        const horiz = p.direction === "bar";
        const groupW = (horiz ? plotH : plotW) / n;
        const bw = (groupW * 0.7) / p.data.length;
        s.values.forEach((v, i) => {
          const len = (v / max) * (horiz ? plotW : plotH);
          out.push(horiz
            ? `<rect x="${plotX}" y="${plotY + i * groupW + groupW * 0.15 + si * bw}" width="${len}" height="${bw}" fill="${color}"/>`
            : `<rect x="${plotX + i * groupW + groupW * 0.15 + si * bw}" y="${plotY + plotH - len}" width="${bw}" height="${len}" fill="${color}"/>`);
        });
      } else {
        const pts = s.values.map((v, i) => `${plotX + (i + 0.5) * (plotW / n)},${plotY + plotH - (v / max) * plotH}`);
        if (p.type === "area") {
          out.push(`<polygon points="${plotX + 0.5 * (plotW / n)},${plotY + plotH} ${pts.join(" ")} ${plotX + (n - 0.5) * (plotW / n)},${plotY + plotH}" fill="${color}" fill-opacity="0.35"/>`);
        }
        out.push(`<polyline points="${pts.join(" ")}" fill="none" stroke="${color}" stroke-width="2.5"/>`);
      }
    });
    return out.join("");
  }
}

function arcPath(cx: number, cy: number, R: number, r: number, a0: number, a1: number) {
  const large = a1 - a0 > Math.PI ? 1 : 0;
  const p = (rad: number, a: number) => `${(cx + rad * Math.cos(a)).toFixed(2)} ${(cy + rad * Math.sin(a)).toFixed(2)}`;
  if (r <= 0) return `M${cx} ${cy} L${p(R, a0)} A${R} ${R} 0 ${large} 1 ${p(R, a1)} Z`;
  return `M${p(R, a0)} A${R} ${R} 0 ${large} 1 ${p(R, a1)} L${p(r, a1)} A${r} ${r} 0 ${large} 0 ${p(r, a0)} Z`;
}

function esc(s: string) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function attr(s: string) {
  return esc(s).replace(/"/g, "&quot;");
}
