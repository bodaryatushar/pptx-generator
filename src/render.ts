import PptxGenJSImport from "pptxgenjs";
import JSZip from "jszip";
import { Fragment, type Child, type PNode } from "./jsx-runtime.js";
import { parsePath } from "./path.js";
import type {
  ChartProps,
  Color,
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
} from "./types.js";

// pptxgenjs ships CJS with a default export; normalise across loaders.
const PptxGenJS: new () => Pptx = (PptxGenJSImport as any).default ?? (PptxGenJSImport as any);

// The pptxgenjs typings describe the instance through a namespace; we only need a few members.
type Pptx = {
  addSlide(): any;
  defineLayout(l: { name: string; width: number; height: number }): void;
  layout: string;
  title: string;
  author: string;
  company: string;
  write(o: { outputType: "uint8array" }): Promise<unknown>;
};
type PSlide = ReturnType<Pptx["addSlide"]>;

interface HostNode {
  type: string;
  props: Record<string, any>;
}

/** Expand components, fragments and arrays into a flat list of intrinsic nodes. */
export function flatten(child: Child): HostNode[] {
  if (child == null || child === false || child === true) return [];
  if (Array.isArray(child)) return child.flatMap(flatten);
  if (typeof child === "string" || typeof child === "number") {
    return [{ type: "#text", props: { value: String(child) } }];
  }
  const node = child as PNode;
  if (node.type === Fragment) return flatten(node.props.children);
  if (typeof node.type === "function") return flatten(node.type(node.props));
  return [{ type: node.type as string, props: node.props }];
}

const SLIDE_WIDTH_IN = 13.333;

class Renderer {
  readonly pptx: Pptx;
  readonly k: number; // inches per design px
  private gradients = new Map<string, Exclude<Fill, string | { color: Color }>>();
  private seq = 0;
  private defaults: { font?: string; color?: Color };
  readonly width: number;
  readonly height: number;

  constructor(props: PresentationProps) {
    const w = props.width ?? 1280;
    const h = props.height ?? 720;
    this.width = w;
    this.height = h;
    this.k = SLIDE_WIDTH_IN / w;
    this.pptx = new PptxGenJS();
    this.pptx.defineLayout({ name: "PPTX_JSX", width: SLIDE_WIDTH_IN, height: h * this.k });
    this.pptx.layout = "PPTX_JSX";
    if (props.title) this.pptx.title = props.title;
    if (props.author) this.pptx.author = props.author;
    if (props.company) this.pptx.company = props.company;
    this.defaults = { font: props.font, color: props.color };
  }

  in = (px: number) => px * this.k;
  pt = (px: number) => +(px * 0.75).toFixed(2);

  private nextName(prefix: string) {
    return `${prefix} ${++this.seq}`;
  }

  // ---- fills -------------------------------------------------------------

  private fill(fill: Fill | undefined, name: string) {
    if (fill == null) return { type: "none" as const };
    if (typeof fill === "string") return { color: hex(fill) };
    if ("color" in fill) return { color: hex(fill.color), transparency: fill.transparency };
    // Gradient: emit a placeholder solid fill and patch the XML after generation.
    this.gradients.set(name, fill);
    return { color: hex(fill.stops[0]?.color ?? "000000") };
  }

  private stroke(s: Stroke | undefined) {
    if (!s) return undefined;
    return { color: hex(s.color), width: this.pt(s.width ?? 1), dashType: s.dash };
  }

  private shadow(s: Shadow | undefined) {
    if (!s) return undefined;
    return {
      type: "outer" as const,
      color: hex(s.color ?? "000000"),
      blur: this.pt(s.blur ?? 12),
      offset: this.pt(s.offset ?? 4),
      angle: s.angle ?? 90,
      opacity: s.opacity ?? 0.18,
    };
  }

  // ---- elements ----------------------------------------------------------

  slide(props: SlideProps) {
    const slide = this.pptx.addSlide();
    const bg = props.background;
    if (bg && typeof bg === "object" && "image" in bg) {
      slide.background = { path: bg.image };
    } else if (typeof bg === "string") {
      slide.background = { color: hex(bg) };
    } else if (bg) {
      // Gradient backgrounds: a full-bleed shape behind everything.
      this.shape(slide, { x: 0, y: 0, w: this.width, h: this.height, fill: bg, name: "Background" });
    }
    if (props.notes) slide.addNotes(props.notes);
    for (const child of flatten(props.children)) this.element(slide, child);
  }

  private element(slide: PSlide, node: HostNode) {
    const p = node.props;
    switch (node.type) {
      case "text":
        return this.text(slide, p as TextProps);
      case "shape":
        return this.shape(slide, p as ShapeProps);
      case "path":
        return this.path(slide, p as PathProps);
      case "line":
        return this.line(slide, p as LineProps);
      case "image":
        return this.image(slide, p as ImageProps);
      case "chart":
        return this.chart(slide, p as ChartProps);
      case "ring":
        return this.ring(slide, p as RingProps);
      case "slideNumber":
        return this.slideNumber(slide, p as SlideNumberProps);
      case "group":
        for (const c of flatten(p.children)) this.element(slide, c);
        return;
      case "#text":
        if (p.value.trim()) throw new Error(`Bare text "${p.value}" must be inside <text>`);
        return;
      default:
        throw new Error(`<${node.type}> is not allowed inside <slide>`);
    }
  }

  private textStyle(s: SpanProps | TextProps) {
    return {
      fontFace: s.font,
      fontSize: s.size != null ? this.pt(s.size) : undefined,
      color: s.color ? hex(s.color) : undefined,
      bold: s.bold,
      italic: s.italic,
      underline: s.underline ? { style: "sng" as const } : undefined,
      charSpacing: s.tracking != null ? this.pt(s.tracking) : undefined,
    };
  }

  text(slide: PSlide, p: TextProps) {
    const name = p.name ?? this.nextName("Text");
    const runs: { text: string; options?: Record<string, unknown> }[] = [];
    for (const c of flatten(p.children)) {
      if (c.type === "#text") runs.push({ text: c.props.value });
      else if (c.type === "span") {
        const sp = c.props as SpanProps;
        const text = flatten(sp.children)
          .map((n) => (n.type === "#text" ? n.props.value : ""))
          .join("");
        runs.push({ text, options: { ...clean(this.textStyle(sp)), breakLine: sp.br } });
      } else throw new Error(`<${c.type}> is not allowed inside <text>`);
    }
    const base = clean({
      fontFace: this.defaults.font,
      color: this.defaults.color ? hex(this.defaults.color) : undefined,
      ...clean(this.textStyle(p)),
    });
    slide.addText(runs.length ? runs : "", {
      ...base,
      objectName: name,
      x: this.in(p.x),
      y: this.in(p.y),
      w: this.in(p.w),
      h: this.in(p.h),
      align: p.align ?? "left",
      valign: p.valign ?? "top",
      lineSpacingMultiple: p.lineHeight,
      margin: p.inset != null ? this.pt(p.inset) : 0,
      fill: p.fill ? (this.fill(p.fill, name) as any) : undefined,
      fit: p.fit ? "shrink" : "none",
      wrap: p.wrap ?? true,
      rotate: p.rotate,
    });
  }

  shape(slide: PSlide, p: ShapeProps) {
    const name = p.name ?? this.nextName("Shape");
    const kind = p.kind ?? (p.radius ? "roundRect" : "rect");
    slide.addShape(kind as any, {
      objectName: name,
      x: this.in(p.x),
      y: this.in(p.y),
      w: this.in(p.w),
      h: this.in(p.h),
      fill: this.fill(p.fill, name) as any,
      line: this.stroke(p.stroke) ?? { type: "none" },
      rectRadius: p.radius != null ? this.in(Math.min(p.radius, p.w / 2, p.h / 2)) : undefined,
      shadow: this.shadow(p.shadow),
      rotate: p.rotate,
      flipH: p.flipH,
      flipV: p.flipV,
    } as any);
  }

  path(slide: PSlide, p: PathProps) {
    const name = p.name ?? this.nextName("Freeform");
    const { box, points } = parsePath(p.d);
    const rel = (x: number, y: number) => ({ x: this.in(x - box.x), y: this.in(y - box.y) });
    const pts = points.map((pt) => {
      if (pt.kind === "close") return { close: true as const };
      const xy = rel(pt.x, pt.y);
      if (pt.kind === "move") return { ...xy, moveTo: true };
      if (pt.kind === "quad") {
        const c = rel(pt.x1, pt.y1);
        return { ...xy, curve: { type: "quadratic" as const, x1: c.x, y1: c.y } };
      }
      if (pt.kind !== "cubic") return xy;
      const c1 = rel(pt.x1, pt.y1);
      const c2 = rel(pt.x2, pt.y2);
      return { ...xy, curve: { type: "cubic" as const, x1: c1.x, y1: c1.y, x2: c2.x, y2: c2.y } };
    });
    slide.addShape("custGeom" as any, {
      objectName: name,
      x: this.in(box.x),
      y: this.in(box.y),
      w: this.in(box.w),
      h: this.in(box.h),
      points: pts,
      fill: this.fill(p.fill, name) as any,
      line: this.stroke(p.stroke) ?? { type: "none" },
      shadow: this.shadow(p.shadow),
    } as any);
  }

  line(slide: PSlide, p: LineProps) {
    const x = Math.min(p.x1, p.x2);
    const y = Math.min(p.y1, p.y2);
    slide.addShape("line" as any, {
      objectName: p.name ?? this.nextName("Line"),
      x: this.in(x),
      y: this.in(y),
      w: this.in(Math.abs(p.x2 - p.x1)),
      h: this.in(Math.abs(p.y2 - p.y1)),
      flipV: (p.x2 - p.x1) * (p.y2 - p.y1) < 0,
      line: { color: hex(p.color ?? "000000"), width: this.pt(p.width ?? 1), dashType: p.dash },
    } as any);
  }

  image(slide: PSlide, p: ImageProps) {
    const box = { x: this.in(p.x), y: this.in(p.y), w: this.in(p.w), h: this.in(p.h) };
    slide.addImage({
      objectName: p.name ?? this.nextName("Picture"),
      path: p.src,
      data: p.data,
      ...box,
      sizing: p.fit && p.fit !== "stretch" ? { type: p.fit, w: box.w, h: box.h } : undefined,
      rounding: p.rounding,
      transparency: p.transparency,
    } as any);
  }

  chart(slide: PSlide, p: ChartProps) {
    const colors = p.colors?.map(hex);
    slide.addChart(p.type as any, p.data as any, {
      objectName: p.name ?? this.nextName("Chart"),
      x: this.in(p.x),
      y: this.in(p.y),
      w: this.in(p.w),
      h: this.in(p.h),
      chartColors: colors,
      holeSize: p.hole,
      barDir: p.direction,
      barGrouping: p.stacked ? "stacked" : undefined,
      showLegend: !!p.legend,
      legendPos: p.legend || undefined,
      legendFontFace: this.defaults.font,
      showValue: !!p.labels,
      showPercent: false,
      showLabel: false,
      showTitle: false,
      dataBorder: { pt: 0, color: "FFFFFF" },
      firstSliceAng: 0,
      catAxisLabelFontFace: this.defaults.font,
      valAxisLabelFontFace: this.defaults.font,
      dataLabelFontFace: this.defaults.font,
      ...p.options,
    } as any);
  }

  slideNumber(slide: PSlide, p: SlideNumberProps) {
    slide.slideNumber = {
      ...clean({ fontFace: this.defaults.font, ...clean(this.textStyle(p)) }),
      x: this.in(p.x),
      y: this.in(p.y),
      w: this.in(p.w),
      h: this.in(p.h),
      align: p.align ?? "center",
      valign: p.valign ?? "middle",
      margin: 0,
    };
  }

  /** Progress ring: a full track plus a block-arc for the value, starting at 12 o'clock. */
  ring(slide: PSlide, p: RingProps) {
    const t = p.thickness ?? Math.max(2, p.r * 0.22);
    const box = { x: this.in(p.cx - p.r), y: this.in(p.cy - p.r), w: this.in(p.r * 2), h: this.in(p.r * 2) };
    const ratio = Math.min(1, t / p.r);
    const name = p.name ?? this.nextName("Progress");
    const arc = (color: Color, from: number, to: number, suffix: string) =>
      slide.addShape("blockArc" as any, {
        objectName: `${name} ${suffix}`,
        ...box,
        fill: { color: hex(color) },
        line: { type: "none" },
        angleRange: [from, to],
        arcThicknessRatio: ratio,
      } as any);
    const v = Math.max(0, Math.min(1, p.value));
    if (v < 1) arc(p.track ?? "E5E7EB", 0, 359.99, "track");
    if (v >= 0.999) arc(p.color, 0, 359.99, "value");
    else if (v > 0) arc(p.color, 270, (270 + v * 360) % 360, "value");
  }

  // ---- output ------------------------------------------------------------

  async toBytes(): Promise<Uint8Array> {
    const raw = (await this.pptx.write({ outputType: "uint8array" })) as Uint8Array;
    if (!this.gradients.size) return raw;
    const zip = await JSZip.loadAsync(raw);
    const slides = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
    for (const f of slides) {
      let xml = await zip.file(f)!.async("string");
      let changed = false;
      for (const [name, g] of this.gradients) {
        const idx = xml.indexOf(`name="${escapeXml(name)}"`);
        if (idx < 0) continue;
        const sp = xml.indexOf("<p:spPr", idx);
        const end = xml.indexOf("</p:spPr>", sp);
        const segment = xml.slice(sp, end);
        const patched = segment.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, gradientXml(g));
        xml = xml.slice(0, sp) + patched + xml.slice(end);
        changed = true;
      }
      if (changed) zip.file(f, xml);
    }
    return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }
}

function gradientXml(g: Exclude<Fill, string | { color: Color }>): string {
  const stops = g.stops
    .map((s) => {
      const alpha = s.transparency ? `<a:alpha val="${Math.round((100 - s.transparency) * 1000)}"/>` : "";
      return `<a:gs pos="${Math.round(s.pos * 1000)}"><a:srgbClr val="${hex(s.color)}">${alpha}</a:srgbClr></a:gs>`;
    })
    .join("");
  if (g.type === "radial") {
    return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>`;
  }
  // CSS angle (0 = to top, clockwise) -> OOXML angle (0 = to right, clockwise)
  const ooxml = (((g.angle - 90) % 360) + 360) % 360;
  return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:lin ang="${Math.round(ooxml * 60000)}" scaled="0"/></a:gradFill>`;
}

export function hex(c: Color) {
  return c.replace(/^#/, "").toUpperCase();
}

function escapeXml(s: string) {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function clean<T extends Record<string, unknown>>(o: T): Partial<T> {
  return Object.fromEntries(Object.entries(o).filter(([, v]) => v !== undefined)) as Partial<T>;
}

function build(tree: Child): Renderer {
  const roots = flatten(tree);
  const root = roots.length === 1 && roots[0].type === "presentation" ? roots[0] : null;
  if (!root) throw new Error("render() expects a single <presentation> root");
  const r = new Renderer(root.props as PresentationProps);
  for (const node of flatten(root.props.children)) {
    if (node.type === "slide") r.slide(node.props as SlideProps);
    else if (node.type !== "#text" || node.props.value.trim()) {
      throw new Error(`Only <slide> is allowed inside <presentation>, got <${node.type}>`);
    }
  }
  return r;
}

export const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

/** Render a JSX tree to .pptx bytes. Works in Node and in the browser. */
export async function renderToBuffer(tree: Child): Promise<Uint8Array> {
  return build(tree).toBytes();
}

/** Browser: render to a Blob, e.g. for a download link. */
export async function renderToBlob(tree: Child): Promise<Blob> {
  return new Blob([await renderToBuffer(tree)], { type: PPTX_MIME });
}

/** Node: render and write to disk. */
export async function renderToFile(tree: Child, file: string): Promise<string> {
  const { writeFile } = await import("node:fs/promises");
  await writeFile(file, await renderToBuffer(tree));
  return file;
}
