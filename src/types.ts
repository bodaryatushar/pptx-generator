import type { Child } from "./jsx-runtime.js";


/** Hex color without "#", e.g. "24074A". A leading "#" is accepted and stripped. */
export type Color = string;

export interface GradientStop {
  /** 0–100 */
  pos: number;
  color: Color;
  /** 0–100, default 0 */
  transparency?: number;
}

export interface LinearGradient {
  type: "linear";
  /** CSS-style angle in degrees: 0 = to top, 90 = to right, 45 = to top-right. */
  angle: number;
  stops: GradientStop[];
}

export interface RadialGradient {
  type: "radial";
  stops: GradientStop[];
}

export type Fill = Color | LinearGradient | RadialGradient | { color: Color; transparency?: number };

export interface Stroke {
  color: Color;
  /** px */
  width?: number;
  dash?: "solid" | "dash" | "dashDot" | "lgDash" | "sysDash" | "sysDot";
}

export interface Shadow {
  color?: Color;
  /** px */
  blur?: number;
  /** px */
  offset?: number;
  /** degrees */
  angle?: number;
  /** 0–1 */
  opacity?: number;
}

/** All geometry is in design px (default canvas 1280×720). Font sizes are px too. */
export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface TextStyle {
  font?: string;
  size?: number;
  color?: Color;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  /** Letter spacing in px */
  tracking?: number;
}

export interface TextProps extends Box, TextStyle {
  align?: "left" | "center" | "right" | "justify";
  valign?: "top" | "middle" | "bottom";
  /** Line height multiplier, e.g. 1.4 */
  lineHeight?: number;
  /** Inner padding in px (default 0) */
  inset?: number;
  fill?: Fill;
  /** Shrink text to fit the box when PowerPoint opens it */
  fit?: boolean;
  wrap?: boolean;
  rotate?: number;
  name?: string;
  children?: Child;
}

export interface SpanProps extends TextStyle {
  /** Break the line after this run */
  br?: boolean;
  children?: Child;
}

export type ShapeKind =
  | "rect"
  | "roundRect"
  | "ellipse"
  | "triangle"
  | "diamond"
  | "chevron"
  | "pentagon"
  | "hexagon"
  | "rightArrow"
  | "star5";

export interface ShapeProps extends Box {
  kind?: ShapeKind;
  fill?: Fill;
  stroke?: Stroke;
  /** Corner radius in px (roundRect only) */
  radius?: number;
  shadow?: Shadow;
  rotate?: number;
  flipH?: boolean;
  flipV?: boolean;
  name?: string;
}

export interface PathProps {
  /** SVG path data in slide px, absolute commands M L H V C Q Z */
  d: string;
  fill?: Fill;
  stroke?: Stroke;
  shadow?: Shadow;
  name?: string;
}

export interface LineProps {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
  color?: Color;
  width?: number;
  dash?: Stroke["dash"];
  name?: string;
}

export interface ImageProps extends Box {
  /** File path or URL */
  src?: string;
  /** base64 data URI */
  data?: string;
  fit?: "cover" | "contain" | "stretch";
  rounding?: boolean;
  transparency?: number;
  name?: string;
}

export interface ChartSeries {
  name: string;
  labels: string[];
  values: number[];
}

export interface ChartProps extends Box {
  type: "doughnut" | "pie" | "bar" | "line" | "area";
  data: ChartSeries[];
  colors?: Color[];
  /** Doughnut hole, 10–90 (%) */
  hole?: number;
  /** Bar direction */
  direction?: "bar" | "col";
  stacked?: boolean;
  legend?: false | "top" | "bottom" | "left" | "right";
  labels?: boolean;
  /** Raw pptxgenjs chart options merged last */
  options?: Record<string, unknown>;
  name?: string;
}

export interface RingProps {
  /** center x */
  cx: number;
  /** center y */
  cy: number;
  /** outer radius px */
  r: number;
  /** 0–1 */
  value: number;
  color: Color;
  track?: Color;
  /** ring thickness px */
  thickness?: number;
  name?: string;
}

/** Native PowerPoint slide-number field (renumbers automatically when slides move). */
export interface SlideNumberProps extends Box, TextStyle {
  align?: "left" | "center" | "right";
  valign?: "top" | "middle" | "bottom";
}

export interface SlideProps {
  background?: Fill | { image: string };
  notes?: string;
  children?: Child;
}

export interface PresentationProps {
  /** Design canvas width in px (default 1280). Output is always 13.333in wide. */
  width?: number;
  /** Design canvas height in px (default 720) */
  height?: number;
  title?: string;
  author?: string;
  company?: string;
  /** Default text style for every <text> */
  font?: string;
  color?: Color;
  children?: Child;
}

export interface IntrinsicProps {
  presentation: PresentationProps;
  slide: SlideProps;
  text: TextProps;
  span: SpanProps;
  shape: ShapeProps;
  path: PathProps;
  line: LineProps;
  image: ImageProps;
  chart: ChartProps;
  ring: RingProps;
  slideNumber: SlideNumberProps;
  group: { children?: Child };
}
