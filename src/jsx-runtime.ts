import type { IntrinsicProps } from "./types.js";

export type Child = PNode | string | number | boolean | null | undefined | Child[];

/** A node in the slide tree before rendering. */
export interface PNode {
  type: string | Component<any> | typeof Fragment;
  props: Record<string, any> & { children?: Child };
}

export type Component<P = {}> = (props: P & { children?: Child }) => Child;

export const Fragment = Symbol.for("pptx-jsx.fragment");

export function jsx(type: PNode["type"], props: Record<string, any>, key?: unknown): PNode {
  return { type, props: props ?? {} };
}
export const jsxs = jsx;
export const jsxDEV = jsx;

export namespace JSX {
  export type Element = PNode;
  /** Components may return arrays (e.g. auto-paginated slides), fragments or null. */
  export type ElementType = keyof IntrinsicProps | ((props: any) => Child);
  export interface ElementChildrenAttribute {
    children: {};
  }
  export interface IntrinsicAttributes {
    key?: string | number;
  }
  export type IntrinsicElements = IntrinsicProps;
}
