# pptx-jsx

Turn JSX components into **native, editable PowerPoint** files: real shapes, text boxes, charts and freeform paths. No screenshots.

```tsx
/** @jsxImportSource pptx-jsx */
import { renderToFile } from "pptx-jsx";

const Title = ({ children }) => <text x={40} y={30} w={1200} h={50} size={28} color="24074A">{children}</text>;

await renderToFile(
  <presentation font="Poppins">
    <slide background="FFFFFF">
      <Title>Q3 results</Title>
      <path d="M700 0 C800 200 600 500 760 720 L1280 720 L1280 0 Z"
            fill={{ type: "linear", angle: 45, stops: [{ pos: 0, color: "FFD600" }, { pos: 100, color: "FF7A00" }] }} />
      <ring cx={100} cy={200} r={20} value={0.72} color="0FAF1A" />
    </slide>
  </presentation>,
  "deck.pptx",
);
```

## Model

* Everything uses design **px** on a 1280×720 canvas (`<presentation width height>` changes it). Output is 13.333in × 7.5in (16:9).
* Components are plain functions. They can return a **single slide, an array of slides, or a fragment**, so a `TaskTableSlides` component can split its rows across as many slides as it needs.
* Colors are hex strings without `#`.

## Elements

| Element | Purpose |
|---|---|
| `<presentation>` | Root. `font`, `color`, `title`, `author`, `width`, `height` |
| `<slide>` | `background` (color, gradient, `{image}`), `notes` |
| `<text>` | Box text. Children are strings or `<span>` runs with their own style. Supports `align`, `valign`, `lineHeight`, `fit` (shrink on overflow) |
| `<shape>` | `rect`, `roundRect` (`radius` in px), `ellipse`, `diamond`, … with `fill`, `stroke`, `shadow` |
| `<path>` | SVG path data (`M L H V C S Q Z`) converted to a native freeform. Use it for waves and blobs |
| `<line>` | `x1 y1 x2 y2` |
| `<image>` | `src`/`data`, `fit="cover" \| "contain"` |
| `<chart>` | Native, editable chart: `doughnut`, `pie`, `bar`, `line`, `area` |
| `<ring>` | Progress ring built from block arcs |
| `<slideNumber>` | Native slide-number field (renumbers when slides are moved) |
| `<group>` | Logical grouping only |

**Gradients** (`{type:"linear", angle, stops}` or `radial`) work on any `fill`. pptxgenjs can't create gradients, so the renderer writes them into the OOXML afterwards.

**Layout helpers:** `textHeight`, `wrapText`, `clampText` and `paginate` estimate how much text fits. Use them for row heights and to split content across slides.

## Example: Enavia launch deck

`examples/enavia` rebuilds the reference deck: a wave cover with a gradient, section dividers, donut KPI slides, budget slides, a Gantt timeline with a "today" marker, and task tables with progress rings. Timelines and task tables move rows onto new slides automatically when they don't fit.

```bash
npm install
npm run example:enavia          # -> out/enavia-launch.pptx
npm run preview -- out/enavia-launch.pptx   # PNGs via LibreOffice (optional)
```

Add `examples/enavia/assets/logo.png` to use the real logo instead of the placeholder wordmark.

## Use it in any project

1. Install it (from git, or publish it to npm or your private registry), then point JSX at it in `tsconfig.json`:
   `"jsx": "react-jsx", "jsxImportSource": "pptx-jsx"`. You don't need React.
2. **Node / API routes:** `renderToFile(<Deck />, "deck.pptx")`, or send `await renderToBuffer(<Deck />)` with the `PPTX_MIME` content type.
3. **Browser (React, Vue, plain JS):** `renderToSvg(<Deck />)` returns one SVG string per slide for previews, and `renderToBlob(<Deck />)` builds the real `.pptx` in the browser.
   In React: `<div dangerouslySetInnerHTML={{ __html: slides[0].svg }} />`.

## Storybook

```bash
npm run storybook   # -> out/storybook.html (single self-contained file)
```

The storybook lists every component with a live SVG preview, editable JSON props, a usage snippet, and a button that builds the `.pptx` in the browser. Stories live in `storybook/stories.tsx`; add an entry to show a new component.

## Notes

* **Fonts:** PowerPoint can't embed fonts from code. Viewers need Poppins installed, or you can switch `theme.font` to a safe font.
* **Doughnut hole size** is respected by PowerPoint and Keynote. LibreOffice ignores it in previews.
* Use from another project: run `npm run build`, then set `"jsxImportSource": "pptx-jsx"` in `tsconfig.json`.
