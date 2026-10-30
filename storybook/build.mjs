// Bundles the storybook into one self-contained HTML file: out/storybook.html
import { build } from "esbuild";
import { mkdir, readFile, writeFile } from "node:fs/promises";

const root = new URL("..", import.meta.url).pathname;

const result = await build({
  entryPoints: [`${root}storybook/main.ts`],
  bundle: true,
  format: "iife",
  platform: "browser",
  minify: true,
  write: false,
  jsx: "automatic",
  jsxImportSource: "pptx-jsx",
  tsconfig: `${root}tsconfig.examples.json`,
  alias: { "pptx-jsx": `${root}src/index.ts`, "pptx-jsx/jsx-runtime": `${root}src/jsx-runtime.ts` },
  external: ["node:fs/promises", "node:fs", "node:url"],
});

const js = result.outputFiles[0].text.replace(/<\/script/gi, "<\\/script");
const html = (await readFile(`${root}storybook/index.html`, "utf8")).replace("/*BUNDLE*/", () => js);
await mkdir(`${root}out`, { recursive: true });
await writeFile(`${root}out/storybook.html`, html);
console.log(`wrote out/storybook.html (${Math.round(html.length / 1024)} KB)`);
