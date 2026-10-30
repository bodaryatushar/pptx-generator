import { existsSync } from "node:fs";
import { mkdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { renderToFile } from "pptx-jsx";
import { EnaviaDeck } from "./deck.js";
import { brand } from "./theme.js";

// Drop your logo at examples/enavia/assets/logo.png to replace the wordmark.
const logo = fileURLToPath(new URL("./assets/logo.png", import.meta.url));
if (existsSync(logo)) brand.logo = logo;

await mkdir("out", { recursive: true });
console.log("wrote", await renderToFile(<EnaviaDeck />, "out/enavia-launch.pptx"));
