import { renderToBlob, renderToSvg, type SvgOptions } from "pptx-jsx";
import { stories, type Story } from "./stories.js";

const svgOptions: SvgOptions = { fontMap: { "Poppins Medium": { family: "Poppins", weight: 500 } } };

const $ = <T extends HTMLElement>(sel: string) => document.querySelector(sel) as T;
const nav = $("#nav");
const title = $("#story-title");
const desc = $("#story-desc");
const group = $("#story-group");
const slidesEl = $("#slides");
const countEl = $("#slide-count");
const propsEl = $<HTMLTextAreaElement>("#props");
const propsError = $("#props-error");
const codeEl = $("#code");
const downloadBtn = $<HTMLButtonElement>("#download");
const statusEl = $("#status");
const resetBtn = $<HTMLButtonElement>("#reset");
const copyBtn = $<HTMLButtonElement>("#copy");

const argsById = new Map<string, Record<string, unknown>>();
let current: Story = stories[0];

function argsOf(s: Story) {
  if (!argsById.has(s.id)) argsById.set(s.id, structuredClone(s.args));
  return argsById.get(s.id)!;
}

function buildNav() {
  const groups = new Map<string, Story[]>();
  for (const s of stories) groups.set(s.group, [...(groups.get(s.group) ?? []), s]);
  nav.innerHTML = "";
  for (const [g, list] of groups) {
    const h = document.createElement("div");
    h.className = "nav-group";
    h.textContent = g;
    nav.append(h);
    for (const s of list) {
      const a = document.createElement("a");
      a.href = `#${s.id}`;
      a.textContent = s.name;
      a.dataset.id = s.id;
      nav.append(a);
    }
  }
}

function renderPreview() {
  const args = argsOf(current);
  try {
    const slides = renderToSvg(current.render(args), svgOptions);
    slidesEl.innerHTML = slides
      .map((s, i) => `<figure class="slide"><div class="frame">${s.svg}</div><figcaption>Slide ${i + 1}</figcaption></figure>`)
      .join("");
    slidesEl.classList.toggle("single", slides.length === 1);
    countEl.textContent = slides.length === 1 ? "1 slide" : `${slides.length} slides`;
    codeEl.textContent = current.code(args);
    propsError.hidden = true;
  } catch (e) {
    propsError.textContent = `Could not render: ${(e as Error).message}`;
    propsError.hidden = false;
  }
}

function select(id: string) {
  current = stories.find((s) => s.id === id) ?? stories[0];
  for (const a of nav.querySelectorAll("a")) a.classList.toggle("active", a.dataset.id === current.id);
  title.textContent = current.name;
  desc.textContent = current.description;
  group.textContent = current.group;
  const args = argsOf(current);
  const hasArgs = Object.keys(args).length > 0;
  propsEl.value = hasArgs ? JSON.stringify(args, null, 2) : "{}";
  propsEl.disabled = !hasArgs;
  resetBtn.disabled = !hasArgs;
  statusEl.textContent = "";
  renderPreview();
}

let timer: number | undefined;
propsEl.addEventListener("input", () => {
  clearTimeout(timer);
  timer = window.setTimeout(() => {
    try {
      argsById.set(current.id, JSON.parse(propsEl.value));
      renderPreview();
    } catch (e) {
      propsError.textContent = `Props must be valid JSON: ${(e as Error).message}`;
      propsError.hidden = false;
    }
  }, 250);
});

resetBtn.addEventListener("click", () => {
  argsById.delete(current.id);
  select(current.id);
});

copyBtn.addEventListener("click", async () => {
  try {
    await navigator.clipboard.writeText(codeEl.textContent ?? "");
    copyBtn.textContent = "Copied";
  } catch {
    const r = document.createRange();
    r.selectNodeContents(codeEl);
    getSelection()?.removeAllRanges();
    getSelection()?.addRange(r);
    copyBtn.textContent = "Selected";
  }
  setTimeout(() => (copyBtn.textContent = "Copy"), 1500);
});

// The .pptx is generated right here in the browser by the same renderer the Node build uses.
const downloadsPromise: Promise<any> = (window as any).claude?.use?.("downloads") ?? Promise.resolve(null);
downloadBtn.addEventListener("click", async () => {
  downloadBtn.disabled = true;
  statusEl.textContent = "Building .pptx…";
  try {
    const blob = await renderToBlob(current.render(argsOf(current)));
    const kb = Math.round(blob.size / 1024);
    const downloads = await downloadsPromise;
    if (!downloads) {
      statusEl.textContent = `Built ${kb} KB, but saving files isn't available in this view.`;
      return;
    }
    await downloads.save({ filename: `${current.id}.pptx`, data: blob });
    statusEl.textContent = `Saved ${current.id}.pptx (${kb} KB)`;
  } catch (e: any) {
    statusEl.textContent =
      e?.code === "declined" ? "Download cancelled." : `Couldn't build the file: ${e?.message ?? e}`;
  } finally {
    downloadBtn.disabled = false;
  }
});

buildNav();
const fromHash = () => select(location.hash.slice(1) || stories[0].id);
addEventListener("hashchange", fromHash);
fromHash();
// Re-render once Poppins has loaded so foreignObject text uses it.
document.fonts?.ready.then(renderPreview);
