export type PathPoint =
  | { kind: "move" | "line"; x: number; y: number }
  | { kind: "quad"; x: number; y: number; x1: number; y1: number }
  | { kind: "cubic"; x: number; y: number; x1: number; y1: number; x2: number; y2: number }
  | { kind: "close" };

/**
 * Parse SVG path data (M L H V C S Q Z, absolute or relative) into absolute points
 * plus the bounding box of all points, including control points.
 */
export function parsePath(d: string) {
  const tokens = d.match(/[a-zA-Z]|-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/g) ?? [];
  const points: PathPoint[] = [];
  let i = 0;
  let cmd = "";
  let cx = 0, cy = 0, sx = 0, sy = 0;
  let lastC2: [number, number] | null = null;
  const num = () => parseFloat(tokens[i++]);

  while (i < tokens.length) {
    if (/[a-zA-Z]/.test(tokens[i])) cmd = tokens[i++];
    else if (cmd.toUpperCase() === "Z" || !cmd) throw new Error(`Malformed path data: ${d}`);
    const rel = cmd === cmd.toLowerCase();
    const ox = rel ? cx : 0, oy = rel ? cy : 0;
    switch (cmd.toUpperCase()) {
      case "M": {
        cx = ox + num(); cy = oy + num(); sx = cx; sy = cy;
        points.push({ kind: "move", x: cx, y: cy });
        cmd = rel ? "l" : "L"; // subsequent pairs are implicit line-tos
        lastC2 = null;
        break;
      }
      case "L": cx = ox + num(); cy = oy + num(); points.push({ kind: "line", x: cx, y: cy }); lastC2 = null; break;
      case "H": cx = (rel ? cx : 0) + num(); points.push({ kind: "line", x: cx, y: cy }); lastC2 = null; break;
      case "V": cy = (rel ? cy : 0) + num(); points.push({ kind: "line", x: cx, y: cy }); lastC2 = null; break;
      case "C": {
        const x1 = ox + num(), y1 = oy + num(), x2 = ox + num(), y2 = oy + num();
        cx = ox + num(); cy = oy + num();
        points.push({ kind: "cubic", x1, y1, x2, y2, x: cx, y: cy });
        lastC2 = [x2, y2];
        break;
      }
      case "S": {
        const [x1, y1] = lastC2 ? [2 * cx - lastC2[0], 2 * cy - lastC2[1]] : [cx, cy];
        const x2 = ox + num(), y2 = oy + num();
        cx = ox + num(); cy = oy + num();
        points.push({ kind: "cubic", x1, y1, x2, y2, x: cx, y: cy });
        lastC2 = [x2, y2];
        break;
      }
      case "Q": {
        const x1 = ox + num(), y1 = oy + num();
        cx = ox + num(); cy = oy + num();
        points.push({ kind: "quad", x1, y1, x: cx, y: cy });
        lastC2 = null;
        break;
      }
      case "Z": points.push({ kind: "close" }); cx = sx; cy = sy; lastC2 = null; break;
      default: throw new Error(`Unsupported path command "${cmd}" in: ${d}`);
    }
  }

  const xs: number[] = [], ys: number[] = [];
  for (const p of points) {
    if (p.kind === "close") continue;
    xs.push(p.x); ys.push(p.y);
    if (p.kind === "quad" || p.kind === "cubic") { xs.push(p.x1); ys.push(p.y1); }
    if (p.kind === "cubic") { xs.push(p.x2); ys.push(p.y2); }
  }
  const x = Math.min(...xs), y = Math.min(...ys);
  const box = { x, y, w: Math.max(...xs) - x || 1, h: Math.max(...ys) - y || 1 };
  return { box, points };
}
