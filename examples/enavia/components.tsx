import { clampText, paginate, textHeight, textWidth, type Child } from "pptx-jsx";
import { brand, sunset, sunsetBand, theme, type Status } from "./theme.js";

const W = 1280;
const H = 720;
const C = theme.color;

/** The wave artwork was drawn on a 2000px-wide canvas; scale it to the 1280px slide. */
const scaled = (d: string) => d.replace(/-?\d+(\.\d+)?/g, (n) => (parseFloat(n) * (W / 2000)).toFixed(2));

const WAVE = {
  lavender: "M870 0 C960 90 1080 190 1135 265 C1090 330 1030 400 1000 440 C880 600 790 760 815 950 C822 1030 835 1090 850 1130 L2000 1130 L2000 0 Z",
  purple: "M1090 0 L1145 120 C1165 230 1100 340 1000 440 C880 560 775 680 775 790 C775 930 860 1040 945 1130 L2000 1130 L2000 0 Z",
  sun: "M870 0 C950 55 1050 90 1145 120 C1240 160 1280 220 1278 290 C1275 420 1160 590 1120 720 C1060 900 1130 1040 1270 1130 L2000 1130 L2000 0 Z",
  canvas: "M935 0 C1010 45 1090 75 1170 100 C1290 145 1332 215 1330 290 C1327 420 1212 590 1172 720 C1112 900 1182 1040 1322 1130 L2000 1130 L2000 0 Z",
};

// ---------------------------------------------------------------------------
// Brand chrome
// ---------------------------------------------------------------------------

export function Logo({ x, y, h = 40 }: { x: number; y: number; h?: number }) {
  if (brand.logo) return <image src={brand.logo} x={x} y={y} w={h * 2.6} h={h} fit="contain" name="Logo" />;
  return (
    <text x={x} y={y} w={h * 3} h={h} size={h * 0.62} bold color={C.ink} valign="bottom" name="Logo">
      enavia
      <span color={C.orange}>+</span>
    </text>
  );
}

function Waves({ variant }: { variant: "cover" | "section" }) {
  return (
    <>
      <path d={scaled(WAVE.lavender)} fill={C.lavender} name="Wave back" />
      <path d={scaled(WAVE.purple)} fill={C.ink} name="Wave purple" />
      <path d={scaled(WAVE.sun)} fill={variant === "cover" ? sunset : sunsetBand} name="Wave sun" />
      {variant === "section" && <path d={scaled(WAVE.canvas)} fill={C.canvas} name="Wave canvas" />}
    </>
  );
}

export function CoverSlide({ title, meta }: { title: string; meta?: string }) {
  return (
    <slide background={C.white}>
      <Waves variant="cover" />
      <Logo x={13} y={52} />
      <text x={13} y={250} w={470} h={130} size={38} font={theme.fontMedium} color="000000" valign="bottom" lineHeight={1.05} name="Title">
        {title}
      </text>
      {meta && (
        <text x={13} y={396} w={470} h={22} size={13} color={C.muted} name="Meta">
          {meta}
        </text>
      )}
    </slide>
  );
}

export function SectionSlide(props: { title: string; subtitle?: string; meta?: string; image?: string }) {
  return (
    <slide background={C.white}>
      <Waves variant="section" />
      {props.image ? (
        <image src={props.image} x={860} y={0} w={420} h={H} fit="cover" name="Section image" />
      ) : (
        <text x={840} y={348} w={360} h={24} size={11} color={C.muted} align="center" name="Image placeholder">
          Add image
        </text>
      )}
      <Logo x={13} y={52} />
      <text x={13} y={292} w={480} h={56} size={38} font={theme.fontMedium} color="000000" valign="bottom" fit name="Title">
        {props.title}
      </text>
      {props.subtitle && (
        <text x={13} y={356} w={480} h={26} size={16} color={C.ink} name="Subtitle">
          {props.subtitle}
        </text>
      )}
      {props.meta && (
        <text x={13} y={391} w={480} h={22} size={13} color={C.muted} name="Meta">
          {props.meta}
        </text>
      )}
    </slide>
  );
}

/** Standard content slide: title, rule, logo, slide number and bottom bar. */
export function ContentSlide({ title, children }: { title: string; children?: Child }) {
  return (
    <slide background={C.white}>
      {title && (
        <>
          <text x={26} y={18} w={1228} h={42} size={26} font={theme.fontMedium} color={C.ink} valign="middle" fit name="Title">
            {title}
          </text>
          <line x1={26} y1={75.5} x2={1254} y2={75.5} color={C.rule} width={1} name="Title rule" />
        </>
      )}
      {children}
      <Logo x={13} y={659} />
      <shape x={1240} y={666} w={26} h={26} fill={C.ink} name="Number badge" />
      <slideNumber x={1240} y={666} w={26} h={26} size={11} color={C.white} />
      <shape x={0} y={704} w={W} h={16} fill={C.ink} name="Footer bar" />
    </slide>
  );
}

// ---------------------------------------------------------------------------
// Donut stats
// ---------------------------------------------------------------------------

export interface DonutItem {
  label: string;
  value: number;
  color: string;
}

function Donut(props: { x: number; w: number; caption: string; center: string; items: DonutItem[]; unit: string }) {
  const r = 157;
  const cx = props.x + props.w / 2;
  const cy = 270;
  const legendX = props.x + props.w / 2 - 166;
  const legendW = 376;
  return (
    <>
      <chart
        type="doughnut"
        x={cx - r}
        y={cy - r}
        w={r * 2}
        h={r * 2}
        hole={80}
        colors={props.items.map((i) => i.color)}
        data={[{ name: props.caption, labels: props.items.map((i) => i.label), values: props.items.map((i) => i.value) }]}
        options={{ layout: { x: 0, y: 0, w: 1, h: 1 }, dataBorder: { pt: 0.75, color: "FFFFFF" } }}
        name={`${props.caption} chart`}
      />
      <text x={cx - 90} y={cy - 26} w={180} h={20} size={11} color={C.text} align="center">
        {props.caption}
      </text>
      <text x={cx - 90} y={cy - 4} w={180} h={28} size={15} color={C.text} align="center">
        {props.center}
      </text>
      {props.items.map((item, i) => {
        const y = 472 + i * 29;
        return (
          <>
            <shape x={legendX} y={y + 6} w={10} h={10} radius={1.5} fill={item.color} />
            <text x={legendX + 17} y={y} w={240} h={22} size={16} color={C.ink} valign="middle">
              {item.label}
            </text>
            <text x={legendX + legendW - 140} y={y} w={140} h={22} size={16} color={C.muted} align="right" valign="middle">
              {`${item.value} ${props.unit}`}
            </text>
          </>
        );
      })}
    </>
  );
}

export function TaskProgressSlide(props: { title: string; byStatus: DonutItem[]; readiness: { completed: number; remaining: number } }) {
  const total = props.byStatus.reduce((s, i) => s + i.value, 0);
  const { completed, remaining } = props.readiness;
  const pct = completed + remaining ? (completed / (completed + remaining)) * 100 : 0;
  return (
    <ContentSlide title={props.title}>
      <Donut x={0} w={640} caption="Total" center={String(total)} items={props.byStatus} unit="Tasks" />
      <line x1={626} y1={102} x2={626} y2={618} color={C.rowRule} />
      <Donut
        x={612}
        w={640}
        caption="Readiness (%)"
        center={pct.toFixed(2)}
        unit="Tasks"
        items={[
          { label: "Completed", value: completed, color: theme.status.completed },
          { label: "Remaining", value: remaining, color: theme.status.delayed },
        ]}
      />
    </ContentSlide>
  );
}

// ---------------------------------------------------------------------------
// Task tables with progress rings (auto-paginated)
// ---------------------------------------------------------------------------

export interface Task {
  id: string;
  title: string;
  comment?: string;
  /** 0–100 */
  progress: number;
  status: Status;
}

const COLS = { num: 13, title: 66, comment: 614, progress: 1141, end: 1267 };
const BODY = 14;
const LINE = 1.45;
const TABLE_TOP = 89;
const HEADER_H = 51;
const TABLE_BOTTOM = 640;

function rowHeight(t: Task) {
  const titleH = textHeight(t.title, { size: BODY, width: COLS.comment - COLS.title - 20, lineHeight: LINE });
  const commentH = t.comment ? textHeight(t.comment, { size: BODY, width: COLS.progress - COLS.comment - 20, lineHeight: LINE }) : 0;
  return Math.max(42, Math.max(titleH, commentH) + 18);
}

function ProgressCell({ task, y, h }: { task: Task; y: number; h: number }) {
  const color = task.progress >= 100 ? theme.status.completed : theme.status[task.status];
  const cy = y + h / 2;
  return (
    <>
      <ring cx={COLS.progress + 63} cy={cy} r={9.5} thickness={2.6} value={task.progress / 100} color={color} track="E5E7EB" />
      <text x={COLS.progress + 78} y={cy - 10} w={60} h={20} size={12} color={color} valign="middle">
        {`${Math.round(task.progress)}%`}
      </text>
    </>
  );
}

function TaskTable({ tasks }: { tasks: Task[] }) {
  let y = TABLE_TOP + HEADER_H;
  const header = (
    <>
      <shape x={COLS.num} y={TABLE_TOP} w={COLS.end - COLS.num} h={HEADER_H} radius={8} fill={C.ink} name="Table header" />
      <shape x={COLS.num} y={TABLE_TOP + HEADER_H / 2} w={COLS.end - COLS.num} h={HEADER_H / 2} fill={C.ink} />
      {[COLS.title, COLS.comment, COLS.progress].map((x) => (
        <line x1={x} y1={TABLE_TOP} x2={x} y2={TABLE_TOP + HEADER_H} color="3A2163" width={0.75} />
      ))}
      <text x={COLS.num + 13} y={TABLE_TOP} w={30} h={HEADER_H} size={17} font={theme.fontMedium} color={C.white} valign="middle">#</text>
      <text x={COLS.title + 7} y={TABLE_TOP} w={540} h={HEADER_H} size={17} font={theme.fontMedium} color={C.white} valign="middle">
        Current Quarter Key Initiatives &amp; Accomplishments
      </text>
      <text x={COLS.comment + 7} y={TABLE_TOP} w={500} h={HEADER_H} size={17} font={theme.fontMedium} color={C.white} valign="middle">
        Comments / Risk
      </text>
      <text x={COLS.progress} y={TABLE_TOP} w={COLS.end - COLS.progress} h={HEADER_H} size={17} font={theme.fontMedium} color={C.white} align="center" valign="middle">
        Progress
      </text>
    </>
  );
  const rows = tasks.map((t, i) => {
    const h = rowHeight(t);
    const top = y;
    y += h;
    return (
      <>
        {i % 2 === 1 && <shape x={COLS.num} y={top} w={COLS.end - COLS.num} h={h} fill="FAF9FC" />}
        <text x={COLS.num + 13} y={top} w={COLS.title - COLS.num - 13} h={h} size={BODY} color={C.text} valign="middle">
          {`${t.id}.`}
        </text>
        <text x={COLS.title + 7} y={top} w={COLS.comment - COLS.title - 14} h={h} size={BODY} color={C.text} valign="middle" lineHeight={LINE}>
          {t.title}
        </text>
        {t.comment && (
          <text x={COLS.comment + 7} y={top} w={COLS.progress - COLS.comment - 14} h={h} size={BODY} color={C.text} valign="middle" lineHeight={LINE}>
            {t.comment}
          </text>
        )}
        <ProgressCell task={t} y={top} h={h} />
        <line x1={COLS.num} y1={top + h} x2={COLS.end} y2={top + h} color={C.rowRule} width={1} />
      </>
    );
  });
  return (
    <>
      {header}
      {rows}
    </>
  );
}

/** One or more slides; rows flow onto new slides when they no longer fit. */
export function TaskTableSlides({ title, tasks }: { title: string; tasks: Task[] }) {
  const pages = paginate(tasks, rowHeight, TABLE_BOTTOM - TABLE_TOP - HEADER_H);
  return pages.map((page, i) => (
    <ContentSlide title={pages.length > 1 ? `${title} (${i + 1}/${pages.length})` : title}>
      {page.length ? (
        <TaskTable tasks={page} />
      ) : (
        <text x={26} y={140} w={1228} h={40} size={22} color={C.muted} align="center">No tasks this quarter</text>
      )}
    </ContentSlide>
  ));
}

// ---------------------------------------------------------------------------
// Gantt timeline (auto-paginated)
// ---------------------------------------------------------------------------

export interface TimelineItem {
  label: string;
  /** ISO date */
  start: string;
  /** ISO date; omit for a milestone */
  end?: string;
  kind: "group" | "task" | "milestone";
}

const G = { left: 40, right: 1239, top: 89, header: 40, row: 36, bottom: 640 };

export function TimelineSlides(props: { title: string; from: number; to: number; today: string; items: TimelineItem[] }) {
  const years = Array.from({ length: props.to - props.from + 1 }, (_, i) => props.from + i);
  const spanW = G.right - G.left;
  const t0 = Date.UTC(props.from, 0, 1);
  const t1 = Date.UTC(props.to + 1, 0, 1);
  const xOf = (iso: string) => G.left + ((Date.parse(iso) - t0) / (t1 - t0)) * spanW;
  const clampX = (x: number) => Math.max(G.left, Math.min(G.right, x));
  const perPage = Math.floor((G.bottom - G.top - G.header - 12) / G.row);
  const pages = paginate(props.items, () => 1, perPage);
  const yearW = spanW / years.length;

  return pages.map((page, pi) => {
    const rowsBottom = G.top + G.header + 8 + Math.max(page.length, 1) * G.row;
    const todayX = xOf(props.today);
    return (
      <ContentSlide title={pages.length > 1 ? `${props.title} (${pi + 1}/${pages.length})` : props.title}>
        <shape x={G.left} y={G.top} w={spanW} h={G.header} fill={C.ink} name="Timeline header" />
        {years.map((yr, i) => {
          const x = G.left + i * yearW;
          return (
            <>
              {i > 0 && <line x1={x} y1={G.top} x2={x} y2={rowsBottom} color={i > 0 ? "E9E7EF" : C.ink} width={0.75} />}
              {i > 0 && <line x1={x} y1={G.top} x2={x} y2={G.top + G.header} color="4A3470" width={0.75} />}
              <text x={x} y={G.top + 3} w={yearW} h={16} size={10} color={C.white} align="center" bold>
                {String(yr)}
              </text>
              {[1, 2, 3, 4].map((q) => (
                <text x={x + ((q - 1) * yearW) / 4} y={G.top + 20} w={yearW / 4} h={16} size={9} color="D9D2E9" align="center">
                  {`Q${q}`}
                </text>
              ))}
            </>
          );
        })}
        {page.length > 0 && todayX >= G.left && todayX <= G.right && (
          <>
            <line x1={todayX} y1={G.top + G.header} x2={todayX} y2={rowsBottom} color={C.today} width={1.5} name="Today" />
            <shape kind="roundRect" x={todayX - 22} y={rowsBottom + 4} w={44} h={16} radius={8} fill={C.today} />
            <text x={todayX - 22} y={rowsBottom + 4} w={44} h={16} size={9} bold color={C.white} align="center" valign="middle">
              Today
            </text>
          </>
        )}
        {page.length === 0 && (
          <text x={G.left} y={G.top + G.header + 20} w={spanW} h={30} size={14} color={C.muted} align="center">
            No initiatives scheduled
          </text>
        )}
        {page.map((item, i) => {
          const rowY = G.top + G.header + 8 + i * G.row;
          const midY = rowY + G.row / 2;
          const x1 = clampX(xOf(item.start));
          const x2 = item.end ? clampX(xOf(item.end)) : x1;
          const isGroup = item.kind === "group";
          const barH = isGroup ? 18 : 16;
          const size = isGroup ? 11 : 10.5;
          const labelW = textWidth(item.label, size) + 12;
          const color = isGroup ? C.ink : C.lilac;

          // Label: inside a long group bar, otherwise right of the bar, or left of it near the right edge.
          const inside = isGroup && x2 - x1 > labelW + 12;
          const fitsRight = x2 + 10 + labelW <= G.right;
          const label = clampText(item.label, { size, width: 420, maxLines: 1 });
          const labelBox = inside
            ? { x: x1 + 8, w: x2 - x1 - 16, color: C.white, align: "left" as const }
            : fitsRight
              ? { x: x2 + 10, w: Math.min(labelW, G.right - x2 - 10), color: C.text, align: "left" as const }
              : { x: Math.max(G.left, x1 - 10 - labelW), w: Math.min(labelW, x1 - 10 - G.left), color: C.text, align: "right" as const };

          return (
            <>
              {item.kind === "milestone" ? (
                <shape kind="diamond" x={x1 - 7} y={midY - 7} w={14} h={14} fill={C.orange} />
              ) : (
                <shape x={x1} y={midY - barH / 2} w={Math.max(4, x2 - x1)} h={barH} radius={isGroup ? 4 : 3} fill={color} />
              )}
              <text x={labelBox.x} y={rowY} w={labelBox.w} h={G.row} size={size} color={labelBox.color} bold={isGroup} align={labelBox.align} valign="middle" wrap={false}>
                {label}
              </text>
            </>
          );
        })}
      </ContentSlide>
    );
  });
}

// ---------------------------------------------------------------------------
// Budget
// ---------------------------------------------------------------------------

export interface BudgetLine {
  label: string;
  amount: number;
}

const PALETTE = ["24074A", "FF7A00", "9B7FBF", "FFC21A", "D9488F", "3B82F6", "0FAF1A", "F4511E", "5B4B8A", "14B8A6", "A3A3A3", "84CC16"];

export function BudgetSlide(props: { title: string; currency: string; total: number; spend: number; lines: BudgetLine[] }) {
  const fmt = (n: number) =>
    new Intl.NumberFormat("en-GB", { style: "currency", currency: props.currency, maximumFractionDigits: 0 }).format(n);
  const allocated = props.lines.reduce((s, l) => s + l.amount, 0);
  const lines = [...props.lines].sort((a, b) => b.amount - a.amount);
  const spentPct = props.total ? Math.min(1, props.spend / props.total) : 0;
  const rowH = Math.min(34, 430 / Math.max(lines.length, 1));
  const listX = 560;
  const listW = 680;
  return (
    <ContentSlide title={props.title}>
      <text x={26} y={92} w={700} h={34} size={22} color={C.ink} valign="middle">
        Total budget
        <span color={C.orange}>{`  ${fmt(props.total)}`}</span>
      </text>
      <text x={26} y={128} w={700} h={22} size={13} color={C.text} valign="middle">
        Spent
        <span color={C.orange} bold>{`  ${fmt(props.spend)}`}</span>
        <span color={C.muted}>{`  ·  ${(spentPct * 100).toFixed(0)}% of budget`}</span>
      </text>
      {/* spend meter */}
      <shape x={26} y={158} w={480} h={8} radius={4} fill="EEEAF4" />
      <shape x={26} y={158} w={Math.max(8, 480 * spentPct)} h={8} radius={4} fill={sunset} />

      <chart
        type="doughnut"
        x={86}
        y={200}
        w={380}
        h={380}
        hole={72}
        colors={lines.map((_, i) => PALETTE[i % PALETTE.length])}
        data={[{ name: "Allocation", labels: lines.map((l) => l.label), values: lines.map((l) => l.amount) }]}
        options={{ layout: { x: 0, y: 0, w: 1, h: 1 } }}
      />
      <text x={176} y={364} w={200} h={20} size={12} color={C.muted} align="center">Allocated</text>
      <text x={176} y={384} w={200} h={30} size={20} bold color={C.ink} align="center">{fmt(allocated)}</text>

      <text x={listX} y={200} w={300} h={20} size={11} bold color={C.muted}>CATEGORY</text>
      <text x={listX + listW - 220} y={200} w={140} h={20} size={11} bold color={C.muted} align="right">AMOUNT</text>
      <text x={listX + listW - 70} y={200} w={70} h={20} size={11} bold color={C.muted} align="right">SHARE</text>
      {lines.map((l, i) => {
        const y = 226 + i * rowH;
        const share = allocated ? l.amount / allocated : 0;
        return (
          <>
            <shape x={listX} y={y + rowH / 2 - 5} w={10} h={10} radius={2} fill={PALETTE[i % PALETTE.length]} />
            <text x={listX + 18} y={y} w={listW - 260} h={rowH} size={13} color={C.ink} valign="middle" wrap={false}>
              {clampText(l.label, { size: 13, width: listW - 270, maxLines: 1 })}
            </text>
            <text x={listX + listW - 220} y={y} w={140} h={rowH} size={13} color={C.text} align="right" valign="middle">
              {fmt(l.amount)}
            </text>
            <text x={listX + listW - 70} y={y} w={70} h={rowH} size={13} color={C.muted} align="right" valign="middle">
              {`${(share * 100).toFixed(1)}%`}
            </text>
            <line x1={listX} y1={y + rowH} x2={listX + listW} y2={y + rowH} color={C.rowRule} />
          </>
        );
      })}
    </ContentSlide>
  );
}

// ---------------------------------------------------------------------------
// Simple slides
// ---------------------------------------------------------------------------

export function MessageSlide({ title, message }: { title: string; message: string }) {
  return (
    <ContentSlide title={title}>
      <shape kind="ellipse" x={604} y={250} w={72} h={72} fill="F1EEF7" />
      <shape kind="ellipse" x={622} y={268} w={36} h={36} fill={C.white} stroke={{ color: theme.status.completed, width: 3 }} />
      <text x={26} y={340} w={1228} h={40} size={24} color={C.text} align="center">
        {message}
      </text>
    </ContentSlide>
  );
}

export function ThankYouSlide() {
  return (
    <ContentSlide title="">
      <text x={0} y={320} w={W} h={70} size={44} font={theme.fontMedium} color="000000" align="center" valign="middle">
        Thank you
      </text>
    </ContentSlide>
  );
}
