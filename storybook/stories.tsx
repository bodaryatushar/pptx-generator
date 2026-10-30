import type { Child } from "pptx-jsx";
import {
  BudgetSlide,
  CoverSlide,
  MessageSlide,
  SectionSlide,
  TaskProgressSlide,
  TaskTableSlides,
  ThankYouSlide,
  TimelineSlides,
} from "../examples/enavia/components.js";
import { departments, EnaviaDeck } from "../examples/enavia/deck.js";
import { theme } from "../examples/enavia/theme.js";

export interface Story {
  id: string;
  group: string;
  name: string;
  description: string;
  /** Editable props shown in the Props panel */
  args: Record<string, unknown>;
  render: (args: any) => Child;
  /** Usage snippet shown in the Code panel */
  code: (args: any) => string;
}

const clinical = departments.find((d) => d.name === "Clinical")!;

const wrap = (slides: Child) => (
  <presentation font={theme.font} color={theme.color.text}>
    {slides}
  </presentation>
);

const jsxProps = (args: Record<string, unknown>) =>
  Object.entries(args)
    .map(([k, v]) => {
      if (typeof v === "string") return `  ${k}="${v}"`;
      const json = JSON.stringify(v, null, 2);
      return `  ${k}={${json.length > 120 ? json.replace(/\n/g, "\n  ") : JSON.stringify(v)}}`;
    })
    .join("\n");

const usage = (name: string) => (args: Record<string, unknown>) =>
  `import { renderToFile } from "pptx-jsx";
import { ${name} } from "./components";

await renderToFile(
  <presentation font="Poppins">
    <${name}
${jsxProps(args).replace(/^/gm, "    ")}
    />
  </presentation>,
  "deck.pptx",
);`;

export const stories: Story[] = [
  {
    id: "cover",
    group: "Title & section",
    name: "Cover",
    description: "Opening slide with the purple wave and a sunset gradient. The waves are native freeform shapes drawn from SVG path data.",
    args: { title: "Neurostax Full Launch 2027", meta: "Version: 30th October 2026" },
    render: (a) => wrap(<CoverSlide {...a} />),
    code: usage("CoverSlide"),
  },
  {
    id: "section",
    group: "Title & section",
    name: "Section divider",
    description: "Department divider. The grey panel holds an optional image (pass image as a file path or URL).",
    args: { title: "Clinical", subtitle: "Key Initiatives and Accomplishments", meta: "Q4 2026 and Q1 2027" },
    render: (a) => wrap(<SectionSlide {...a} />),
    code: usage("SectionSlide"),
  },
  {
    id: "task-progress",
    group: "Data",
    name: "Task progress",
    description: "Two native doughnut charts with a total in the middle and a legend. The chart data stays editable in PowerPoint.",
    args: {
      title: "Neurostax Full Launch 2027 Task Progress",
      byStatus: [
        { label: "Not Started", value: 32, color: theme.status.notStarted },
        { label: "In Progress", value: 1, color: theme.status.inProgress },
        { label: "Almost delayed", value: 29, color: theme.status.almostDelayed },
        { label: "Delayed", value: 82, color: theme.status.delayed },
        { label: "Completed", value: 42, color: theme.status.completed },
      ],
      readiness: { completed: 10, remaining: 5 },
    },
    render: (a) => wrap(<TaskProgressSlide {...a} />),
    code: usage("TaskProgressSlide"),
  },
  {
    id: "budget",
    group: "Data",
    name: "Budget",
    description: "Total and spend headline, a spend meter, an allocation doughnut and a ranked category list.",
    args: {
      title: "Strategic Imperatives budget",
      currency: "EUR",
      total: 1500000,
      spend: 1017355,
      lines: [
        { label: "Prep the Market", amount: 640000 },
        { label: "Prep the Brand", amount: 520000 },
        { label: "Prep the Organisation", amount: 340000 },
      ],
    },
    render: (a) => wrap(<BudgetSlide {...a} />),
    code: usage("BudgetSlide"),
  },
  {
    id: "task-table",
    group: "Data",
    name: "Task table",
    description: "Tasks with comments and progress rings. Rows grow with their text, and the table continues on a new slide when it runs out of room. Add tasks in Props to see it split.",
    args: { title: "Clinical department tasks", tasks: clinical.tasks.slice(0, 9) },
    render: (a) => wrap(<TaskTableSlides {...a} />),
    code: usage("TaskTableSlides"),
  },
  {
    id: "timeline",
    group: "Data",
    name: "Timeline",
    description: "Gantt view by quarter with group bars, task bars, milestones and a Today line. Labels move inside, right or left of a bar so they never get cut off.",
    args: { title: "Clinical Key Initiative & Accomplishment Timeline", from: 2026, to: 2030, today: "2026-10-30", items: clinical.timeline.slice(0, 12) },
    render: (a) => wrap(<TimelineSlides {...a} />),
    code: usage("TimelineSlides"),
  },
  {
    id: "message",
    group: "Simple",
    name: "Message",
    description: "Empty state, used for sections with nothing to report.",
    args: { title: "Risks", message: "No risks found" },
    render: (a) => wrap(<MessageSlide {...a} />),
    code: usage("MessageSlide"),
  },
  {
    id: "thank-you",
    group: "Simple",
    name: "Thank you",
    description: "Closing slide.",
    args: {},
    render: () => wrap(<ThankYouSlide />),
    code: usage("ThankYouSlide"),
  },
  {
    id: "primitives",
    group: "Library",
    name: "Primitives",
    description: "The building blocks every component above is made from: text runs, shapes, gradients, SVG paths, rings and charts.",
    args: { accent: "FF7A00", ink: "24074A", progress: 0.72 },
    render: (a) => (
      <presentation font="Poppins" color="374151">
        <slide background="FFFFFF">
          <path
            d="M820 0 C900 160 760 330 840 470 C900 580 1000 650 1080 720 L1280 720 L1280 0 Z"
            fill={{ type: "linear", angle: 45, stops: [{ pos: 0, color: "FFD600" }, { pos: 100, color: a.accent }] }}
          />
          <text x={60} y={60} w={700} h={60} size={40} color={a.ink} font="Poppins Medium">
            Hello from <span color={a.accent}>JSX</span>
          </text>
          <text x={60} y={128} w={640} h={60} size={16} lineHeight={1.5}>
            Every element below is a native PowerPoint object you can select and edit.
          </text>
          <shape x={60} y={220} w={180} h={110} radius={14} fill={a.ink} shadow={{ blur: 18, offset: 6, opacity: 0.25 }} />
          <text x={60} y={220} w={180} h={110} size={14} color="FFFFFF" align="center" valign="middle">roundRect + shadow</text>
          <shape kind="ellipse" x={270} y={220} w={110} h={110} fill={{ type: "radial", stops: [{ pos: 0, color: "FFD600" }, { pos: 100, color: a.accent }] }} />
          <shape kind="diamond" x={410} y={235} w={80} h={80} fill="9B7FBF" />
          <shape kind="chevron" x={520} y={245} w={150} h={60} fill="EDE9F6" stroke={{ color: a.ink, width: 1.5 }} />
          <ring cx={120} cy={430} r={50} thickness={12} value={a.progress} color="0FAF1A" />
          <text x={70} y={410} w={100} h={40} size={18} align="center" valign="middle" color={a.ink}>
            {`${Math.round(a.progress * 100)}%`}
          </text>
          <chart
            type="bar"
            x={230}
            y={370}
            w={460}
            h={250}
            colors={[a.ink, a.accent]}
            data={[
              { name: "Plan", labels: ["Q1", "Q2", "Q3", "Q4"], values: [12, 18, 22, 30] },
              { name: "Actual", labels: ["Q1", "Q2", "Q3", "Q4"], values: [10, 19, 25, 27] },
            ]}
          />
        </slide>
      </presentation>
    ),
    code: () => `<slide background="FFFFFF">
  <path d="M820 0 C900 160 760 330 840 470 … Z"
        fill={{ type: "linear", angle: 45, stops: [{ pos: 0, color: "FFD600" }, { pos: 100, color: "FF7A00" }] }} />
  <text x={60} y={60} w={700} h={60} size={40}>Hello from <span color="FF7A00">JSX</span></text>
  <shape x={60} y={220} w={180} h={110} radius={14} fill="24074A" shadow={{ blur: 18 }} />
  <ring cx={120} cy={430} r={50} thickness={12} value={0.72} color="0FAF1A" />
  <chart type="bar" x={230} y={370} w={460} h={250} data={[…]} />
</slide>`,
  },
  {
    id: "full-deck",
    group: "Decks",
    name: "Enavia launch deck",
    description: "The complete deck rebuilt from the reference PDF, generated from sample data.",
    args: {},
    render: () => <EnaviaDeck />,
    code: () => `import { renderToFile } from "pptx-jsx";
import { EnaviaDeck } from "./deck";

await renderToFile(<EnaviaDeck />, "enavia-launch.pptx");`,
  },
];
