import {
  BudgetSlide,
  CoverSlide,
  MessageSlide,
  SectionSlide,
  TaskProgressSlide,
  TaskTableSlides,
  ThankYouSlide,
  TimelineSlides,
  type Task,
  type TimelineItem,
} from "./components.js";
import { theme } from "./theme.js";

// Sample data based on the reference deck. In an app this comes from your API.
export const LOREM =
  "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua.";

export const departments: { name: string; tasks: Task[]; timeline: TimelineItem[] }[] = [
  {
    name: "BP&A",
    tasks: [
      { id: "1", title: "Develop patient flow and assess opportunities", comment: "Please make sure to mark the point where we lose patients for later analysis and strategising.", progress: 100, status: "completed" },
      { id: "2", title: "Identify local trial requirements for regulatory submission", comment: LOREM, progress: 100, status: "completed" },
    ],
    timeline: [],
  },
  {
    name: "Clinical",
    tasks: [
      { id: "1", title: "Develop patient flow and assess opportunities", comment: "Please make sure to mark the point where we lose patients for later analysis and strategising.", progress: 100, status: "completed" },
      { id: "2", title: "Create TPP incl upside/downside scenarios", comment: LOREM, progress: 100, status: "completed" },
      { id: "3", title: "Develop life cycle management plan", comment: LOREM, progress: 88, status: "inProgress" },
      { id: "4", title: "Identify local trial requirements for regulatory submission", comment: LOREM, progress: 100, status: "completed" },
      { id: "5", title: "Build scientific platform", progress: 100, status: "completed" },
      { id: "6", title: "Define segmentation and positioning", progress: 100, status: "completed" },
      { id: "7", title: "Marketing Materials", progress: 0, status: "almostDelayed" },
      { id: "7.1", title: "Develop brand book, branding guidance", progress: 100, status: "completed" },
      { id: "7.2", title: "Develop core campaign materials", progress: 0, status: "delayed" },
      { id: "7.3", title: "Develop branded materials", progress: 0, status: "almostDelayed" },
      { id: "7.4", title: "Develop core campaign materials", progress: 0, status: "delayed" },
      { id: "7.5", title: "Develop branded materials", progress: 0, status: "almostDelayed" },
      { id: "8", title: "Customer Segmentation and Targeting", progress: 0, status: "almostDelayed" },
      { id: "8.1", title: "Customer segmentation and Targeting studies", progress: 0, status: "delayed" },
      { id: "8.2", title: "Customer segmentation and Targeting studies", progress: 0, status: "almostDelayed" },
      { id: "9", title: "Identify patient groups and develop engagement strategy", progress: 100, status: "completed" },
      { id: "10", title: "Competition Analysis", progress: 0, status: "notStarted" },
      { id: "10.1", title: "Competitor wargames to challenge strategy", progress: 100, status: "completed" },
      { id: "10.2", title: "Competitor wargames to challenge strategy", progress: 0, status: "almostDelayed" },
      { id: "10.3", title: "Competitor landscape research", progress: 100, status: "completed" },
      { id: "10.4", title: "Competitor landscape research", progress: 72, status: "almostDelayed" },
      { id: "10.5", title: "Competitor resources research", progress: 81, status: "almostDelayed" },
      { id: "11", title: "Physicians insights disease awareness and unmet need (ad boards)", progress: 100, status: "completed" },
    ],
    timeline: [
      { kind: "group", label: "Define segmentation and positioning", start: "2025-06-01", end: "2028-10-15" },
      { kind: "group", label: "Marketing Materials", start: "2025-09-01", end: "2026-10-30" },
      { kind: "task", label: "Develop brand book, branding guidance", start: "2025-10-01", end: "2026-02-15" },
      { kind: "task", label: "Develop core campaign materials", start: "2026-04-01", end: "2026-05-15" },
      { kind: "task", label: "Develop branded materials", start: "2026-09-01", end: "2026-10-30" },
      { kind: "task", label: "Develop core campaign materials", start: "2026-05-15", end: "2026-08-15" },
      { kind: "task", label: "Develop branded materials", start: "2026-10-01", end: "2026-10-30" },
      { kind: "group", label: "Customer Segmentation and Targeting", start: "2025-10-01", end: "2026-10-30" },
      { kind: "task", label: "Customer segmentation and Targeting studies", start: "2026-01-01", end: "2026-03-01" },
      { kind: "task", label: "Customer segmentation and Targeting studies", start: "2026-10-01", end: "2026-10-30" },
      { kind: "group", label: "Competition Analysis", start: "2025-11-01", end: "2027-01-31" },
      { kind: "task", label: "Competitor wargames to challenge strategy", start: "2026-03-01", end: "2026-06-01" },
      { kind: "milestone", label: "Competitor wargames to challenge strategy", start: "2027-01-31" },
      { kind: "milestone", label: "Competitor landscape research", start: "2026-10-31" },
      { kind: "milestone", label: "Competitor resources research", start: "2026-10-31" },
    ],
  },
  {
    name: "Marketing",
    tasks: [
      { id: "1", title: "Develop life cycle management plan", comment: LOREM, progress: 88, status: "inProgress" },
      { id: "2", title: "Develop brand vision and values", comment: LOREM, progress: 100, status: "completed" },
      { id: "3", title: "Develop brand strategy", progress: 100, status: "completed" },
      { id: "4", title: "Testing brand name(s), logos, images", progress: 100, status: "completed" },
      { id: "5", title: "Capture and share outcomes and best practices from early launch markets to accelerate smaller / later ones", progress: 0, status: "notStarted" },
      { id: "6", title: "External Launch event", progress: 0, status: "notStarted" },
    ],
    timeline: [
      { kind: "group", label: "Develop brand strategy", start: "2025-06-01", end: "2028-10-15" },
      { kind: "group", label: "Capture and share outcomes and best practices from early launch markets to accelerate smaller / later ones", start: "2027-10-01", end: "2028-03-31" },
      { kind: "milestone", label: "External Launch event", start: "2027-10-01" },
    ],
  },
  { name: "Commercial Excellence", tasks: [
      { id: "1", title: "Identify local trial requirements for regulatory submission", comment: LOREM, progress: 100, status: "completed" },
      { id: "2", title: "Assess partnering options", progress: 100, status: "completed" },
      { id: "3", title: "Value based negotiation training", progress: 81, status: "delayed" },
    ], timeline: [] },
  { name: "Medical Affairs", tasks: [
      { id: "1", title: "Define SKU and pack size strategy", progress: 100, status: "completed" },
      { id: "2", title: "Medical landscape mapping: global KOLs / network / CoE", progress: 100, status: "completed" },
      { id: "3", title: "Develop core claims and evidence pack", progress: 100, status: "completed" },
    ], timeline: [] },
];

export const EnaviaDeck = () => (
  <presentation title="Neurostax Full Launch 2027" company="Enavia" font={theme.font} color={theme.color.text}>
    <CoverSlide title="Neurostax Full Launch 2027" meta="Version: 30th October 2026" />

    <BudgetSlide
      title="Strategic Imperatives budget"
      currency="EUR"
      total={1_500_000}
      spend={1_017_355}
      lines={[
        { label: "Prep the Market", amount: 640_000 },
        { label: "Prep the Brand", amount: 520_000 },
        { label: "Prep the Organisation", amount: 340_000 },
      ]}
    />
    <BudgetSlide
      title="Categories budget"
      currency="EUR"
      total={1_500_000}
      spend={1_017_355}
      lines={[
        { label: "Access Strategy", amount: 180_000 },
        { label: "Regulatory Dossier Submission", amount: 160_000 },
        { label: "Define Vision", amount: 60_000 },
        { label: "Brand Positioning", amount: 110_000 },
        { label: "Core branded Campaign", amount: 210_000 },
        { label: "Customer Insights", amount: 95_000 },
        { label: "Brand Messaging", amount: 85_000 },
        { label: "Market Insights", amount: 120_000 },
        { label: "Training", amount: 70_000 },
        { label: "Financial Planning", amount: 45_000 },
        { label: "Payor Engagement Plan", amount: 130_000 },
        { label: "Launch Plan", amount: 135_000 },
      ]}
    />

    <TaskProgressSlide
      title="Neurostax Full Launch 2027 Task Progress"
      byStatus={[
        { label: "Not Started", value: 32, color: theme.status.notStarted },
        { label: "In Progress", value: 1, color: theme.status.inProgress },
        { label: "Almost delayed", value: 29, color: theme.status.almostDelayed },
        { label: "Delayed", value: 82, color: theme.status.delayed },
        { label: "Completed", value: 42, color: theme.status.completed },
      ]}
      readiness={{ completed: 10, remaining: 5 }}
    />

    {departments.map((d) => (
      <>
        <SectionSlide title={d.name} subtitle="Key Initiatives and Accomplishments" meta="Q4 2026 and Q1 2027" />
        <TimelineSlides
          title={`${d.name} Key Initiative & Accomplishment Timeline`}
          from={2026}
          to={2030}
          today="2026-10-30"
          items={d.timeline}
        />
        <TaskTableSlides title={`${d.name} department tasks`} tasks={d.tasks} />
      </>
    ))}

    <MessageSlide title="Risks" message="No risks found" />
    <ThankYouSlide />
  </presentation>
);

