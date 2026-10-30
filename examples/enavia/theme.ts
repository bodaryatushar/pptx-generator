import type { LinearGradient } from "pptx-jsx";

export const theme = {
  font: "Poppins",
  fontMedium: "Poppins Medium",
  color: {
    ink: "24074A", // brand deep purple
    text: "374151",
    muted: "6B7280",
    rule: "CDBFD8",
    rowRule: "EEEEF2",
    lavender: "E6E3EE",
    canvas: "E4E4EA",
    white: "FFFFFF",
    orange: "FF7A00",
    yellow: "FFD600",
    lilac: "9B7FBF",
    today: "FF7A00",
  },
  status: {
    notStarted: "9CA3AF",
    inProgress: "3B82F6",
    almostDelayed: "FACC15",
    delayed: "F4511E",
    completed: "0FAF1A",
  },
};

/** Logo image (file path in Node, URL/data URI in the browser). Unset = text wordmark. */
export const brand: { logo?: string } = {};

export const sunset: LinearGradient = {
  type: "linear",
  angle: 45,
  stops: [
    { pos: 0, color: theme.color.yellow },
    { pos: 100, color: theme.color.orange },
  ],
};

export const sunsetBand: LinearGradient = { ...sunset, angle: 15 };

export type Status = keyof typeof theme.status;
