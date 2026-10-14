export interface Indicator {
  name: string;
  weight: number;
}

/** [target, realisasi, ach (%), skor] */
export type Cell = [number, number, number, number];

export interface Row {
  no: number;
  witel: string;
  telda: string;
  v: Cell[];
  total: number;
  rank: number;
}

export interface Dataset {
  indicators: Indicator[];
  rows: Row[];
  warn?: string;
}

export type Tab = "data" | "viz";
export type CaptureMode = Tab;
