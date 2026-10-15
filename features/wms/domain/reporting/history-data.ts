// Data historis WMS diambil dari sheet "REAL" pada workbook
// R-25-PRQ-JJS-RWMS_Reporting_WMS_2026 - Newest.xlsx (nilai input/hardcode di sheet tsb).
//
// prevDays  : REAL!B:AF  -> realisasi harian bulan sebelumnya (sumber "REAL MTD bulan lalu")
// prevTotal : REAL!AG    -> kolom "Total" bulan sebelumnya (sumber "REAL FM bulan lalu")
// y2026     : REAL!AL:AT -> total bulanan Jan..Sep 2026 (dasar "REAL YTD")
// y2025     : REAL!BA:BL -> total bulanan Jan..Des 2025 (dasar "GR YTD")
//
// Berlaku untuk laporan periode `reportPeriodKey`. Untuk periode lain, data bulan
// lalu dihitung dari baris MASTER yang tersimpan di dashboard.

export type WmsHistoryEntry = {
  prevDays: readonly number[];
  prevTotal: number;
  y2025: readonly number[];
  y2026: readonly number[];
};

export const wmsHistory = {
  reportPeriodKey: "202610",
  lastClosedMonth2026: 9,
  witel: {
    "BALI": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,5,0,10,0], prevTotal: 209, y2026: [8,7,17,11,51,20,30,209,15], y2025: [0,0,1,0,0,1,0,3,0,0,1,0] },
    "JATIM BARAT": { prevDays: [1,1,1,0,0,0,0,1,1,1,0,0,0,1,1,0,2,2,0,1,2,4,2,5,3,0,2,0,1,3,0], prevTotal: 17, y2026: [11,16,5,7,13,13,9,17,35], y2025: [7,10,19,8,20,16,13,10,3,9,6,5] },
    "JATIM TIMUR": { prevDays: [1,1,1,3,1,0,6,1,1,0,1,4,1,1,3,3,2,6,2,0,5,9,3,6,12,6,1,4,8,9,0], prevTotal: 109, y2026: [12,17,10,14,12,9,53,109,101], y2025: [13,34,26,23,29,16,23,17,11,15,12,23] },
    "NUSA TENGGARA": { prevDays: [0,0,2,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,1,2,0,0], prevTotal: 1, y2026: [31,7,12,4,14,9,3,1,6], y2025: [3,6,10,11,10,9,7,1,1,5,2,8] },
    "SEMARANG JATENG UTARA": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,1,0,0,0,0,0,1,0,0,0,0,0,0,0], prevTotal: 2, y2026: [1,4,4,2,2,3,5,2,3], y2025: [3,7,4,5,8,6,6,4,7,6,2,5] },
    "SOLO JATENG TIMUR": { prevDays: [1,0,0,0,0,0,0,0,0,0,1,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0], prevTotal: 0, y2026: [1,0,1,3,1,0,3,0,4], y2025: [4,13,8,12,25,4,4,13,7,8,3,7] },
    "SURAMADU": { prevDays: [0,0,2,0,0,0,0,5,0,0,2,0,0,6,0,1,0,0,1,2,8,0,0,4,3,0,1,1,3,7,0], prevTotal: 77, y2026: [3,22,5,20,16,21,63,77,46], y2025: [21,18,28,31,37,28,25,12,16,15,21,6] },
    "YOGYA JATENG SELATAN": { prevDays: [0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,2,0], prevTotal: 5, y2026: [2,4,9,3,12,5,9,5,4], y2025: [0,2,8,5,7,7,5,3,6,1,4,1] },
  } as Record<string, WmsHistoryEntry>,
  telda: {
    "BANJARNEGARA": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], prevTotal: 0, y2026: [0,0,0,0,0,0,0,0,0], y2025: [0,0,0,0,0,1,1,0,0,0,0,0] },
    "BANTUL": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], prevTotal: 0, y2026: [0,0,0,0,0,0,0,0,0], y2025: [0,0,0,0,0,0,0,0,0,0,0,0] },
    "CILACAP": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], prevTotal: 0, y2026: [0,0,0,0,0,0,3,0,0], y2025: [0,0,1,0,1,1,0,0,0,0,0,0] },
    "GUNUNG KIDUL": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], prevTotal: 0, y2026: [0,0,0,0,0,0,0,0,0], y2025: [0,0,0,0,1,0,0,0,0,0,0,0] },
    "KEBUMEN": { prevDays: [0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3], prevTotal: 4, y2026: [0,0,0,0,2,0,2,4,1], y2025: [0,0,0,0,1,1,0,0,0,0,0,0] },
    "MAGELANG": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], prevTotal: 0, y2026: [0,0,2,0,0,0,0,0,0], y2025: [0,1,1,1,1,0,0,1,0,0,1,0] },
    "MEA YOGYAKARTA": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], prevTotal: 0, y2026: [1,1,3,2,0,3,0,0,0], y2025: [0,1,0,1,0,0,0,0,2,0,0,0] },
    "MUNTILAN": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], prevTotal: 0, y2026: [0,0,0,0,0,0,0,0,0], y2025: [0,0,0,0,0,0,0,0,0,0,0,0] },
    "PURBALINGGA": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], prevTotal: 0, y2026: [1,0,3,1,4,0,2,0,1], y2025: [0,0,1,0,0,0,1,0,1,0,1,0] },
    "PURWOKERTO": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], prevTotal: 0, y2026: [0,3,1,0,3,0,1,0,2], y2025: [0,0,1,0,0,1,0,0,0,0,0,0] },
    "PURWOREJO": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0], prevTotal: 1, y2026: [0,0,0,0,0,1,1,1,0], y2025: [0,0,0,1,0,1,1,0,1,1,0,1] },
    "SLEMAN": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], prevTotal: 0, y2026: [0,0,0,0,3,0,0,0,0], y2025: [0,0,0,0,1,0,0,2,0,0,0,0] },
    "TEMANGGUNG": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], prevTotal: 0, y2026: [0,0,0,0,0,1,0,0,0], y2025: [0,0,1,0,0,1,1,0,1,0,0,0] },
    "WONOSOBO": { prevDays: [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0], prevTotal: 0, y2026: [0,0,0,0,0,0,0,0,0], y2025: [0,0,3,2,2,1,1,0,1,0,2,0] },
  } as Record<string, WmsHistoryEntry>,
};

// Telda (14 baris REAL) -> nama TELKOM_DAERAH di MASTER (kolom K).
// MEA YOGYAKARTA di MASTER tertulis "INNER - YOGYA JATENG SELATAN" (lihat Report Sementara!AZ).
export const wmsTeldaMasterNames: Record<string, readonly string[]> = {
  "BANJARNEGARA": ["BANJARNEGARA"],
  "BANTUL": ["BANTUL"],
  "CILACAP": ["CILACAP"],
  "GUNUNG KIDUL": ["GUNUNG KIDUL"],
  "KEBUMEN": ["KEBUMEN"],
  "MAGELANG": ["MAGELANG"],
  "MEA YOGYAKARTA": ["INNER - YOGYA JATENG SELATAN", "MEA YOGYAKARTA"],
  "MUNTILAN": ["MUNTILAN"],
  "PURBALINGGA": ["PURBALINGGA"],
  "PURWOKERTO": ["PURWOKERTO"],
  "PURWOREJO": ["PURWOREJO"],
  "SLEMAN": ["SLEMAN"],
  "TEMANGGUNG": ["TEMANGGUNG"],
  "WONOSOBO": ["WONOSOBO"],
};

// Baris "BRANCH" pada laporan = penjumlahan beberapa telda
// (Report Sementara!F22:F29 dan kolom AO/AQ/AT).
export const wmsBranchMembers: Record<string, readonly string[]> = {
  "BANTUL": ["BANTUL", "GUNUNG KIDUL"],
  "CILACAP": ["CILACAP"],
  "KEBUMEN": ["KEBUMEN", "PURWOREJO"],
  "MAGELANG": ["MAGELANG", "MUNTILAN"],
  "PURBALINGGA": ["PURBALINGGA", "BANJARNEGARA"],
  "PURWOKERTO": ["PURWOKERTO"],
  "WONOSOBO": ["WONOSOBO", "TEMANGGUNG"],
  "YOGYAKARTA": ["MEA YOGYAKARTA", "SLEMAN"],
};
