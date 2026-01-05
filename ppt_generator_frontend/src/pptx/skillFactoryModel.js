const MONTHS_SHORT_EN_GB = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

function pad2(n) {
  return String(n).padStart(2, "0");
}

// PUBLIC_INTERFACE
export function formatDateRangeDDMonYYYY(startISO, endISO) {
  /** This is a public function. */
  const fmt = (iso) => {
    if (!iso) return "";
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return "";
    const dd = pad2(d.getDate());
    const mon = MONTHS_SHORT_EN_GB[d.getMonth()] ?? "";
    const yyyy = d.getFullYear();
    return `${dd} ${mon} ${yyyy}`;
  };

  const a = fmt(startISO);
  const b = fmt(endISO);
  if (!a || !b) return "";
  return `${a} – ${b}`;
}

// PUBLIC_INTERFACE
export function coerceBulletLines(text) {
  /** This is a public function. */
  return String(text ?? "")
    .split(/\r?\n/g)
    .map((l) => l.trim())
    .map((l) => l.replace(/^[-•]\s+/, "").trim())
    .filter(Boolean);
}

// PUBLIC_INTERFACE
export function createDefaultTeamTable() {
  /** This is a public function. */
  return {
    columns: ["Name", "Role", "Location"],
    rows: [
      ["", "", ""],
      ["", "", ""],
      ["", "", ""],
    ],
  };
}

// PUBLIC_INTERFACE
export function addTeamRow(table) {
  /** This is a public function. */
  const cols = table?.columns?.length ?? 0;
  const nextRows = [...(table?.rows ?? [])];
  nextRows.push(Array.from({ length: cols }, () => ""));
  return { ...table, rows: nextRows };
}

// PUBLIC_INTERFACE
export function removeTeamRow(table, rowIndex) {
  /** This is a public function. */
  const nextRows = [...(table?.rows ?? [])];
  if (rowIndex < 0 || rowIndex >= nextRows.length) return table;
  nextRows.splice(rowIndex, 1);
  return { ...table, rows: nextRows.length ? nextRows : [Array.from({ length: table.columns.length }, () => "")] };
}

// PUBLIC_INTERFACE
export function addTeamColumn(table, headerLabel = "") {
  /** This is a public function. */
  const nextColumns = [...(table?.columns ?? []), headerLabel];
  const nextRows = (table?.rows ?? []).map((r) => [...r, ""]);
  return { ...table, columns: nextColumns, rows: nextRows.length ? nextRows : [Array.from({ length: nextColumns.length }, () => "")] };
}

// PUBLIC_INTERFACE
export function removeTeamColumn(table, colIndex) {
  /** This is a public function. */
  const nextColumns = [...(table?.columns ?? [])];
  if (colIndex < 0 || colIndex >= nextColumns.length) return table;
  if (nextColumns.length <= 1) return table;

  nextColumns.splice(colIndex, 1);
  const nextRows = (table?.rows ?? []).map((r) => {
    const nr = [...r];
    nr.splice(colIndex, 1);
    return nr;
  });

  return { ...table, columns: nextColumns, rows: nextRows };
}

// PUBLIC_INTERFACE
export function updateTeamCell(table, rowIndex, colIndex, value) {
  /** This is a public function. */
  const nextRows = (table?.rows ?? []).map((r) => [...r]);
  if (!nextRows[rowIndex]) return table;
  if (colIndex < 0 || colIndex >= (table?.columns?.length ?? 0)) return table;
  nextRows[rowIndex][colIndex] = value;
  return { ...table, rows: nextRows };
}

// PUBLIC_INTERFACE
export function updateTeamColumnHeader(table, colIndex, value) {
  /** This is a public function. */
  const nextColumns = [...(table?.columns ?? [])];
  if (colIndex < 0 || colIndex >= nextColumns.length) return table;
  nextColumns[colIndex] = value;
  return { ...table, columns: nextColumns };
}

// PUBLIC_INTERFACE
export function createDefaultSkillFactoryForm() {
  /** This is a public function. */
  return {
    factoryName: "Java", // appended after 'Digital Applications:'
    sprintNumber: "14",
    sprintStartISO: "",
    sprintEndISO: "",
    projectHighlightsText: "",
    projectLowlightsText: "",
    teamTable: createDefaultTeamTable(),
    smeName: "",
    keyActivitiesCompletedText: "",
    keyActivitiesPlannedText: "",
  };
}
