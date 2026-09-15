import path from "node:path";
import zlib from "node:zlib";
import JSZip from "jszip";
import XlsxPopulate from "xlsx-populate";

type Duty = "I" | "AA" | "S";

type TimetableRow = {
  date: Date;
  time: string;
  code: string;
  paper: string;
  subjectName: string;
  moa: string;
  duration: string;
  remarks: string;
};

type ExamRow = {
  date: string;
  session: string;
  paper: string;
  subjectName: string;
  normal: number;
  aa: number;
};

type Group = {
  date: string;
  session: string;
  papers: string[];
  subjects: string[];
};

type Paper = {
  column: number;
  unavailabilityColumn: number;
  date: string;
  code: string;
  subject: string;
  session: string;
  required: Record<Duty, number>;
  normalMinutes: number;
  aaMinutes: number;
};

type Invigilator = {
  row: number;
  name: string;
  subjectTokens: Set<string>;
  unavailableColumns: Set<number>;
  assignments: Map<number, Duty>;
  counts: Record<Duty, number>;
  minutes: number;
  amCount: number;
  pmCount: number;
  dateSessions: Map<string, Set<string>>;
};

const TEMPLATE_DIR = path.join(process.cwd(), "templates");
const EXAM_TEMPLATE = path.join(TEMPLATE_DIR, "Exam Dates Template.xlsx");
const MAIN_EXAM_START = new Date(2025, 9, 27);
const EXCLUDED_PAPERS = new Set(["9539/03"]);
const YELLOW = "FFFFFF00";
const BLACK = "FF000000";
const ORANGE = "FFFFC0C0";
const HEADER_GREY = "FFD9D9D9";
const RED = "FFFF0000";
const CENTER = "center";

function pad(number: number) {
  return String(number).padStart(2, "0");
}

function formatDate(date: Date) {
  return `${pad(date.getDate())}/${pad(date.getMonth() + 1)}/${date.getFullYear()}`;
}

function parseDate(value: string) {
  const match = value.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!match) return null;
  return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
}

function cellText(value: unknown) {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return formatDate(value);
  if (typeof value === "number") return Number.isInteger(value) ? String(value) : String(value);
  return String(value).replace(/\s+/g, " ").trim();
}

function clean(value: unknown) {
  return cellText(value).replace(/\s+/g, " ").trim();
}

function asInt(value: unknown) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Math.round(value);
  const parsed = Number(String(value).trim());
  return Number.isFinite(parsed) ? Math.round(parsed) : 0;
}

function convertToMinutes(value: unknown) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return Math.round(value < 1 ? value * 1440 : value);
  const text = String(value).trim();
  const match = text.match(/^(\d{1,2}):(\d{2})$/);
  if (match) return Number(match[1]) * 60 + Number(match[2]);
  const parsed = Number(text);
  return Number.isFinite(parsed) ? Math.round(parsed) : 0;
}

function sessionFromTime(time: string) {
  const match = time.match(/^\s*(\d{1,2}):/);
  return match && Number(match[1]) >= 12 ? "PM" : "AM";
}

function durationToMinutes(duration: string) {
  const match = duration.trim().match(/^(\d+):(\d{2})$/);
  if (!match) return 0;
  return Number(match[1]) * 60 + Number(match[2]);
}

function defaultAaMinutes(normalMinutes: number) {
  if (normalMinutes <= 0) return normalMinutes;
  return Math.ceil((normalMinutes * 1.25) / 5) * 5;
}

function timeRangeEnd(timeText: string) {
  return timeText.match(/-\s*(\d{1,2}:\d{2})\s*$/)?.[1] ?? "";
}

function addMinutesToClock(clockText: string, minutes: number) {
  const match = clockText.match(/^(\d{1,2}):(\d{2})$/);
  if (!match) return clockText;
  const date = new Date(2000, 0, 1, Number(match[1]), Number(match[2]) + minutes);
  return `${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function columnName(columnNumber: number) {
  let name = "";
  let column = columnNumber;
  while (column > 0) {
    column -= 1;
    name = String.fromCharCode(65 + (column % 26)) + name;
    column = Math.floor(column / 26);
  }
  return name;
}

function decodePdfString(value: string) {
  let text = value;
  if (text.startsWith("(") && text.endsWith(")")) text = text.slice(1, -1);
  let output = "";
  for (let index = 0; index < text.length; index++) {
    const char = text[index];
    if (char === "\\" && index + 1 < text.length) {
      const next = text[++index];
      if (next === "n") output += "\n";
      else if (next === "r") output += "\r";
      else if (next === "t") output += "\t";
      else if (next === "b") output += "\b";
      else if (next === "f") output += "\f";
      else output += next;
    } else {
      output += char;
    }
  }
  return output;
}

function inflateMaybe(bytes: Buffer) {
  const candidates = [bytes, bytes.length > 6 ? bytes.subarray(2, bytes.length - 4) : null].filter(Boolean) as Buffer[];
  for (const candidate of candidates) {
    try {
      return zlib.inflateSync(candidate);
    } catch {
      try {
        return zlib.inflateRawSync(candidate);
      } catch {
        // Try the next stream wrapper.
      }
    }
  }
  return null;
}

function pdfTextItems(pdfBuffer: Buffer) {
  const latin = pdfBuffer.toString("latin1");
  const items: Array<{ stream: number; x: number; y: number; text: string }> = [];
  let position = 0;
  let streamNumber = 0;

  while (position >= 0) {
    const streamIndex = latin.indexOf("stream", position);
    if (streamIndex < 0) break;
    let dataStart = streamIndex + "stream".length;
    if (pdfBuffer[dataStart] === 13 && pdfBuffer[dataStart + 1] === 10) dataStart += 2;
    else if (pdfBuffer[dataStart] === 10 || pdfBuffer[dataStart] === 13) dataStart += 1;
    const endIndex = latin.indexOf("endstream", dataStart);
    if (endIndex < 0) break;
    let dataEnd = endIndex;
    while (dataEnd > dataStart && (pdfBuffer[dataEnd - 1] === 10 || pdfBuffer[dataEnd - 1] === 13)) dataEnd--;
    const dictStart = latin.lastIndexOf("<<", streamIndex);
    const dictText = dictStart >= 0 ? latin.slice(dictStart, streamIndex) : "";
    let streamBytes = pdfBuffer.subarray(dataStart, dataEnd);
    if (dictText.includes("/FlateDecode")) {
      const inflated = inflateMaybe(streamBytes);
      if (!inflated) {
        position = endIndex + "endstream".length;
        continue;
      }
      streamBytes = inflated;
    }
    const text = streamBytes.toString("latin1");
    if (/\bBT\b/.test(text) && (/\bTJ\b/.test(text) || /\bTj\b/.test(text))) {
      streamNumber++;
      const pattern = /(?:-?\d+(?:\.\d+)?\s+){4}(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+Tm([\s\S]*?)(?:ET)/g;
      for (const match of text.matchAll(pattern)) {
        const parts = [...match[3].matchAll(/\((?:\\.|[^\\)])*\)/g)].map((part) => decodePdfString(part[0]));
        const value = parts.join("").replace(/\s+/g, " ").trim();
        if (value) items.push({ stream: streamNumber, x: Number(match[1]), y: Number(match[2]), text: value });
      }
    }
    position = endIndex + "endstream".length;
  }
  return items;
}

function firstCell(cells: Array<{ x: number; text: string }>, minX: number, maxX: number) {
  return cells.find((cell) => cell.x >= minX && cell.x < maxX)?.text ?? "";
}

function joinCell(cells: Array<{ x: number; text: string }>, minX: number, maxX: number) {
  return cells
    .filter((cell) => cell.x >= minX && cell.x < maxX)
    .sort((a, b) => a.x - b.x)
    .map((cell) => cell.text)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

function timetableRowsFromPdf(pdfBuffer: Buffer) {
  const groups = new Map<string, Array<{ x: number; text: string }>>();
  for (const item of pdfTextItems(pdfBuffer)) {
    if (item.x < 45 || item.x > 540 || item.y <= 50) continue;
    const key = `${item.stream}|${Math.round(item.y * 10) / 10}`;
    const cells = groups.get(key) ?? [];
    cells.push({ x: item.x, text: item.text });
    groups.set(key, cells);
  }

  const rows: TimetableRow[] = [];
  for (const cells of groups.values()) {
    cells.sort((a, b) => a.x - b.x);
    const dateText = firstCell(cells, 45, 80);
    const code = firstCell(cells, 125, 153);
    const date = parseDate(dateText);
    if (!date || !/^\d{4}$/.test(code)) continue;
    rows.push({
      date,
      time: firstCell(cells, 80, 125),
      code,
      paper: firstCell(cells, 153, 180),
      subjectName: joinCell(cells, 180, 290),
      moa: firstCell(cells, 290, 380),
      duration: firstCell(cells, 380, 420),
      remarks: joinCell(cells, 420, 540)
    });
  }
  return rows;
}

function applyBorder(range: any) {
  range.style("border", true);
}

function setFill(cellOrRange: any, color?: string) {
  if (!color) cellOrRange.style("fill", undefined);
  else cellOrRange.style("fill", { type: "solid", color: { rgb: color } });
}

function setText(sheet: any, row: number, column: number, value: unknown) {
  sheet.cell(row, column).value(value === undefined ? "" : value);
}

async function normalizeXlsxPackage(buffer: Buffer) {
  const zip = await JSZip.loadAsync(buffer);

  const workbookFile = zip.file("xl/workbook.xml");
  if (workbookFile) {
    let workbookXml = await workbookFile.async("string");
    const rootMatch = workbookXml.match(/<workbook\b[^>]*>/);
    if (rootMatch && workbookXml.includes("r:id") && !rootMatch[0].includes("xmlns:r=")) {
      workbookXml = workbookXml.replace(
        /<workbook\b/,
        '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
      );
      zip.file("xl/workbook.xml", workbookXml);
    }
  }

  const contentTypesFile = zip.file("[Content_Types].xml");
  if (contentTypesFile) {
    let contentTypesXml = await contentTypesFile.async("string");
    const worksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    const worksheetFiles = Object.keys(zip.files).filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name));
    for (const worksheet of worksheetFiles) {
      const partName = `/${worksheet}`;
      if (!contentTypesXml.includes(`PartName="${partName}"`)) {
        contentTypesXml = contentTypesXml.replace(
          "</Types>",
          `<Override PartName="${partName}" ContentType="${worksheetContentType}"/></Types>`
        );
      }
    }
    zip.file("[Content_Types].xml", contentTypesXml);
  }

  const stylesFile = zip.file("xl/styles.xml");
  if (stylesFile) {
    const stylesXml = (await stylesFile.async("string")).replaceAll("<fill/>", "<fill><patternFill/></fill>");
    zip.file("xl/styles.xml", stylesXml);
  }

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

async function outputWorkbook(workbook: any) {
  const output = await workbook.outputAsync() as Buffer;
  return normalizeXlsxPackage(output);
}

function usedLastRow(sheet: any) {
  const range = sheet.usedRange();
  return range ? range.endCell().rowNumber() : 1;
}

function usedLastColumn(sheet: any) {
  const range = sheet.usedRange();
  return range ? range.endCell().columnNumber() : 1;
}

async function readSubjectCodes(buffer: Buffer) {
  const workbook = await XlsxPopulate.fromDataAsync(buffer);
  const sheet = workbook.sheet(0);
  const codes = new Set<string>();
  for (let row = 2; row <= usedLastRow(sheet); row++) {
    const value = clean(sheet.cell(row, 1).value());
    if (/^\d{4}$/.test(value)) codes.add(value);
  }
  if (codes.size === 0) throw new Error("No four-digit subject codes found in the selected workbook.");
  return codes;
}

export async function createExamDateSheet(subjectCodesBuffer: Buffer, timetablePdfBuffer: Buffer) {
  const subjectCodes = await readSubjectCodes(subjectCodesBuffer);
  const allRows = timetableRowsFromPdf(timetablePdfBuffer);
  if (allRows.length === 0) throw new Error("No timetable rows could be extracted from the PDF. The PDF may be scanned or use an unsupported encoding.");

  const rows = allRows
    .filter((row) => row.date >= MAIN_EXAM_START && subjectCodes.has(row.code) && row.moa === "WRITTEN" && !EXCLUDED_PAPERS.has(`${row.code}/${row.paper}`.toUpperCase()))
    .sort((a, b) => a.date.getTime() - b.date.getTime() || a.time.localeCompare(b.time) || a.code.localeCompare(b.code) || a.paper.localeCompare(b.paper));
  if (rows.length === 0) throw new Error("No matching written exam rows found. Check that the subject codes workbook and timetable PDF are for the same exam year.");

  const workbook = await XlsxPopulate.fromFileAsync(EXAM_TEMPLATE);
  const sheet = workbook.sheet("Exam Dates");
  const lastRow = Math.max(3, rows.length + 3);
  for (let row = 4; row <= Math.max(usedLastRow(sheet), lastRow); row++) {
    for (let column = 1; column <= 23; column++) {
      sheet.cell(row, column).value(null);
      sheet.cell(row, column).formula(null);
    }
    if (row <= lastRow) applyBorder(sheet.range(`A${row}:W${row}`));
  }

  rows.forEach((row, index) => {
    const outRow = index + 4;
    const dateText = formatDate(row.date);
    const paperText = `${row.code}/${row.paper}`;
    const normalMinutes = durationToMinutes(row.duration);
    const aaMinutes = defaultAaMinutes(normalMinutes);
    const endTime = timeRangeEnd(row.time);
    const ecEnd = addMinutesToClock(endTime, aaMinutes - normalMinutes);
    setText(sheet, outRow, 1, dateText);
    setText(sheet, outRow, 2, row.time);
    setText(sheet, outRow, 3, row.code);
    setText(sheet, outRow, 4, row.paper);
    setText(sheet, outRow, 5, row.subjectName);
    setText(sheet, outRow, 7, row.moa);
    setText(sheet, outRow, 8, row.duration);
    setText(sheet, outRow, 9, row.remarks);
    sheet.cell(outRow, 14).formula(`SUM(J${outRow}:M${outRow})`);
    setText(sheet, outRow, 15, ecEnd);
    setText(sheet, outRow, 19, dateText);
    setText(sheet, outRow, 20, paperText);
    setText(sheet, outRow, 21, sessionFromTime(row.time));
    setText(sheet, outRow, 22, normalMinutes);
    setText(sheet, outRow, 23, aaMinutes);
  });

  return outputWorkbook(workbook);
}

function joinDeduped(values: string[]) {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const value of values) {
    const text = value.trim();
    const key = text.replace(/\s+/g, " ").toUpperCase();
    if (!text || seen.has(key)) continue;
    seen.add(key);
    result.push(text);
  }
  return result.join(", ");
}

function collectExamRows(examSheet: any) {
  const rows: ExamRow[] = [];
  const hasV3DurationColumns = clean(examSheet.cell(4, 19).value()).toLowerCase() === "session";
  for (let row = 4; row <= usedLastRow(examSheet); row++) {
    const sourceDate = clean(examSheet.cell(row, 1).value());
    const timeText = clean(examSheet.cell(row, 2).value());
    const code = clean(examSheet.cell(row, 3).value());
    const paperNo = clean(examSheet.cell(row, 4).value());
    const subjectName = clean(examSheet.cell(row, 5).value());
    const remarks = clean(examSheet.cell(row, 9).value());
    if (!sourceDate && !code && !paperNo && !subjectName) continue;
    const helperDate = hasV3DurationColumns ? sourceDate : clean(examSheet.cell(row, 19).value()) || sourceDate;
    const paper = hasV3DurationColumns ? `${code}/${paperNo}` : clean(examSheet.cell(row, 20).value()) || `${code}/${paperNo}`;
    const session = (hasV3DurationColumns ? clean(examSheet.cell(row, 19).value()) : clean(examSheet.cell(row, 21).value())).toUpperCase() || sessionFromTime(timeText);
    const normal = convertToMinutes(hasV3DurationColumns ? examSheet.cell(row, 20).value() : examSheet.cell(row, 22).value()) || convertToMinutes(examSheet.cell(row, 8).value());
    let aa = convertToMinutes(hasV3DurationColumns ? examSheet.cell(row, 21).value() : examSheet.cell(row, 23).value()) || normal;
    if (/admin\s+break/i.test(remarks) && normal) aa = Math.round((normal * 1.25) / 5) * 5;
    if (!helperDate || !paper || !session) continue;
    rows.push({ date: helperDate, session, paper, subjectName, normal, aa });
  }
  if (rows.length === 0) throw new Error("No usable exam rows found.");
  return rows;
}

function ensureCombinedExamDatesSheet(workbook: any) {
  const firstSheet = workbook.sheet(0);
  const existingExamDates = workbook.sheet("Exam Dates");
  if (existingExamDates && existingExamDates !== firstSheet) workbook.deleteSheet(existingExamDates);
  firstSheet.name("Exam Dates");

  for (const name of ["Unavailability", "Schedule", "Requirements"]) {
    const sheet = workbook.sheet(name);
    if (sheet) workbook.deleteSheet(sheet);
  }

  setText(firstSheet, 3, 16, "No. of Invigilators");
  setText(firstSheet, 3, 19, "Duration of Papers");
  setText(firstSheet, 3, 20, "");
  setText(firstSheet, 3, 21, "");
  setText(firstSheet, 3, 22, "");
  setText(firstSheet, 3, 23, "");
  setText(firstSheet, 4, 16, "Normal");
  setText(firstSheet, 4, 17, "AA/Prompter");
  setText(firstSheet, 4, 18, "Standby");
  setText(firstSheet, 4, 19, "Session");
  setText(firstSheet, 4, 20, "Normal");
  setText(firstSheet, 4, 21, "AA");
  setText(firstSheet, 4, 22, "");
  setText(firstSheet, 4, 23, "");

  let examRows = 0;
  for (let row = 5; row <= usedLastRow(firstSheet); row++) {
    const sourceDate = clean(firstSheet.cell(row, 1).value());
    const timeText = clean(firstSheet.cell(row, 2).value());
    const code = clean(firstSheet.cell(row, 3).value());
    const paperNo = clean(firstSheet.cell(row, 4).value());
    const subjectName = clean(firstSheet.cell(row, 5).value());
    if (!sourceDate && !code && !paperNo && !subjectName) continue;

    const normalMinutes = convertToMinutes(firstSheet.cell(row, 8).value());
    const existingAaMinutes = convertToMinutes(firstSheet.cell(row, 21).value()) || convertToMinutes(firstSheet.cell(row, 23).value());
    const aaMinutes = existingAaMinutes || defaultAaMinutes(normalMinutes);
    firstSheet.cell(row, 14).formula(`SUM(J${row}:M${row})`);
    setText(firstSheet, row, 19, sessionFromTime(timeText));
    setText(firstSheet, row, 20, normalMinutes);
    setText(firstSheet, row, 21, aaMinutes);
    setText(firstSheet, row, 22, "");
    setText(firstSheet, row, 23, "");
    examRows++;
  }

  if (examRows === 0) throw new Error("No exam rows found in the combined examination timetable input.");
  firstSheet.range("P3:R3").merged(true);
  firstSheet.range("S3:U3").merged(true);
  firstSheet.range("P3:U4").style("bold", true).style("fontFamily", "Aptos Narrow").style("fontSize", 11).style("horizontalAlignment", CENTER);
  applyBorder(firstSheet.range("P3:U4"));
  setFill(firstSheet.range("P3:U4"), HEADER_GREY);
  firstSheet.range(`S5:U${usedLastRow(firstSheet)}`).style("fontFamily", "Aptos Narrow").style("fontSize", 11).style("horizontalAlignment", CENTER);
  applyBorder(firstSheet.range(`S5:U${usedLastRow(firstSheet)}`));
  return firstSheet;
}

export async function createCombinedStepOneOutput(combinedTimetableBuffer: Buffer) {
  const workbook = await XlsxPopulate.fromDataAsync(combinedTimetableBuffer);
  ensureCombinedExamDatesSheet(workbook);
  const normalized = await outputWorkbook(workbook);
  return createUnavailabilitySheet(normalized);
}

function setGridBase(sheet: any, lastColumn: number, labelColumn: number, firstGroupColumn: number, includeSummary: boolean) {
  sheet.range("A1:AZ55").style("fontFamily", "Aptos Narrow").style("fontSize", 11);
  sheet.column("A").width(24.66);
  sheet.column("B").width(17.44);
  sheet.column("C").width(13);
  if (includeSummary) {
    for (let col = 4; col <= 10; col++) sheet.column(col).width(13);
    sheet.column(11).width(19.66);
  }
  sheet.column(labelColumn).width(13);
  for (let col = firstGroupColumn; col <= lastColumn; col++) sheet.column(col).width(18);
  [1, 2].forEach((row) => sheet.row(row).height(24));
  sheet.row(3).height(54);
  sheet.row(4).height(72);
  setFill(sheet.range(`A1:${includeSummary ? "J4" : "C4"}`), BLACK);
  sheet.range(`${columnName(labelColumn)}1:${columnName(labelColumn)}4`).style("bold", true).style("fontFamily", "Arial").style("fontSize", 9).style("wrapText", true).style("verticalAlignment", CENTER);
  sheet.range(`${columnName(firstGroupColumn)}1:${columnName(lastColumn)}4`).style("fontFamily", "Arial").style("fontSize", 9).style("wrapText", true).style("verticalAlignment", CENTER);
  sheet.range(`${columnName(firstGroupColumn)}1:${columnName(lastColumn)}1`).style("horizontalAlignment", CENTER);
  sheet.range(`${columnName(firstGroupColumn)}2:${columnName(lastColumn)}3`).style("horizontalAlignment", "left");
  sheet.range(`${columnName(firstGroupColumn)}4:${columnName(lastColumn)}4`).style("horizontalAlignment", CENTER);
  sheet.range("A5:C5").style("bold", true);
  if (includeSummary) sheet.range("D5:K5").style("bold", true).style("wrapText", true);
  applyBorder(sheet.range(`A5:${includeSummary ? "K" : "C"}55`));
  applyBorder(sheet.range(`${columnName(labelColumn)}1:${columnName(lastColumn)}4`));
}

export async function createUnavailabilitySheet(examDatesBuffer: Buffer) {
  const workbook = await XlsxPopulate.fromDataAsync(examDatesBuffer);
  const examSheet = workbook.sheet("Exam Dates");
  if (!examSheet) throw new Error("The input workbook must contain a sheet named 'Exam Dates'.");
  for (const name of ["Unavailability", "Schedule", "Requirements"]) {
    const sheet = workbook.sheet(name);
    if (sheet) workbook.deleteSheet(sheet);
  }
  const examRows = collectExamRows(examSheet);
  const groups: Group[] = [];
  for (const exam of examRows) {
    let group = groups.find((item) => item.date === exam.date && item.session === exam.session);
    if (!group) {
      group = { date: exam.date, session: exam.session, papers: [], subjects: [] };
      groups.push(group);
    }
    group.papers.push(exam.paper);
    group.subjects.push(exam.subjectName);
  }

  const unavailability = workbook.addSheet("Unavailability");
  const lastColumn = 4 + groups.length;
  setText(unavailability, 1, 4, "Date");
  setText(unavailability, 2, 4, "Session");
  setText(unavailability, 3, 4, "Subject Code");
  setText(unavailability, 4, 4, "Subject Name");
  setText(unavailability, 5, 1, "Name");
  setText(unavailability, 5, 2, "Teaching Subject");
  setText(unavailability, 5, 3, "Teaching classes");
  for (let row = 6; row <= 55; row++) setText(unavailability, row, 1, `Inv${row - 5}`);
  groups.forEach((group, index) => {
    const column = 5 + index;
    setText(unavailability, 1, column, group.date);
    setText(unavailability, 2, column, group.session);
    setText(unavailability, 3, column, group.papers.join(", "));
    setText(unavailability, 4, column, joinDeduped(group.subjects));
  });
  setGridBase(unavailability, lastColumn, 4, 5, false);
  if (lastColumn >= 6) {
    for (let col = 6; col <= lastColumn; col++) unavailability.column(col).width(13);
  }
  [1, 2].forEach((row) => unavailability.row(row).height(19.2));
  unavailability.row(3).height(43.2);
  unavailability.row(4).height(57.6);
  if (lastColumn >= 5) {
    applyBorder(unavailability.range(`E1:${columnName(lastColumn)}4`));
    setFill(unavailability.range(`E6:${columnName(lastColumn)}55`), YELLOW);
    applyBorder(unavailability.range(`E6:${columnName(lastColumn)}55`));
  }
  return outputWorkbook(workbook);
}

function normalizeSubject(value: string) {
  return value.replace(/\s+/g, " ").trim().toUpperCase().replaceAll("H1 ", "").replaceAll("H2 ", "").replace("(REVISED)", "").replace(/\s+/g, " ").trim();
}

function subjectTokens(value: string) {
  const tokens = new Set<string>();
  for (const part of value.split(",")) {
    const norm = normalizeSubject(part);
    if (!norm || norm === "OTHERS" || norm === "OTHER") continue;
    tokens.add(norm);
  }
  return tokens;
}

function teachesSubject(inv: Invigilator, subject: string) {
  for (const part of subject.split(",")) {
    const norm = normalizeSubject(part);
    if (norm && inv.subjectTokens.has(norm)) return true;
    if ((norm === "PE" && inv.subjectTokens.has("PHYSICAL EDUCATION")) || (norm === "PHYSICAL EDUCATION" && inv.subjectTokens.has("PE"))) return true;
  }
  return false;
}

function getDoubleDays(inv: Invigilator) {
  let count = 0;
  for (const sessions of inv.dateSessions.values()) if (sessions.has("AM") && sessions.has("PM")) count++;
  return count;
}

function dutyMinutes(paper: Paper, duty: Duty) {
  if (duty === "I") return paper.normalMinutes;
  if (duty === "AA") return paper.aaMinutes;
  return 0;
}

function addAssignment(inv: Invigilator, paper: Paper, duty: Duty) {
  inv.assignments.set(paper.column, duty);
  inv.counts[duty]++;
  inv.minutes += dutyMinutes(paper, duty);
  if (paper.session === "AM") inv.amCount++;
  if (paper.session === "PM") inv.pmCount++;
  if (!inv.dateSessions.has(paper.date)) inv.dateSessions.set(paper.date, new Set());
  inv.dateSessions.get(paper.date)!.add(paper.session);
}

function rebuildDateSessions(inv: Invigilator, paperByColumn: Map<number, Paper>) {
  inv.dateSessions = new Map();
  for (const column of inv.assignments.keys()) {
    const paper = paperByColumn.get(column)!;
    if (!inv.dateSessions.has(paper.date)) inv.dateSessions.set(paper.date, new Set());
    inv.dateSessions.get(paper.date)!.add(paper.session);
  }
}

function removeAssignment(inv: Invigilator, paper: Paper, duty: Duty, paperByColumn: Map<number, Paper>) {
  inv.assignments.delete(paper.column);
  inv.counts[duty]--;
  inv.minutes -= dutyMinutes(paper, duty);
  if (paper.session === "AM") inv.amCount--;
  if (paper.session === "PM") inv.pmCount--;
  rebuildDateSessions(inv, paperByColumn);
}

function canAssign(inv: Invigilator, paper: Paper) {
  return !inv.unavailableColumns.has(paper.unavailabilityColumn) && !inv.assignments.has(paper.column);
}

function findHeaderColumn(sheet: any, row: number, header: string) {
  for (let column = 1; column <= usedLastColumn(sheet); column++) {
    if (clean(sheet.cell(row, column).value()).toLowerCase() === header.toLowerCase()) return column;
  }
  throw new Error(`Could not find '${header}' in row ${row} of sheet '${sheet.name()}'.`);
}

function lastNamedRow(sheet: any) {
  for (let row = usedLastRow(sheet); row >= 6; row--) {
    if (clean(sheet.cell(row, 1).value())) return row;
  }
  throw new Error(`No invigilators found in '${sheet.name()}'.`);
}

function createScheduleSheet(workbook: any, unavailability: any) {
  const existing = workbook.sheet("Schedule");
  if (existing) workbook.deleteSheet(existing);
  const schedule = workbook.addSheet("Schedule");
  const unavLabelColumn = findHeaderColumn(unavailability, 3, "Subject Code");
  const unavFirst = unavLabelColumn + 1;
  const unavLast = usedLastColumn(unavailability);
  const lastColumn = 12 + Math.max(0, unavLast - unavFirst + 1);
  for (let row = 1; row <= 55; row++) for (let col = 1; col <= 3; col++) setText(schedule, row, col, unavailability.cell(row, col).value());
  setText(schedule, 5, 4, "no. of normal invigilator session");
  setText(schedule, 5, 5, "no. of AA invigilator session");
  setText(schedule, 5, 6, "total no. of normal + AA duty");
  setText(schedule, 5, 7, "no. of standby session");
  setText(schedule, 5, 8, "total no. of minutes");
  setText(schedule, 5, 9, "no. of AM session");
  setText(schedule, 5, 10, "no. of PM session");
  setText(schedule, 5, 11, "no. of days with double invigilation duty");
  setText(schedule, 1, 12, "Date");
  setText(schedule, 2, 12, "Session");
  setText(schedule, 3, 12, "Subject Code");
  setText(schedule, 4, 12, "Subject Name");
  for (let unavCol = unavFirst; unavCol <= unavLast; unavCol++) {
    const dest = 13 + (unavCol - unavFirst);
    for (let row = 1; row <= 4; row++) setText(schedule, row, dest, unavailability.cell(row, unavCol).value());
  }
  setGridBase(schedule, lastColumn, 12, 13, true);
  setFill(schedule.range("D1:K4"), BLACK);
  setFill(schedule.range("L1:L4"));
  schedule.range("D5:L55").style("wrapText", true);
  schedule.range(`M6:${columnName(lastColumn)}55`).style("horizontalAlignment", CENTER).style("verticalAlignment", CENTER);
  applyBorder(schedule.range(`K1:${columnName(lastColumn)}4`));
  return schedule;
}

function collectPapers(examDates: any, schedule: any, unavailability: any) {
  const reqByCode = new Map<string, any>();
  const hasV3DurationColumns = clean(examDates.cell(4, 19).value()).toLowerCase() === "session";
  for (let row = 4; row <= usedLastRow(examDates); row++) {
    const sourceCode = clean(examDates.cell(row, 3).value());
    const paperNo = clean(examDates.cell(row, 4).value());
    const code = hasV3DurationColumns && sourceCode && paperNo ? `${sourceCode}/${paperNo}` : clean(examDates.cell(row, 20).value());
    if (!code) continue;
    reqByCode.set(code, {
      normalRequired: asInt(examDates.cell(row, 16).value()),
      aaRequired: asInt(examDates.cell(row, 17).value()),
      standbyRequired: asInt(examDates.cell(row, 18).value()),
      normalMinutes: asInt(examDates.cell(row, hasV3DurationColumns ? 20 : 22).value()),
      aaMinutes: asInt(examDates.cell(row, hasV3DurationColumns ? 21 : 23).value())
    });
  }
  const scheduleFirst = findHeaderColumn(schedule, 3, "Subject Code") + 1;
  const unavFirst = findHeaderColumn(unavailability, 3, "Subject Code") + 1;
  const unavByKey = new Map<string, number>();
  for (let col = unavFirst; col <= usedLastColumn(unavailability); col++) {
    const key = `${clean(unavailability.cell(1, col).value())}|${clean(unavailability.cell(2, col).value()).toUpperCase()}|${clean(unavailability.cell(3, col).value())}`;
    if (key !== "||") unavByKey.set(key, col);
  }
  const papers: Paper[] = [];
  for (let col = scheduleFirst; col <= usedLastColumn(schedule); col++) {
    const codeText = clean(schedule.cell(3, col).value());
    if (!codeText) continue;
    const date = clean(schedule.cell(1, col).value());
    const session = clean(schedule.cell(2, col).value()).toUpperCase();
    const subject = clean(schedule.cell(4, col).value());
    let requiredI = 0;
    let requiredAA = 0;
    let requiredS = 0;
    let normalMinutes = 0;
    let aaMinutes = 0;
    for (const code of codeText.split(",")) {
      const trimmed = code.trim();
      if (!trimmed) continue;
      const info = reqByCode.get(trimmed);
      if (!info) throw new Error(`Could not find requirements for paper '${trimmed}'.`);
      requiredI += info.normalRequired;
      requiredAA += info.aaRequired;
      requiredS += info.standbyRequired;
      normalMinutes = Math.max(normalMinutes, info.normalMinutes);
      aaMinutes = Math.max(aaMinutes, info.aaMinutes);
    }
    if (aaMinutes === 0) aaMinutes = normalMinutes;
    const unavKey = `${date}|${session}|${codeText}`;
    papers.push({
      column: col,
      unavailabilityColumn: unavByKey.get(unavKey) ?? unavFirst + (col - scheduleFirst),
      date,
      code: codeText,
      subject,
      session,
      required: { I: requiredI, AA: requiredAA, S: requiredS },
      normalMinutes,
      aaMinutes
    });
  }
  return papers;
}

function collectInvigilators(schedule: any, unavailability: any) {
  const unavRowByName = new Map<string, number>();
  for (let row = 6; row <= lastNamedRow(unavailability); row++) {
    const name = clean(unavailability.cell(row, 1).value());
    if (name) unavRowByName.set(name, row);
  }
  const invigilators: Invigilator[] = [];
  for (let row = 6; row <= lastNamedRow(schedule); row++) {
    const name = clean(schedule.cell(row, 1).value());
    if (!name || !unavRowByName.has(name)) continue;
    const unavRow = unavRowByName.get(name)!;
    const unavailableColumns = new Set<number>();
    for (let col = 1; col <= usedLastColumn(unavailability); col++) {
      if (clean(unavailability.cell(unavRow, col).value()).toUpperCase() === "X") unavailableColumns.add(col);
    }
    const subjectText = clean(unavailability.cell(unavRow, 2).value()) || clean(schedule.cell(row, 2).value());
    invigilators.push({
      row,
      name,
      subjectTokens: subjectTokens(subjectText),
      unavailableColumns,
      assignments: new Map(),
      counts: { I: 0, AA: 0, S: 0 },
      minutes: 0,
      amCount: 0,
      pmCount: 0,
      dateSessions: new Map()
    });
  }
  return invigilators;
}

function buildTargets(papers: Paper[], invigilators: Invigilator[]) {
  const targets: Record<string, Map<string, number>> = {};
  for (const duty of ["I", "AA", "S", "NormalAA"]) {
    let total = 0;
    for (const paper of papers) total += duty === "NormalAA" ? paper.required.I + paper.required.AA : paper.required[duty as Duty];
    const base = Math.floor(total / invigilators.length);
    const extra = total % invigilators.length;
    targets[duty] = new Map(invigilators.map((inv, index) => [inv.name, base + (index < extra ? 1 : 0)]));
  }
  return targets;
}

function score(inv: Invigilator, paper: Paper, duty: Duty, targets: Record<string, Map<string, number>>, targetMinutes: number) {
  const normalAaTarget = targets.NormalAA.get(inv.name) ?? 0;
  const aaTarget = targets.AA.get(inv.name) ?? 0;
  const standbyTarget = targets.S.get(inv.name) ?? 0;
  const currentNormalAa = inv.counts.I + inv.counts.AA;
  let projectedNormalAa = currentNormalAa + (duty === "I" || duty === "AA" ? 1 : 0);
  let projectedAa = inv.counts.AA + (duty === "AA" ? 1 : 0);
  const projectedStandby = inv.counts.S + (duty === "S" ? 1 : 0);
  let normalAaBelowRank = 10000 - Math.max(0, normalAaTarget - currentNormalAa);
  let normalAaAboveAfter = Math.max(0, projectedNormalAa - normalAaTarget);
  let aaBelowRank = 10000 - Math.max(0, aaTarget - inv.counts.AA);
  let aaAboveAfter = Math.max(0, projectedAa - aaTarget);
  const standbyBelowRank = 10000 - Math.max(0, standbyTarget - inv.counts.S);
  const standbyAboveAfter = Math.max(0, projectedStandby - standbyTarget);
  const projectedMinutes = inv.minutes + dutyMinutes(paper, duty);
  const doubleAfter = inv.dateSessions.has(paper.date) && !inv.dateSessions.get(paper.date)!.has(paper.session) ? 1 : 0;
  const subjectConflict = teachesSubject(inv, paper.subject) ? 1 : 0;
  const minuteDelta = Math.round(Math.abs(projectedMinutes - targetMinutes));
  if (duty === "S") {
    normalAaBelowRank = 0;
    normalAaAboveAfter = 0;
    projectedNormalAa = 0;
    aaBelowRank = 0;
    aaAboveAfter = 0;
    projectedAa = 0;
  }
  return [
    subjectConflict * 5,
    normalAaBelowRank,
    normalAaAboveAfter * 40,
    projectedNormalAa,
    aaBelowRank,
    aaAboveAfter * 40,
    projectedAa,
    doubleAfter * 5,
    standbyBelowRank,
    standbyAboveAfter * 40,
    projectedStandby,
    minuteDelta,
    projectedMinutes
  ].map((number) => String(number).padStart(8, "0")).concat(inv.name).join("|");
}

function makeSchedule(papers: Paper[], invigilators: Invigilator[], targetMinutes: number) {
  const targets = buildTargets(papers, invigilators);
  const shortages: Array<{ code: string; duty: Duty; required: number; assigned: number }> = [];
  for (const duty of ["AA", "I", "S"] as Duty[]) {
    const ordered = papers
      .filter((paper) => paper.required[duty] > 0)
      .sort((a, b) => invigilators.filter((inv) => canAssign(inv, a)).length - invigilators.filter((inv) => canAssign(inv, b)).length || b.required[duty] - a.required[duty] || a.date.localeCompare(b.date) || a.session.localeCompare(b.session) || a.code.localeCompare(b.code));
    for (const paper of ordered) {
      let assigned = 0;
      for (let slot = 0; slot < paper.required[duty]; slot++) {
        const candidates = invigilators.filter((inv) => canAssign(inv, paper));
        if (candidates.length === 0) break;
        candidates.sort((a, b) => score(a, paper, duty, targets, targetMinutes).localeCompare(score(b, paper, duty, targets, targetMinutes)));
        addAssignment(candidates[0], paper, duty);
        assigned++;
      }
      if (assigned !== paper.required[duty]) shortages.push({ code: paper.code, duty, required: paper.required[duty], assigned });
    }
  }
  return shortages;
}

function countRange(values: number[]) {
  return values.length ? Math.max(...values) - Math.min(...values) : 0;
}

function scheduleObjective(invigilators: Invigilator[], targetMinutes: number) {
  const normalAa = invigilators.map((inv) => inv.counts.I + inv.counts.AA);
  const aa = invigilators.map((inv) => inv.counts.AA);
  const standby = invigilators.map((inv) => inv.counts.S);
  const minutes = invigilators.map((inv) => inv.minutes);
  const doubleTotal = invigilators.reduce((sum, inv) => sum + getDoubleDays(inv), 0);
  const minuteSquares = invigilators.reduce((sum, inv) => sum + (inv.minutes - targetMinutes) * (inv.minutes - targetMinutes), 0);
  return [
    countRange(normalAa),
    countRange(aa),
    doubleTotal,
    countRange(standby),
    countRange(minutes),
    Math.round(minuteSquares)
  ].map((number, index) => String(number).padStart(index === 5 ? 16 : 8, "0")).join("|");
}

function improveByTransfer(invigilators: Invigilator[], paperByColumn: Map<number, Paper>, targetMinutes: number) {
  for (let iteration = 0; iteration < 1200; iteration++) {
    const current = scheduleObjective(invigilators, targetMinutes);
    let changed = false;
    const highList = [...invigilators].sort((a, b) => (b.counts.I + b.counts.AA) - (a.counts.I + a.counts.AA) || b.counts.AA - a.counts.AA || getDoubleDays(b) - getDoubleDays(a) || b.counts.S - a.counts.S || b.minutes - a.minutes || a.name.localeCompare(b.name));
    const lowList = [...invigilators].sort((a, b) => (a.counts.I + a.counts.AA) - (b.counts.I + b.counts.AA) || a.counts.AA - b.counts.AA || getDoubleDays(a) - getDoubleDays(b) || a.counts.S - b.counts.S || a.minutes - b.minutes || a.name.localeCompare(b.name));
    for (const high of highList) {
      for (const low of lowList) {
        if (high.name === low.name) continue;
        if (high.counts.I + high.counts.AA <= low.counts.I + low.counts.AA + 1) continue;
        for (const column of [...high.assignments.keys()].sort((a, b) => a - b)) {
          const duty = high.assignments.get(column)!;
          if (duty !== "I" && duty !== "AA") continue;
          const paper = paperByColumn.get(column)!;
          if (!canAssign(low, paper) || teachesSubject(low, paper.subject)) continue;
          removeAssignment(high, paper, duty, paperByColumn);
          addAssignment(low, paper, duty);
          if (scheduleObjective(invigilators, targetMinutes) < current) {
            changed = true;
            break;
          }
          removeAssignment(low, paper, duty, paperByColumn);
          addAssignment(high, paper, duty);
        }
        if (changed) break;
      }
      if (changed) break;
    }
    if (!changed) break;
  }
}

function formulaTerms(terms: string[]) {
  return terms.length ? terms.join("+") : "0";
}

function setSummaryFormulas(schedule: any, papers: Paper[], firstRow: number, lastRow: number) {
  const firstColumn = papers[0].column;
  const lastColumn = papers[papers.length - 1].column;
  const firstName = columnName(firstColumn);
  const lastName = columnName(lastColumn);
  const normalTerms: string[] = [];
  const aaTerms: string[] = [];
  const amTerms: string[] = [];
  const pmTerms: string[] = [];
  const dates = [...new Set(papers.map((paper) => paper.date))];
  for (const paper of [...papers].sort((a, b) => a.column - b.column)) {
    const col = columnName(paper.column);
    normalTerms.push(`IF(${col}{r}="I",${paper.normalMinutes},0)`);
    aaTerms.push(`IF(${col}{r}="AA",${paper.aaMinutes},0)`);
    const countTerm = `COUNTIF(${col}{r},"I")+COUNTIF(${col}{r},"AA")+COUNTIF(${col}{r},"S")`;
    if (paper.session === "AM") amTerms.push(countTerm);
    if (paper.session === "PM") pmTerms.push(countTerm);
  }
  for (let row = firstRow; row <= lastRow; row++) {
    const range = `${firstName}${row}:${lastName}${row}`;
    schedule.cell(row, 4).formula(`COUNTIF(${range},"I")`);
    schedule.cell(row, 5).formula(`COUNTIF(${range},"AA")`);
    schedule.cell(row, 6).formula(`D${row}+E${row}`);
    schedule.cell(row, 7).formula(`COUNTIF(${range},"S")`);
    schedule.cell(row, 8).formula(`${formulaTerms(normalTerms).replaceAll("{r}", String(row))}+${formulaTerms(aaTerms).replaceAll("{r}", String(row))}`);
    schedule.cell(row, 9).formula(formulaTerms(amTerms).replaceAll("{r}", String(row)));
    schedule.cell(row, 10).formula(formulaTerms(pmTerms).replaceAll("{r}", String(row)));
    const doubleTerms: string[] = [];
    for (const date of dates) {
      const amRefs: string[] = [];
      const pmRefs: string[] = [];
      for (const paper of papers.filter((item) => item.date === date)) {
        const col = columnName(paper.column);
        const term = `COUNTIF(${col}${row},"I")+COUNTIF(${col}${row},"AA")+COUNTIF(${col}${row},"S")`;
        if (paper.session === "AM") amRefs.push(term);
        if (paper.session === "PM") pmRefs.push(term);
      }
      if (amRefs.length && pmRefs.length) doubleTerms.push(`IF(AND((${amRefs.join("+")})>0,(${pmRefs.join("+")})>0),1,0)`);
    }
    schedule.cell(row, 11).formula(formulaTerms(doubleTerms));
  }
}

function writeSchedule(schedule: any, papers: Paper[], invigilators: Invigilator[], shortages: Array<{ code: string; duty: Duty; required: number; assigned: number }>) {
  const lastRow = lastNamedRow(schedule);
  const shortageCodes = new Set(shortages.filter((item) => item.assigned < item.required).map((item) => item.code));
  for (const paper of papers) {
    if (shortageCodes.has(paper.code)) setFill(schedule.cell(3, paper.column), ORANGE);
    for (let row = 6; row <= lastRow; row++) {
      schedule.cell(row, paper.column).clear();
      setFill(schedule.cell(row, paper.column));
    }
  }
  for (const inv of invigilators) {
    for (const paper of papers) {
      if (inv.unavailableColumns.has(paper.unavailabilityColumn)) setText(schedule, inv.row, paper.column, "X");
    }
    for (const [column, duty] of inv.assignments.entries()) {
      const paper = papers.find((item) => item.column === column)!;
      setText(schedule, inv.row, column, duty);
      setFill(schedule.cell(inv.row, column), teachesSubject(inv, paper.subject) ? RED : YELLOW);
      schedule.cell(inv.row, column).style("horizontalAlignment", CENTER).style("verticalAlignment", CENTER);
    }
  }
  setSummaryFormulas(schedule, papers, 6, lastRow);
  const noteRow = lastRow + 2;
  setText(schedule, noteRow, 1, "Scheduling notes");
  setText(schedule, noteRow, 2, shortages.length === 0 ? "All grouped duty requirements filled." : shortages.map((item) => `${item.code} ${item.duty}: required ${item.required}, assigned ${item.assigned}`).join("; "));
}

export async function createDutyScheduleSheet(unavailabilityBuffer: Buffer) {
  const workbook = await XlsxPopulate.fromDataAsync(unavailabilityBuffer);
  const examDates = workbook.sheet("Exam Dates");
  const unavailability = workbook.sheet("Unavailability");
  if (!examDates || !unavailability) throw new Error("Workbook must contain Exam Dates and Unavailability sheets.");
  const schedule = createScheduleSheet(workbook, unavailability);
  const papers = collectPapers(examDates, schedule, unavailability);
  const invigilators = collectInvigilators(schedule, unavailability);
  if (papers.length === 0) throw new Error("No schedule columns found.");
  if (invigilators.length === 0) throw new Error("No invigilators found.");
  const paperByColumn = new Map(papers.map((paper) => [paper.column, paper]));
  let targetMinutes = papers.reduce((sum, paper) => sum + paper.required.I * paper.normalMinutes + paper.required.AA * paper.aaMinutes, 0) / invigilators.length;
  const shortages = makeSchedule(papers, invigilators, targetMinutes);
  targetMinutes = invigilators.reduce((sum, inv) => sum + inv.minutes, 0) / invigilators.length;
  improveByTransfer(invigilators, paperByColumn, targetMinutes);
  writeSchedule(schedule, papers, invigilators, shortages);
  return outputWorkbook(workbook);
}
