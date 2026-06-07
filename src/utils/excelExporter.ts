import * as XLSX from "xlsx";

type JsonPrimitive = null | boolean | number | string;
type JsonValue = JsonPrimitive | JsonValue[] | { [key: string]: JsonValue };

export type ExcelRow = Record<string, JsonPrimitive>;
export type ExcelSheet = { name: string; rows: ExcelRow[]; headers: string[] };
export type ExcelData = { sheets: ExcelSheet[] };

function flattenObject(obj: Record<string, JsonValue>, prefix = ""): ExcelRow {
  const result: ExcelRow = {};
  for (const [key, value] of Object.entries(obj)) {
    const fullKey = prefix ? `${prefix}.${key}` : key;
    if (value !== null && typeof value === "object" && !Array.isArray(value)) {
      Object.assign(result, flattenObject(value as Record<string, JsonValue>, fullKey));
    } else if (Array.isArray(value)) {
      result[fullKey] = JSON.stringify(value);
    } else {
      result[fullKey] = value as JsonPrimitive;
    }
  }
  return result;
}

function toRows(value: JsonValue): ExcelRow[] {
  if (Array.isArray(value)) {
    return value.map((item): ExcelRow => {
      if (item !== null && typeof item === "object" && !Array.isArray(item)) {
        return flattenObject(item);
      }
      // Arrays-within-arrays get stringified; primitives/null pass through
      return { value: Array.isArray(item) ? JSON.stringify(item) : item };
    });
  }
  if (value !== null && typeof value === "object") {
    return [flattenObject(value as Record<string, JsonValue>)];
  }
  return [{ value }];
}

function collectHeaders(rows: ExcelRow[]): string[] {
  const set = new Set<string>();
  rows.forEach((row) => Object.keys(row).forEach((k) => set.add(k)));
  return Array.from(set);
}

function sanitizeSheetName(name: string): string {
  return name.replace(/[\\/?*[\]:]/g, "_").slice(0, 31) || "Sheet";
}

function makeSheet(name: string, rows: ExcelRow[]): ExcelSheet {
  return { name: sanitizeSheetName(name), rows, headers: collectHeaders(rows) };
}

export const prepareExcelData = (json: string): ExcelData => {
  const parsed = JSON.parse(json) as JsonValue;

  // Array at root → single sheet
  if (Array.isArray(parsed)) {
    return { sheets: [makeSheet("Sheet1", toRows(parsed))] };
  }

  // Single object → split array fields into their own sheets
  if (parsed !== null && typeof parsed === "object") {
    const obj = parsed as Record<string, JsonValue>;
    const arrayFields: [string, JsonValue[]][] = [];
    const scalarObj: Record<string, JsonValue> = {};

    for (const [key, value] of Object.entries(obj)) {
      if (Array.isArray(value)) {
        arrayFields.push([key, value]);
      } else {
        scalarObj[key] = value;
      }
    }

    const sheets: ExcelSheet[] = [];

    if (Object.keys(scalarObj).length > 0) {
      sheets.push(makeSheet("Main", [flattenObject(scalarObj)]));
    }

    for (const [fieldName, arr] of arrayFields) {
      sheets.push(makeSheet(fieldName, toRows(arr)));
    }

    return { sheets: sheets.length > 0 ? sheets : [makeSheet("Sheet1", [])] };
  }

  // Primitive
  return { sheets: [makeSheet("Sheet1", [{ value: parsed as JsonPrimitive }])] };
};

export const downloadExcel = (data: ExcelData): void => {
  const workbook = XLSX.utils.book_new();
  for (const sheet of data.sheets) {
    const ws = XLSX.utils.json_to_sheet(sheet.rows);
    XLSX.utils.book_append_sheet(workbook, ws, sheet.name);
  }
  XLSX.writeFile(workbook, "output.xlsx");
};
