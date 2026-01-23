export interface SelectionInfo {
  range: string;
  rowCount: number;
  headers: string[];
}

export interface Record {
  [key: string]: unknown;
}

/**
 * Get information about the current selection
 */
export async function getSelectionInfo(): Promise<SelectionInfo> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address", "rowCount", "values"]);
    await context.sync();

    const values = range.values;
    if (!values || values.length < 2) {
      return {
        range: range.address,
        rowCount: 0,
        headers: [],
      };
    }

    // First row is headers
    const headers = (values[0] as unknown[])
      .map((h, idx) => {
        const val = String(h ?? "").trim();
        return val || `Column ${String.fromCharCode(65 + idx)}`;
      })
      .filter((_, idx) => {
        // Only include columns that have at least some data
        return values.some(
          (row, rowIdx) => rowIdx > 0 && row[idx] !== null && row[idx] !== ""
        );
      });

    return {
      range: range.address,
      rowCount: values.length - 1, // Exclude header row
      headers,
    };
  });
}

/**
 * Convert current selection to array of records
 */
export async function selectionToRecords(): Promise<Record[]> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("values");
    await context.sync();

    const values = range.values;
    if (!values || values.length < 2) {
      return [];
    }

    // First row is headers
    const rawHeaders = values[0] as unknown[];
    const headers = rawHeaders.map((h, idx) => {
      const val = String(h ?? "").trim();
      return val || `Column ${String.fromCharCode(65 + idx)}`;
    });

    // Find columns with data
    const validColumns = headers
      .map((_, idx) => idx)
      .filter((idx) =>
        values.some(
          (row, rowIdx) => rowIdx > 0 && row[idx] !== null && row[idx] !== ""
        )
      );

    // Convert to records
    const records: Record[] = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];

      // Skip empty rows
      const hasData = validColumns.some(
        (idx) => row[idx] !== null && row[idx] !== ""
      );
      if (!hasData) continue;

      const record: Record = {};
      for (const idx of validColumns) {
        record[headers[idx]] = row[idx];
      }
      records.push(record);
    }

    return records;
  });
}

/**
 * Read records from a specific range
 */
export async function rangeToRecords(rangeAddress: string): Promise<Record[]> {
  return Excel.run(async (context) => {
    const range = context.workbook.worksheets
      .getActiveWorksheet()
      .getRange(rangeAddress);
    range.load("values");
    await context.sync();

    const values = range.values;
    if (!values || values.length < 2) {
      return [];
    }

    const rawHeaders = values[0] as unknown[];
    const headers = rawHeaders.map((h, idx) => {
      const val = String(h ?? "").trim();
      return val || `Column ${String.fromCharCode(65 + idx)}`;
    });

    const records: Record[] = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const hasData = row.some((cell) => cell !== null && cell !== "");
      if (!hasData) continue;

      const record: Record = {};
      for (let j = 0; j < headers.length; j++) {
        record[headers[j]] = row[j];
      }
      records.push(record);
    }

    return records;
  });
}

/**
 * Write records to a new worksheet
 */
export async function writeResultsToSheet(
  records: Record[],
  sheetNameBase: string = "Results"
): Promise<string> {
  if (records.length === 0) {
    throw new Error("No results to write");
  }

  return Excel.run(async (context) => {
    // Get all unique keys from records
    const allKeys = new Set<string>();
    for (const record of records) {
      for (const key of Object.keys(record)) {
        allKeys.add(key);
      }
    }
    const headers = Array.from(allKeys);

    // Create new sheet with unique name
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    let sheetName = sheetNameBase;
    let counter = 1;
    const existingNames = new Set(sheets.items.map((s) => s.name));
    while (existingNames.has(sheetName)) {
      sheetName = `${sheetNameBase} (${counter})`;
      counter++;
    }

    const newSheet = sheets.add(sheetName);

    // Write headers
    const headerRange = newSheet.getRange(`A1:${getColumnLetter(headers.length)}1`);
    headerRange.values = [headers];
    headerRange.format.font.bold = true;
    headerRange.format.fill.color = "#f0f0f0";

    // Write data
    const dataValues = records.map((record) =>
      headers.map((header) => {
        const value = record[header];
        // Convert objects to JSON strings
        if (typeof value === "object" && value !== null) {
          return JSON.stringify(value);
        }
        return value ?? "";
      })
    );

    const dataRange = newSheet.getRange(
      `A2:${getColumnLetter(headers.length)}${records.length + 1}`
    );
    dataRange.values = dataValues;

    // Format
    newSheet.getRange("1:1").format.rowHeight = 25;
    newSheet.freezePanes.freezeRows(1);

    // Auto-fit columns
    const usedRange = newSheet.getUsedRange();
    usedRange.format.autofitColumns();

    // Activate the new sheet
    newSheet.activate();

    await context.sync();

    return sheetName;
  });
}

/**
 * Convert column index to Excel column letter (0 = A, 25 = Z, 26 = AA, etc.)
 */
function getColumnLetter(colIndex: number): string {
  let letter = "";
  let index = colIndex;
  while (index > 0) {
    const remainder = (index - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    index = Math.floor((index - 1) / 26);
  }
  return letter || "A";
}
