import "@testing-library/jest-dom";

// Mock Office.js
const mockOffice = {
  onReady: (callback: () => void) => callback(),
  context: {
    roamingSettings: {
      get: () => null,
      set: () => {},
      saveAsync: () => {},
    },
  },
};

// @ts-expect-error - Mocking Office global
globalThis.Office = mockOffice;

// Mock Excel.run
const mockExcelRun = async <T>(callback: (context: unknown) => Promise<T>): Promise<T> => {
  const mockContext = {
    workbook: {
      worksheets: {
        items: [{ name: "Sheet1" }],
        load: () => {},
        getItem: () => ({
          getUsedRangeOrNullObject: () => ({
            load: () => {},
            isNullObject: false,
            values: [
              ["Name", "email", "id"],
              ["Peter", "peter@yahoo", 1],
            ],
            address: "A1:C2",
          }),
          activate: () => {},
          getRange: () => ({
            values: null,
            format: { font: {}, fill: {}, rowHeight: 0, autofitColumns: () => {} },
          }),
          freezePanes: { freezeRows: () => {} },
        }),
        getActiveWorksheet: () => ({
          name: "Sheet1",
          load: () => {},
        }),
        add: () => ({
          getRange: () => ({
            values: null,
            format: { font: {}, fill: {}, rowHeight: 0, autofitColumns: () => {} },
          }),
          getUsedRange: () => ({
            format: { autofitColumns: () => {} },
          }),
          freezePanes: { freezeRows: () => {} },
          activate: () => {},
        }),
      },
    },
    sync: async () => {},
  };
  return callback(mockContext);
};

// @ts-expect-error - Mocking Excel global
globalThis.Excel = { run: mockExcelRun };
