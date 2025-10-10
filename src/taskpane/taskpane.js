/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("runA3").onclick = runA3;
    document.getElementById("runChiPhi").onclick = runChiPhi;
  }
});

export async function runChiPhi() {
  try {
    await Excel.run(async (context) => {
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

export async function runA3() {
  try {
    await Excel.run(async (context) => {
      //Get data from current sheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load("values");
      await context.sync();
      const values = usedRange.values;
      let headerRowIndex = -1;
      // Find the header row containing "Tên KH"
      for (let i = 0; i < values.length; i++) {
          if (values[i].some(cell => typeof cell === 'string' && cell.includes("Tên KH"))) {
              headerRowIndex = i;
              break;
          }
      }
      if (headerRowIndex === -1) throw new Error("Header row not found");
      const headers = values[headerRowIndex];
      const dataRows = usedRange.values.slice(1);
      // Find column indices using partial match
      const getColIndex = (headers, partialName) => 
          headers.findIndex(h => h.includes(partialName));
      const cols = {
          tenKH: getColIndex(headers, "Tên KH"),
          maChuyen: getColIndex(headers, "Mãchuyến"),
          ngayGoiXe: getColIndex(headers, "Ngày gọi xe"),
          xeTQ: getColIndex(headers, "Xe TQ"),
          xeVN: getColIndex(headers, "Xe VN"),
          loaiXe: getColIndex(headers, "Loại xe"),
          tenKHDT: getColIndex(headers, "Tên KH Đ. Tác"),
          noiDung: getColIndex(headers, "Nội dung"),
          cungDuong: getColIndex(headers, "Cung Đường")
      };
      // Filter rows where "Tên KH" is "A3"
      const filteredData = dataRows
          .filter(row => row[cols.tenKH] === "A3")
          .map(row => ({
              "Mãchuyến": row[cols.maChuyen],
              "Ngày gọi xe": row[cols.ngayGoiXe],
              "Xe TQ": row[cols.xeTQ],
              "Xe VN": row[cols.xeVN],
              "Loại xe": row[cols.loaiXe],
              "Tên KH Đ. Tác": row[cols.tenKHDT],
              "Nội dung": row[cols.noiDung],
              "Cung Đường": row[cols.cungDuong]
          }));
      console.log(filteredData);
      // Add new A3 sheet
      const fileUrl = "https://raw.githubusercontent.com/duyanhbk93tn/DaiLamMoc/main/tmpA3.xlsx";
      const response = await fetch(fileUrl);
      const arrayBuffer = await response.arrayBuffer();
      const base64Data = window.btoa(
          new Uint8Array(arrayBuffer).reduce((data, byte) => data + String.fromCharCode(byte), '')
      );
      const currentSheetName = context.workbook.worksheets.getActiveWorksheet().load("name");
      await context.sync();
      const options = {
        positionType: Excel.WorksheetPositionType.after,
        relativeTo: currentSheetName
      };
      context.workbook.insertWorksheetsFromBase64(base64Data, options);
      await context.sync();
      const A3sheet = context.workbook.worksheets.getActiveWorksheet();
      const now = new Date();
      const month = String(now.getMonth() + 1).padStart(2, '0'); // mm
      const day = String(now.getDate()).padStart(2, '0'); // dd
      A3sheet.name = "A3-" + `${month}.${day}`;
      const startRow = 14; // Update from row 14
      const A3dataRange = A3sheet.getRangeByIndexes(startRow - 1, 0, filteredData.length, 17); // A to Q
      A3dataRange.load("values");
      await context.sync();
      const updatedValues = A3dataRange.values;
      filteredData.forEach((item, index) => {
          updatedValues[index][0] = "'" + item["Mãchuyến"];        // A
          updatedValues[index][1] = item["Ngày gọi xe"];     // B
          updatedValues[index][4] = item["Xe VN"];           // E
          updatedValues[index][5] = item["Loại xe"];         // F
          updatedValues[index][9] = item["Tên KH Đ. Tác"];   // J
          updatedValues[index][15] = item["Xe TQ"];          // P
          updatedValues[index][16] = item["Cung Đường"];     // Q
      });
      A3dataRange.values = updatedValues;
      await context.sync();
      const endRow = 14 + filteredData.length;
      if (endRow <= 149) {
          const rowsToDelete = A3sheet.getRange(`A${endRow}:Z149`);
          rowsToDelete.delete(Excel.DeleteShiftDirection.up);
      }
      await context.sync();
      const rangeA3 = A3sheet.getRange(`A${startRow}:P${endRow}`);
      rangeA3.load("values");
      await context.sync();0
      const valuesA3 = rangeA3.values;
      let i = 0;
      while (i < valuesA3.length) {
          const currentValue = valuesA3[i][0];
          let mergeCount = 1;
          // Count consecutive same values in column A
          while (i + mergeCount < valuesA3.length && valuesA3[i + mergeCount][0] === currentValue) {
              mergeCount++;
          }
          if (mergeCount > 1) {
              const startRowIndex = startRow + i;
              const endRowIndex = startRow + i + mergeCount - 1;
              // Merge column A
              A3sheet.getRange(`A${startRowIndex}:A${endRowIndex}`).merge();
              await context.sync();
              // For columns B, E, F: merge consecutive same values within [startRowIndex, endRowIndex]
              ["B", "E", "F"].forEach(col => {
                  let i = 0;
                  while (i < (endRowIndex - startRowIndex + 1)) {
                      let count = 1;
                      while (i + count < (endRowIndex - startRowIndex + 1) && valuesA3[i + count][col - "A"] === valuesA3[i][col - "A"]) {
                          count++;
                      }
                      if (count > 1) {
                          A3sheet.getRange(`${col}${startRowIndex + i}:${col}${startRowIndex + i + count - 1}`).merge();
                      }
                      i += count;
                  }
              });
          }
          i += mergeCount;
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
