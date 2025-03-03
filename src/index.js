const ExcelJS = require("exceljs");

// 是否在浏览器环境
const isBrowser = typeof window !== "undefined" && typeof document !== "undefined";

/**
 * 加载工作簿
 * @param {string|Buffer|ArrayBuffer|Blob|File} input
 * @returns {Promise<ExcelJS.Workbook>}
 */
async function loadWorkbook(input) {
  const workbook = new ExcelJS.Workbook();
  if (isBrowser) {
    if (input instanceof Blob || input instanceof File) {
      const arrayBuffer = await input.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);
    } else if (input instanceof ArrayBuffer) {
      await workbook.xlsx.load(input);
    } else {
      throw new Error("Unsupported input type in browser environment. Expected Blob, File, ArrayBuffer.");
    }
  } else {
    if (typeof input === "string") {
      await workbook.xlsx.readFile(input);
    } else if (input instanceof Buffer || input instanceof ArrayBuffer) {
      await workbook.xlsx.load(input);
    } else if (typeof input.pipe === "function") {
      await workbook.xlsx.read(input);
    } else {
      throw new Error(
        "Unsupported input type in Node.js environment. Expected file path, Buffer, ArrayBuffer, or Stream."
      );
    }
  }
  return workbook;
}

/**
 * 填充Excel模板
 * @param {ExcelJS.Workbook} workbook
 * @param {Array} workbookData
 * @returns {Promise<ExcelJS.Workbook>}
 */
async function fillTemplate(workbook, workbookData) {
  // 工作表待合并单元格信息
  const sheetDynamicMerges = {};

  // 第一步：复制行并替换占位符
  // NOTE 工作簿的sheetId是按工作表创建的顺序从1开始递增
  let sheetIndex = 0;
  workbook.eachSheet((worksheet, sheetId) => {
    const worksheetData = workbookData[sheetIndex];
    sheetIndex++;
    if (worksheetData && typeof worksheetData === "object") {
      // NOTE 合并信息是静态的，不会随着行增加而实时更新
      const originMerges = worksheet.model.merges.map((merge) => {
        // C30:D30
        const [startCell, endCell] = merge.split(":");
        return {
          start: { row: worksheet.getCell(startCell).row, col: worksheet.getCell(startCell).col },
          end: { row: worksheet.getCell(endCell).row, col: worksheet.getCell(endCell).col },
        };
      });

      // 替换单字段占位符
      worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          if (typeof cell.value === "string") {
            for (const key in worksheetData) {
              if (typeof worksheetData[key] !== "object" && cell.value.includes(`{{${key}}}`)) {
                cell.value = cell.value.replace(new RegExp(`{{${key}}}`, "g"), worksheetData[key]);
              }
            }
          }
        });
      });

      // 查找所有迭代标签
      const iterationTags = [];
      const iterationRowNumbers = [];
      worksheet.eachRow((row, rowNumber) => {
        // 标记当前行是否已经找到迭代标签
        let isIterationRow = false;
        const iterationRegex = /{{([^\.]+)\.[^}]+}}/;
        row.eachCell((cell, colNumber) => {
          if (!isIterationRow && typeof cell.value === "string" && cell.value.match(iterationRegex)) {
            isIterationRow = true;
            const iterFieldName = cell.value.match(iterationRegex)[1];
            iterationTags.push({ iterStartRow: rowNumber, iterFieldName });
            iterationRowNumbers.push(rowNumber);
          }
        });
      });
      // 迭代行并替换迭代字段占位符
      let iterOffset = 0;
      iterationTags.forEach(({ iterStartRow, iterFieldName }, iterationTagIndex) => {
        if (Array.isArray(worksheetData[iterFieldName])) {
          const iterData = worksheetData[iterFieldName];
          // 调整后的起始行
          const adjustedStartRow = iterStartRow + iterOffset;
          // 多行的情况下，需要复制多行
          if (iterData.length > 1) {
            // 一次性复制多行
            // NOTE 复制的行不会复制合并信息
            worksheet.duplicateRow(adjustedStartRow, iterData.length - 1, true);
            // 筛选出与当前模板行相关的合并单元格信息，并应用到其复制的行
            const merges = originMerges.filter((merge) => {
              return merge.start.row <= iterStartRow && merge.end.row >= iterStartRow;
            });
            if (merges.length > 0) {
              if (!sheetDynamicMerges[sheetId]) {
                sheetDynamicMerges[sheetId] = [];
              }
              // NOTE 在浏览器环境，动态增加的行会使其后面的行取消合并单元格
              const startFixIndex = isBrowser ? (iterationTagIndex === 0 ? 1 : 0) : 1;
              for (let i = startFixIndex; i < iterData.length; i++) {
                for (const merge of merges) {
                  sheetDynamicMerges[sheetId].push([
                    merge.start.row + i + iterOffset,
                    merge.start.col,
                    merge.end.row + i + iterOffset,
                    merge.end.col,
                  ]);
                }
              }
            }
          }
          // 替换迭代行中的占位符
          for (let i = 0; i < iterData.length; i++) {
            const currentRow = worksheet.getRow(adjustedStartRow + i);
            currentRow.eachCell((cell, colNumber) => {
              if (typeof cell.value === "string") {
                // 替换迭代字段占位符
                for (const key in iterData[i]) {
                  if (cell.value.includes(`{{${iterFieldName}.${key}}}`)) {
                    cell.value = cell.value.replace(new RegExp(`{{${iterFieldName}.${key}}}`, "g"), iterData[i][key]);
                  }
                }
              }
            });
          }
          // 更新行号偏移量
          iterOffset += iterData.length - 1;
        }
      });
      // 修正浏览器环境下，迭代行之后的合并单元格信息
      if (isBrowser) {
        originMerges.forEach((merge) => {
          // 迭代后的偏移行
          let mergeOffset = 0;
          iterationTags.forEach(({ iterStartRow, iterFieldName }) => {
            if (Array.isArray(worksheetData[iterFieldName])) {
              if (!iterationRowNumbers.includes(merge.start.row) && merge.start.row > iterStartRow) {
                mergeOffset += worksheetData[iterFieldName].length - 1;
              }
            }
          });
          if (mergeOffset) {
            if (!sheetDynamicMerges[sheetId]) {
              sheetDynamicMerges[sheetId] = [];
            }
            sheetDynamicMerges[sheetId].push([
              merge.start.row + mergeOffset,
              merge.start.col,
              merge.end.row + mergeOffset,
              merge.end.col,
            ]);
          }
        });
      }
    }
  });

  // 第二步：动态行单元格合并处理
  if (Object.keys(sheetDynamicMerges).length > 0) {
    // 将工作簿保存到内存中的缓冲区
    const buffer = await workbook.xlsx.writeBuffer();
    // 从缓冲区重新加载工作簿
    await workbook.xlsx.load(buffer);
    // 处理合并单元格
    workbook.eachSheet((worksheet, sheetId) => {
      if (sheetDynamicMerges[sheetId]) {
        sheetDynamicMerges[sheetId].forEach((merge) => {
          try {
            worksheet.mergeCells(merge);
          } catch (error) {
            console.warn(`Fail to merge cells ${merge}: ${error.message}`);
          }
        });
      }
    });
  }

  return workbook;
}

/**
 * 保存工作簿到文件
 * @param {ExcelJS.Workbook} workbook
 * @param {string} output
 * @returns {Promise<void>}
 */
async function saveWorkbook(workbook, output) {
  if (isBrowser) {
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = output;
    link.click();
    URL.revokeObjectURL(link.href);
  } else {
    await workbook.xlsx.writeFile(output);
  }
}

module.exports = { loadWorkbook, fillTemplate, saveWorkbook };
