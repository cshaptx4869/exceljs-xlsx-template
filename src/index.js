"use strict";

const ExcelJS = require("exceljs");
const { isBrowser, fetchUrlFile } = require("./helpers.js");

/**
 * 加载工作簿
 * @param {string|ArrayBuffer|Blob|Buffer} input - 输入数据，可以是本地路径、URL地址、ArrayBuffer、Blob、Buffer
 * @returns {Promise<ExcelJS.Workbook>}
 */
async function loadWorkbook(input) {
  const workbook = new ExcelJS.Workbook();
  const httpRegex = /^https?:\/\//;
  if (isBrowser) {
    if (typeof input === "string" && httpRegex.test(input)) {
      await workbook.xlsx.load(await fetchUrlFile(input));
    } else if (input instanceof Blob || input instanceof ArrayBuffer) {
      await workbook.xlsx.load(input);
    } else {
      throw new Error("Unsupported input type in browser environment. Expected Blob, File, ArrayBuffer.");
    }
  } else {
    if (typeof input === "string") {
      if (httpRegex.test(input)) {
        await workbook.xlsx.load(await fetchUrlFile(input));
      } else {
        await workbook.xlsx.readFile(input);
      }
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
 * @param {Array<Record<string, any>>} workbookData - 包含模板数据的数组对象
 * @param {boolean} parseImage - 是否解析图片，默认为 false
 * @returns {Promise<ExcelJS.Workbook>}
 */
async function fillTemplate(workbook, workbookData, parseImage = false) {
  // 第一步：复制行并替换占位符
  // 工作表待合并单元格信息
  const sheetDynamicMerges = {};
  // NOTE 工作簿的sheetId是按工作表创建的顺序从1开始递增
  let sheetIndex = 0;
  workbook.eachSheet((worksheet, sheetId) => {
    const sheetData = workbookData[sheetIndex++];
    if (!(sheetData && typeof sheetData === "object" && !Array.isArray(sheetData))) {
      return;
    }
    // 单标签替换和迭代标签信息收集
    const sheetIterTags = [];
    worksheet.eachRow((row, rowNumber) => {
      // 行迭代字段
      const rowIterFields = [];
      row.eachCell((cell, colNumber) => {
        const cellType = cell.type;
        // 字符串值
        if (cellType === ExcelJS.ValueType.String) {
          cell.value = processCellTags(cell.value, sheetData, sheetIterTags, rowIterFields, rowNumber);
        }
        // 富文本值
        else if (cellType === ExcelJS.ValueType.RichText) {
          cell.value.richText.forEach((item) => {
            item.text = processCellTags(item.text, sheetData, sheetIterTags, rowIterFields, rowNumber);
          });
        }
      });
    });
    // 迭代标签处理
    if (sheetIterTags.length === 0) {
      return;
    }
    // 合并单元格信息
    // NOTE 合并信息是静态的，不会随着行增加而实时更新
    const sheetMerges = sheetMergeInfo(worksheet);
    // 迭代行并替换迭代字段占位符
    let iterOffset = 0;
    sheetIterTags.forEach(({ iterStartRow, iterFieldNames, iterFieldName }, iterTagIndex) => {
      // 调整后的起始行
      const adjustedStartRow = iterStartRow + iterOffset;
      // 多行的情况下，需要复制多行
      if (sheetData[iterFieldName].length > 1) {
        // 一次性复制多行
        // NOTE 复制的行不会复制合并信息
        worksheet.duplicateRow(adjustedStartRow, sheetData[iterFieldName].length - 1, true);
        // 筛选出与当前模板行相关的合并单元格信息，并应用到其复制的行
        const merges = sheetMerges.filter((merge) => {
          return merge.start.row <= iterStartRow && merge.end.row >= iterStartRow;
        });
        if (merges.length > 0) {
          if (!sheetDynamicMerges[sheetId]) {
            sheetDynamicMerges[sheetId] = [];
          }
          // NOTE 在浏览器环境，动态增加的行会使其后面的行取消合并单元格
          const startFixIndex = isBrowser ? (iterTagIndex === 0 ? 1 : 0) : 1;
          for (let i = startFixIndex; i < sheetData[iterFieldName].length; i++) {
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
      for (let i = 0; i < sheetData[iterFieldName].length; i++) {
        const currentRow = worksheet.getRow(adjustedStartRow + i);
        currentRow.eachCell((cell, colNumber) => {
          // 字符串值
          if (cell.type === ExcelJS.ValueType.String) {
            for (const iterField of iterFieldNames) {
              const iterFieldData = sheetData[iterField];
              if (cell.value.includes(`{{${iterField}\.`)) {
                if (iterFieldData[i] !== undefined) {
                  for (const key in iterFieldData[i]) {
                    const placeholder = `{{${iterField}.${key}}}`;
                    if (cell.value.includes(placeholder)) {
                      if (cell.value.length === placeholder.length) {
                        cell.value = iterFieldData[i][key];
                      } else {
                        cell.value = cell.value.replace(new RegExp(placeholder, "g"), iterFieldData[i][key]);
                      }
                    }
                  }
                } else {
                  cell.value = null;
                }
              }
            }
          }
          // TODO 迭代标签单元格为富文本值
        });
      }
      // 更新行号偏移量
      iterOffset += sheetData[iterFieldName].length - 1;
    });
    // 修正在浏览器环境，动态增加的行会使其后面的行取消合并单元格
    if (isBrowser) {
      const iterRows = sheetIterTags.map(({ iterStartRow }) => iterStartRow);
      sheetMerges.forEach((merge) => {
        // 迭代后的偏移行
        let mergeOffset = 0;
        sheetIterTags.forEach(({ iterStartRow, iterFieldName }) => {
          if (Array.isArray(sheetData[iterFieldName])) {
            if (!iterRows.includes(merge.start.row) && merge.start.row > iterStartRow) {
              mergeOffset += sheetData[iterFieldName].length - 1;
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
            console.warn(`Fail to merge cells ${merge}`);
          }
        });
      }
    });
  }

  // 第三步：填充图片
  parseImage && (await fillImage(workbook));

  return workbook;
}

/**
 * 保存工作簿到文件
 * @param {ExcelJS.Workbook} workbook
 * @param {string} output - 输出文件路径或文件名
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

/**
 * 获取自定义占位符单元格范围
 * @param {ExcelJS.Worksheet} worksheet
 * @param {string} placeholder - 占位符字符串，默认为 "{{#placeholder}}"
 * @param {boolean} clearMatch - 是否清除占位符，默认为 true
 * @returns {{start: {row: number, col: number}, end: {row: number, col: number}}|null}
 */
function placeholderRange(worksheet, placeholder = "{{#placeholder}}", clearMatch = true) {
  let result = null;
  const sheetMerges = sheetMergeInfo(worksheet);
  // 遍历每一行
  outer: for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
    const row = worksheet.getRow(rowNumber);
    // 遍历每个单元格
    for (let colNumber = 1; colNumber <= row.cellCount; colNumber++) {
      const cell = row.getCell(colNumber);
      // 单元格值中是否包含占位符
      if (typeof cell.value === "string" && cell.value.includes(`${placeholder}`)) {
        const info = sheetMerges.find((merge) => {
          return merge.start.row === rowNumber && merge.start.col === colNumber;
        });
        result = info ?? {
          start: { row: rowNumber, col: colNumber },
          end: { row: rowNumber, col: colNumber },
        };
        // 去除占位符
        if (clearMatch) {
          cell.value = cell.value.replace(new RegExp(`${placeholder}`, "g"), "");
        }
        // 跳出循环
        break outer;
      }
    }
  }
  return result;
}

/**
 * 处理单元格标签(单标签替换和迭代标签信息收集)
 * @param {string} target
 * @param {Record<string, any>} worksheetData
 * @param {Array<{iterStartRow: number, iterFieldNames: string[], iterFieldName: string}>} iterationTags
 * @param {string[]} iterFieldNames
 * @param {number} rowNumber
 * @returns {string}
 */
function processCellTags(target, worksheetData, iterationTags, iterFieldNames, rowNumber) {
  // 单标签占位符替换
  if (/{{\w+}}/.test(target)) {
    // 允许单元格中有多个单标签占位符
    const placeholders = target.match(/{{\w+}}/g);
    placeholders.forEach((placeholder) => {
      const fieldName = placeholder.slice(2, -2);
      if (fieldName in worksheetData) {
        if (target.length === placeholder.length && typeof worksheetData[fieldName] !== "object") {
          target = worksheetData[fieldName];
        } else {
          target = target.replace(placeholder, worksheetData[fieldName]);
        }
      }
    });
  }
  // 迭代标签信息搜集
  else if (/{{\w+\.\w+}}/.test(target)) {
    // TODO 单元格含多个迭代标签
    // 单元格中仅匹配一个迭代标签占位符
    const iterFieldName = target.match(/{{(\w+)\.\w+}}/)[1];
    if (
      iterFieldName in worksheetData &&
      Array.isArray(worksheetData[iterFieldName]) &&
      worksheetData[iterFieldName].length > 0
    ) {
      if (iterFieldNames.length === 0) {
        iterFieldNames.push(iterFieldName);
        iterationTags.push({ iterStartRow: rowNumber, iterFieldNames: iterFieldNames, iterFieldName: iterFieldName });
      } else {
        if (!iterFieldNames.includes(iterFieldName)) {
          iterFieldNames.push(iterFieldName);
          const lastIterationTag = iterationTags[iterationTags.length - 1];
          if (worksheetData[iterFieldName].length > worksheetData[lastIterationTag.iterFieldName].length) {
            lastIterationTag.iterFieldName = iterFieldName;
          }
        }
      }
    }
  }
  return target;
}

/**
 * 获取工作表合并信息
 * @param {ExcelJS.Worksheet} worksheet
 * @returns {Array<{start: {row: number, col: number}, end: {row: number, col: number}}>}
 */
function sheetMergeInfo(worksheet) {
  return worksheet.model.merges.map((merge) => {
    // C30:D30
    const [startCell, endCell] = merge.split(":");
    return {
      start: { row: worksheet.getCell(startCell).row, col: worksheet.getCell(startCell).col },
      end: { row: worksheet.getCell(endCell).row, col: worksheet.getCell(endCell).col },
    };
  });
}

/**
 * 填充图片
 * @param {ExcelJS.Workbook} workbook
 */
async function fillImage(workbook) {
  const workbookImage = {};
  const invalidImages = [];
  const imageRegex = /https?:\/\/[^\s]+?\.(jpe?g|gif|png)/i;
  // NOTE eachSheet、eachRow、eachCell都是同步方法，不会等待异步操作完成
  // 遍历每个工作表
  for (let i = 0; i < workbook.worksheets.length; i++) {
    const worksheet = workbook.worksheets[i];
    const sheetMerges = sheetMergeInfo(worksheet);
    // 遍历每一行
    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      // 遍历每个单元格
      for (let colNumber = 1; colNumber <= row.cellCount; colNumber++) {
        const cell = row.getCell(colNumber);
        // 检查单元格的值是否是图片地址
        if (typeof cell.value === "string" && cell.value.match(imageRegex)) {
          const matches = cell.value.match(imageRegex);
          const imageUrl = matches[0];
          const imageExt = matches[1];
          if (invalidImages.includes(imageUrl)) {
            continue;
          }
          // 如果图片未缓存，则加载图片
          if (workbookImage[imageUrl] === undefined) {
            let fileContent = null;
            try {
              fileContent = await fetchUrlFile(imageUrl);
            } catch {
              invalidImages.push(imageUrl);
              console.warn(`Fail to load image ${imageUrl}`);
              continue;
            }
            // 将图片添加到工作簿中
            workbookImage[imageUrl] = workbook.addImage({
              buffer: fileContent,
              extension: imageExt === "jpg" ? "jpeg" : imageExt,
            });
          }
          // 将图片添加到工作表中
          const merge = sheetMerges.find((merge) => {
            return merge.start.row === rowNumber && merge.start.col === colNumber;
          });
          // 坐标系基于零，A1 的左上角将为 {col：0，row：0}，右下角为 {col：1，row：1}
          worksheet.addImage(workbookImage[imageUrl], {
            // 左上角
            tl: {
              col: merge ? merge.start.col - 1 : colNumber - 1,
              row: merge ? merge.start.row - 1 : rowNumber - 1,
            },
            // 右下角
            br: {
              col: merge ? merge.end.col : colNumber,
              row: merge ? merge.end.row : rowNumber,
            },
          });
          // 去除图片地址
          cell.value = cell.value.replace(imageRegex, "");
        }
      }
    }
  }

  return workbook;
}

module.exports = {
  loadWorkbook,
  fillTemplate,
  saveWorkbook,
  placeholderRange,
};
