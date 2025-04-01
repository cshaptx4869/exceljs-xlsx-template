"use strict";

const ExcelJS = require("exceljs");
const { loadWorkbook, fillTemplate, saveWorkbook, placeholderRange } = require("./core.js");

/**
 * 渲染Xlsx模板
 * @param {string|ArrayBuffer|Blob|Buffer} input - 输入数据，可以是本地路径、URL地址、ArrayBuffer、Blob、Buffer
 * @param {Array<Record<string, any>>} data - 包含模板数据的数组对象
 * @param {string} output - 输出文件路径或文件名
 * @param {{parseImage?: boolean; beforeSave?: (workbook: ExcelJS.Workbook) => void|Promise<void>}} options 配置项
 * @returns {Promise<void>}
 */
async function renderXlsxTemplate(input, data, output, options = { parseImage: false, beforeSave: undefined }) {
  const workbook = await loadWorkbook(input);
  await fillTemplate(workbook, data, options.parseImage === true);
  if (typeof options.beforeSave === "function") {
    await options.beforeSave(workbook);
  }
  await saveWorkbook(workbook, output);
}

module.exports = {
  loadWorkbook,
  fillTemplate,
  saveWorkbook,
  placeholderRange,
  renderXlsxTemplate,
};
