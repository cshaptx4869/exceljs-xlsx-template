import ExcelJS from 'exceljs';

/**
 * 加载工作簿
 * @param {string | Buffer | ArrayBuffer | Blob | File} input
 * @returns {Promise<ExcelJS.Workbook>}
 */
declare function loadWorkbook(input: string | Buffer | ArrayBuffer | Blob | File): Promise<ExcelJS.Workbook>;

/**
 * 填充Excel模板
 * @param {ExcelJS.Workbook} workbook
 * @param {Array<Record<string, any>>} workbookData
 * @returns {Promise<ExcelJS.Workbook>}
 */
declare function fillTemplate(workbook: ExcelJS.Workbook, workbookData: Array<Record<string, any>>): Promise<ExcelJS.Workbook>;

/**
 * 保存工作簿到文件
 * @param {ExcelJS.Workbook} workbook
 * @param {string} output
 * @returns {Promise<void>}
 */
declare function saveWorkbook(workbook: ExcelJS.Workbook, output: string): Promise<void>;

export { loadWorkbook, fillTemplate, saveWorkbook };