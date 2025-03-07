import ExcelJS from "exceljs";

/**
 * 加载工作簿
 * @param {string | ArrayBuffer | Blob | Buffer} input - 输入数据，可以是本地路径、URL地址、ArrayBuffer、Blob、Buffer
 * @returns {Promise<ExcelJS.Workbook>}
 */
declare function loadWorkbook(input: string | ArrayBuffer | Blob | Buffer): Promise<ExcelJS.Workbook>;

/**
 * 填充Excel模板
 * @param {ExcelJS.Workbook} workbook
 * @param {Array<Record<string, any>>} workbookData - 包含模板数据的数组对象
 * @param {boolean} parseImage - 是否解析图片，默认为 false
 * @returns {Promise<ExcelJS.Workbook>}
 */
declare function fillTemplate(
  workbook: ExcelJS.Workbook,
  workbookData: Array<Record<string, any>>,
  parseImage?: boolean
): Promise<ExcelJS.Workbook>;

/**
 * 保存工作簿到文件
 * @param {ExcelJS.Workbook} workbook
 * @param {string} output - 输出文件路径或文件名
 * @returns {Promise<void>}
 */
declare function saveWorkbook(workbook: ExcelJS.Workbook, output: string): Promise<void>;

/**
 * 获取自定义占位符单元格范围
 * @param {ExcelJS.Worksheet} worksheet
 * @param {string} placeholder - 占位符字符串，默认为 "{{#placeholder}}"
 * @param {boolean} clearMatch - 是否清除占位符，默认为 true
 * @returns {{start: {row: number, col: number}, end: {row: number, col: number}}|null}
 */
declare function placeholderRange(
  worksheet: ExcelJS.Worksheet,
  placeholder?: string,
  clearMatch?: boolean
): { start: { row: number; col: number }; end: { row: number; col: number } } | null;

export { loadWorkbook, fillTemplate, saveWorkbook, placeholderRange };
