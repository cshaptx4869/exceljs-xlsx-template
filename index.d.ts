import ExcelJS from "exceljs";

/**
 * 输入数据类型
 * 支持本地路径、URL地址、ArrayBuffer、Blob 或 Buffer。
 */
type Input = string | ArrayBuffer | Blob | Buffer;

/**
 * 模板数据类型
 * 数组对象形式，键值对表示占位符和实际值。
 */
type Data = Array<Record<string, any>>;

/**
 * 渲染Xlsx模板并保存到指定路径。
 *
 * @param input - 输入数据，支持本地路径、URL地址、ArrayBuffer、Blob 或 Buffer。
 * @param data - 模板数据，数组对象形式，键值对表示占位符和实际值。
 * @param output - 输出文件路径或文件名。
 * @param options - 配置项：
 *   - parseImage?: 是否解析图片，默认为 false。
 *   - beforeSave?: 保存前的回调函数，接收工作簿对象，可返回 Promise。
 * @returns - 返回一个 Promise，完成时无返回值。
 */
declare function renderXlsxTemplate(
  input: Input,
  data: Data,
  output: string,
  options?: {
    parseImage?: boolean;
    beforeSave?: (workbook: ExcelJS.Workbook) => void | Promise<void>;
  }
): Promise<void>;

/**
 * 加载Excel工作簿。
 *
 * @param input - 输入数据，支持本地路径、URL地址、ArrayBuffer、Blob 或 Buffer。
 * @returns - 返回一个 Promise，解析为加载的工作簿对象。
 */
declare function loadWorkbook(input: Input): Promise<ExcelJS.Workbook>;

/**
 * 填充Excel模板。
 *
 * @param workbook - 工作簿对象。
 * @param workbookData - 模板数据，数组对象形式，键值对表示占位符和实际值。
 * @param parseImage - 是否解析图片，默认为 false。
 * @returns - 返回一个 Promise，解析为填充后的工作簿对象。
 */
declare function fillTemplate(
  workbook: ExcelJS.Workbook,
  workbookData: Data,
  parseImage?: boolean
): Promise<ExcelJS.Workbook>;

/**
 * 保存工作簿到文件。
 *
 * @param workbook - 工作簿对象。
 * @param output - 输出文件路径或文件名。
 * @returns - 返回一个 Promise，完成时无返回值。
 */
declare function saveWorkbook(workbook: ExcelJS.Workbook, output: string): Promise<void>;

/**
 * 获取自定义占位符单元格范围。
 *
 * @param worksheet - 工作表对象。
 * @param placeholder - 占位符字符串，默认为 "{{#placeholder}}"。
 * @param clearMatch - 是否清除匹配的占位符，默认为 true。
 * @returns - 返回占位符单元格的起始和结束位置，未找到时返回 null。
 */
declare function placeholderRange(
  worksheet: ExcelJS.Worksheet,
  placeholder?: string,
  clearMatch?: boolean
): { start: { row: number; col: number }; end: { row: number; col: number } } | null;

export { loadWorkbook, fillTemplate, saveWorkbook, placeholderRange, renderXlsxTemplate };
