import type { FillOptions, RenderData, RenderInput, RenderOptions } from "./types"
import { fillTemplate, loadWorkbook, placeholderRange, saveWorkbook } from "./core"
import { fetchUrlFile } from "./helpers"

/**
 * 渲染Xlsx模板
 * @param input - 输入数据，可以是本地路径、URL地址、ArrayBuffer、Blob、Buffer
 * @param data - 包含模板数据的数组对象
 * @param output - 输出文件路径或文件名
 * @param options 配置项
 * @returns Promise<void>
 */
async function renderXlsxTemplate(input: RenderInput, data: RenderData, output: string, options: RenderOptions = {}) {
  const { beforeSave, ...fillOptions } = options
  const workbook = await loadWorkbook(input)
  await fillTemplate(workbook, data, fillOptions)
  if (typeof beforeSave === "function") {
    await beforeSave(workbook)
  }
  await saveWorkbook(workbook, output)
}

// 导出核心函数
export {
  fetchUrlFile,
  fillTemplate,
  loadWorkbook,
  placeholderRange,
  renderXlsxTemplate,
  saveWorkbook,
}

// 导出类型定义
export type {
  FillOptions,
  RenderData,
  RenderInput,
  RenderOptions,
}
