import type ExcelJS from "exceljs"
import type { Buffer } from "node:buffer"

export type RenderInput = string | ArrayBuffer | Blob | Buffer

export type RenderData = Record<string, any>[]

export interface RenderOptions {
  parseImage?: boolean
  beforeSave?: (workbook: ExcelJS.Workbook) => void | Promise<void>
}

export interface IterationTag {
  iterStartRow: number
  iterFieldNames: string[]
  iterFieldName: string
}

export type SheetDynamicMerges = Record<string, [number, number, number, number][]>
