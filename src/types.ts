import type ExcelJS from "exceljs"
import type { Buffer } from "node:buffer"
import type { Stream } from "node:stream"

export type RenderInput = string | ArrayBuffer | Blob | Buffer | Stream

export type RenderData = Record<string, any>[]

export interface RenderOptions {
  parseImage?: boolean
  keepDupRowStatic?: boolean
  beforeSave?: (workbook: ExcelJS.Workbook) => void | Promise<void>
}

export type FillOptions = Omit<RenderOptions, "beforeSave">

export interface IterationTag {
  iterStartRow: number
  iterFieldNames: string[]
  iterFieldName: string
}

export type SheetDynamicMerges = Record<string, [number, number, number, number][]>
