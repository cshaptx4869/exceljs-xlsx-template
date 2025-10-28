import type { ImageRange } from "exceljs"
import type { IterationTag, RenderData, RenderInput, SheetDynamicMerges } from "./types"
import ExcelJS from "exceljs"
import { fetchUrlFile, isBrowser } from "./helpers"

/**
 * 加载工作簿
 * @param input - 输入数据，可以是本地路径、URL地址、ArrayBuffer、Blob、Buffer
 * @returns 加载完成的工作簿对象
 */
async function loadWorkbook(input: RenderInput) {
  const workbook = new ExcelJS.Workbook()
  const httpRegex = /^https?:\/\//
  if (isBrowser) {
    if (typeof input === "string" && httpRegex.test(input)) {
      await workbook.xlsx.load(await fetchUrlFile(input))
    }
    else if (input instanceof Blob || input instanceof ArrayBuffer) {
      await workbook.xlsx.load(input as ExcelJS.Buffer)
    }
    else {
      throw new TypeError("Unsupported input type in browser environment. Expected Blob, File, ArrayBuffer.")
    }
  }
  else {
    if (typeof input === "string") {
      if (httpRegex.test(input)) {
        await workbook.xlsx.load(await fetchUrlFile(input))
      }
      else {
        await workbook.xlsx.readFile(input)
      }
    }
    else if (input instanceof (await import("node:buffer")).Buffer || input instanceof ArrayBuffer) {
      await workbook.xlsx.load(input as ExcelJS.Buffer)
    }
    else if (input instanceof (await import("node:stream")).Stream) {
      await workbook.xlsx.read(input)
    }
    else {
      throw new TypeError(
        "Unsupported input type in Node.js environment. Expected file path, Buffer, ArrayBuffer, or Stream.",
      )
    }
  }
  return workbook
}

/**
 * 填充Excel模板
 * @param workbook - 工作簿对象
 * @param workbookData - 包含模板数据的数组对象
 * @param parseImage - 是否解析图片，默认为 false
 * @returns 填充完成的工作簿对象
 */
async function fillTemplate(workbook: ExcelJS.Workbook, workbookData: RenderData, parseImage = false) {
  // 第一步：复制行并替换占位符
  // 工作表待合并单元格信息
  const sheetDynamicMerges: SheetDynamicMerges = {}
  // NOTE 工作簿的sheetId是按工作表创建的顺序从1开始递增
  let sheetIndex = 0
  workbook.eachSheet((worksheet, sheetId) => {
    const sheetData = workbookData[sheetIndex++]
    if (!(sheetData && typeof sheetData === "object" && !Array.isArray(sheetData))) {
      return
    }
    // 普通标签替换和迭代标签信息收集
    const sheetIterTags: IterationTag[] = []
    worksheet.eachRow((row, rowNumber) => {
      // 行迭代字段
      const rowIterFields: string[] = []
      row.eachCell((cell) => {
        const cellType = cell.type
        // 字符串值
        if (cellType === ExcelJS.ValueType.String) {
          // @ts-expect-error 一定是String类型
          cell.value = processCellTags(cell.value, sheetData, sheetIterTags, rowIterFields, rowNumber)
        }
        // 富文本值
        else if (cellType === ExcelJS.ValueType.RichText) {
          // @ts-expect-error 一定是RichText类型
          cell.value.richText.forEach((item) => {
            item.text = processCellTags(item.text, sheetData, sheetIterTags, rowIterFields, rowNumber)
          })
        }
      })
    })
    // 迭代标签处理
    if (sheetIterTags.length === 0) {
      return
    }
    // 合并单元格信息
    // NOTE 合并信息是静态的，不会随着行增加而实时更新
    const sheetMerges = sheetMergeInfo(worksheet)
    // 迭代行并替换迭代字段占位符
    let iterOffset = 0
    sheetIterTags.forEach(({ iterStartRow, iterFieldNames, iterFieldName }, iterTagIndex) => {
      // 调整后的起始行
      const adjustedStartRow = iterStartRow + iterOffset
      const iterDataLength = sheetData[iterFieldName].length
      // 多行的情况下，需要复制多行
      if (iterDataLength > 1) {
        // 一次性复制多行
        // NOTE 复制的行不会复制合并信息
        worksheet.duplicateRow(adjustedStartRow, iterDataLength - 1, true)
        // 筛选出与当前模板行相关的合并单元格信息，并应用到其复制的行
        const merges = sheetMerges.filter((merge) => {
          return merge.start.row <= iterStartRow && merge.end.row >= iterStartRow
        })
        if (merges.length > 0) {
          if (!sheetDynamicMerges[sheetId]) {
            sheetDynamicMerges[sheetId] = []
          }
          // NOTE 在浏览器环境，动态增加的行会使其后面的行取消合并单元格
          const startFixIndex = isBrowser ? (iterTagIndex === 0 ? 1 : 0) : 1
          for (let i = startFixIndex; i < iterDataLength; i++) {
            for (const merge of merges) {
              sheetDynamicMerges[sheetId].push([
                merge.start.row + i + iterOffset,
                merge.start.col,
                merge.end.row + i + iterOffset,
                merge.end.col,
              ])
            }
          }
        }
      }
      // 替换迭代行中的占位符
      for (let i = 0; i < iterDataLength; i++) {
        const currentRow = worksheet.getRow(adjustedStartRow + i)
        // 遍历当前行的单元格
        currentRow.eachCell((cell) => {
          // 字符串值
          if (cell.type === ExcelJS.ValueType.String) {
            // 遍历单元格中的多个迭代字段
            iterFieldNames.forEach((iterField) => {
              // 单元格中包含当前迭代字段的占位符
              // @ts-expect-error 一定是String类型
              if (cell.value.includes(`{{@@${iterField}\.`)) {
                // 当前迭代字段索引数据
                const currentIterFieldData = sheetData[iterField][i]
                if (currentIterFieldData !== undefined) {
                  // 迭代字段数据
                  for (const field in currentIterFieldData) {
                    const placeholder = `{{@@${iterField}.${field}}}`
                    // @ts-expect-error 一定是String类型
                    if (cell.value.includes(placeholder)) {
                      // 完全匹配，替换单元格内容为迭代字段数据
                      // @ts-expect-error 一定是String类型
                      if (cell.value.length === placeholder.length && typeof currentIterFieldData[field] !== "object") {
                        cell.value = currentIterFieldData[field]
                        break
                      }
                      // 包含其他内容，部分替换为迭代字段数据
                      else {
                        // @ts-expect-error 一定是String类型
                        cell.value = cell.value.replace(
                          new RegExp(placeholder, "g"),
                          currentIterFieldData[field] ?? "",
                        )
                      }
                    }
                  }
                }
                else {
                  // 清空单元格内容
                  cell.value = null
                }
              }
            })
          }
          // TODO 迭代标签单元格为富文本值
        })
      }
      // 更新行号偏移量
      iterOffset += iterDataLength - 1
    })
    // 修正在浏览器环境，动态增加的行会使其后面的行取消合并单元格
    if (isBrowser) {
      const iterRows = sheetIterTags.map(({ iterStartRow }) => iterStartRow)
      sheetMerges.forEach((merge) => {
        // 迭代后的偏移行
        let mergeOffset = 0
        sheetIterTags.forEach(({ iterStartRow, iterFieldName }) => {
          if (Array.isArray(sheetData[iterFieldName])) {
            if (!iterRows.includes(merge.start.row) && merge.start.row > iterStartRow) {
              mergeOffset += sheetData[iterFieldName].length - 1
            }
          }
        })
        if (mergeOffset) {
          if (!sheetDynamicMerges[sheetId]) {
            sheetDynamicMerges[sheetId] = []
          }
          sheetDynamicMerges[sheetId].push([
            merge.start.row + mergeOffset,
            merge.start.col,
            merge.end.row + mergeOffset,
            merge.end.col,
          ])
        }
      })
    }
  })

  // 第二步：动态行单元格合并处理
  if (Object.keys(sheetDynamicMerges).length > 0) {
    // 将工作簿保存到内存中的缓冲区
    const buffer = await workbook.xlsx.writeBuffer()
    // 从缓冲区重新加载工作簿
    await workbook.xlsx.load(buffer)
    // 处理合并单元格
    workbook.eachSheet((worksheet, sheetId) => {
      if (sheetDynamicMerges[sheetId]) {
        sheetDynamicMerges[sheetId].forEach((merge) => {
          try {
            worksheet.mergeCells(merge)
          }
          catch {
            console.warn(`Fail to merge cells ${merge}`)
          }
        })
      }
    })
  }

  // 第三步：填充图片
  parseImage && (await fillImage(workbook))

  return workbook
}

/**
 * 保存工作簿到文件
 * @param workbook - 工作簿对象
 * @param output - 输出文件路径或文件名
 * @returns Promise<void>
 */
async function saveWorkbook(workbook: ExcelJS.Workbook, output: string) {
  if (isBrowser) {
    const buffer = await workbook.xlsx.writeBuffer()
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
    const link = document.createElement("a")
    link.href = URL.createObjectURL(blob)
    link.download = output
    link.click()
    URL.revokeObjectURL(link.href)
  }
  else {
    await workbook.xlsx.writeFile(output)
  }
}

/**
 * 获取自定义占位符单元格范围
 * @param worksheet - 工作表对象
 * @param placeholder - 占位符字符串，默认为 "{{#placeholder}}"
 * @param clearMatch - 是否清除占位符，默认为 true
 * @returns 占位符单元格范围信息
 */
function placeholderRange(worksheet: ExcelJS.Worksheet, placeholder = "{{#placeholder}}", clearMatch = true) {
  let result = null
  const sheetMerges = sheetMergeInfo(worksheet)
  // 遍历每一行
  outer: for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
    const row = worksheet.getRow(rowNumber)
    // 遍历每个单元格
    for (let colNumber = 1; colNumber <= row.cellCount; colNumber++) {
      const cell = row.getCell(colNumber)
      // 单元格值中是否包含占位符
      if (typeof cell.value === "string" && cell.value.includes(`${placeholder}`)) {
        const info = sheetMerges.find((merge) => {
          return merge.start.row === rowNumber && merge.start.col === colNumber
        })
        result = info ?? {
          start: { row: rowNumber, col: colNumber },
          end: { row: rowNumber, col: colNumber },
        }
        // 去除占位符
        if (clearMatch) {
          cell.value = cell.value.replace(new RegExp(`${placeholder}`, "g"), "")
        }
        // 跳出循环
        break outer
      }
    }
  }
  return result
}

/**
 * 处理单元格标签(普通标签替换和迭代标签信息收集)
 * @param target 单元格值
 * @param worksheetData 工作表数据
 * @param iterationTags 迭代标签信息
 * @param iterFieldNames 行迭代字段
 * @param rowNumber 行号
 * @returns 处理后的单元格值
 */
function processCellTags(target: string, worksheetData: Record<string, any>, iterationTags: IterationTag[], iterFieldNames: string[], rowNumber: number) {
  // 普通标签占位符替换
  if (/\{\{\w+(\.\w+)*\}\}/.test(target)) {
    // 允许单元格中有多个普通标签占位符
    const placeholders = target.match(/\{\{\w+(\.\w+)*\}\}/g) ?? []
    placeholders.forEach((placeholder) => {
      // 支持 xxx.xxx 格式
      const fields = placeholder.slice(2, -2).split(".")
      let value = JSON.parse(JSON.stringify(worksheetData))
      let isMatched = true
      for (let i = 0; i < fields.length; i++) {
        if (fields[i] in value) {
          value = value[fields[i]]
        }
        else {
          isMatched = false
          break
        }
      }
      // 数据匹配成功
      if (isMatched) {
        if (target.length === placeholder.length && typeof value !== "object") {
          // 无其他多余字符的，直接替换标签内容
          target = value
        }
        else {
          // 替换占位符部分内容
          target = target.replace(placeholder, value ?? "")
        }
      }
    })
  }
  // 迭代标签信息搜集
  else if (/\{\{@@\w+\.\w+\}\}/.test(target)) {
    // 单元格中仅匹配一个迭代标签占位符
    // TODO 单元格中含多个不同的迭代标签
    const iterFieldName = target.match(/\{\{@@(\w+)\.\w+\}\}/)?.[1] ?? ""
    // 数据存在且为数组类型
    if (
      iterFieldName in worksheetData
      && Array.isArray(worksheetData[iterFieldName])
      && worksheetData[iterFieldName].length > 0
    ) {
      if (iterFieldNames.length === 0) {
        iterFieldNames.push(iterFieldName)
        iterationTags.push({ iterStartRow: rowNumber, iterFieldNames, iterFieldName })
      }
      else {
        if (!iterFieldNames.includes(iterFieldName)) {
          iterFieldNames.push(iterFieldName)
          // 迭代标签字段长度不一致，取最长的（后续按最大长度复制行）
          const lastIterationTag = iterationTags[iterationTags.length - 1]
          if (worksheetData[iterFieldName].length > worksheetData[lastIterationTag.iterFieldName].length) {
            lastIterationTag.iterFieldName = iterFieldName
          }
        }
      }
    }
  }
  return target
}

/**
 * 获取工作表合并信息
 * @param worksheet 工作表
 * @returns 工作表合并单元格信息
 */
function sheetMergeInfo(worksheet: ExcelJS.Worksheet) {
  return worksheet.model.merges.map((merge) => {
    // C30:D30
    const [startCell, endCell] = merge.split(":")
    const startCellInfo = worksheet.getCell(startCell)
    const endCellInfo = worksheet.getCell(endCell)
    return {
      start: { row: Number(startCellInfo.row), col: Number(startCellInfo.col) },
      end: { row: Number(endCellInfo.row), col: Number(endCellInfo.col) },
    }
  })
}

/**
 * 填充图片
 * @param workbook
 */
async function fillImage(workbook: ExcelJS.Workbook) {
  const filledImageMap = new Map()
  const invalidImageSet = new Set()
  // eslint-disable-next-line regexp/optimal-quantifier-concatenation
  const urlRegex = /https?:\/\/\S+(?:\.(jpe?g|png|gif))?/i
  const base64Regex = /data:image\/(jpeg|gif|png);base64,\S+/i
  // NOTE eachSheet、eachRow、eachCell都是同步方法，不会等待异步操作完成
  // 遍历每个工作表
  for (let i = 0; i < workbook.worksheets.length; i++) {
    const worksheet = workbook.worksheets[i]
    const sheetMerges = sheetMergeInfo(worksheet)
    // 遍历每一行
    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber)
      // 遍历每个单元格
      for (let colNumber = 1; colNumber <= row.cellCount; colNumber++) {
        const cell = row.getCell(colNumber)
        if (typeof cell.value !== "string") {
          continue
        }
        let targetRegex = null
        let imageId = 0
        // URL图片
        if (urlRegex.test(cell.value)) {
          targetRegex = urlRegex
          const matches = cell.value.match(urlRegex)
          const imageUrl = matches?.[0] ?? ""
          const imageExt = matches?.[1] ?? "png"
          if (invalidImageSet.has(imageUrl)) {
            continue
          }
          if (filledImageMap.has(imageUrl)) {
            imageId = filledImageMap.get(imageUrl)
          }
          else {
            let fileContent = null
            try {
              fileContent = await fetchUrlFile(imageUrl)
            }
            catch {
              invalidImageSet.add(imageUrl)
              console.warn(`Fail to load image ${imageUrl}`)
              continue
            }
            // 将图片添加到工作簿中
            imageId = workbook.addImage({
              buffer: fileContent,
              extension: (imageExt === "jpg" ? "jpeg" : imageExt) as "png" | "jpeg" | "gif",
            })
            filledImageMap.set(imageUrl, imageId)
          }
        }
        // Base64图片
        else if (base64Regex.test(cell.value)) {
          targetRegex = base64Regex
          const matches = cell.value.match(base64Regex)!
          const imageData = matches[0]
          const imageExt = matches[1]
          if (filledImageMap.has(imageData)) {
            imageId = filledImageMap.get(imageData)
          }
          else {
            imageId = workbook.addImage({
              base64: imageData,
              extension: imageExt as "png" | "jpeg" | "gif",
            })
            filledImageMap.set(imageData, imageId)
          }
        }
        if (!targetRegex) {
          continue
        }
        // 将图片添加到工作表中
        const merge = sheetMerges.find((merge) => {
          return merge.start.row === rowNumber && merge.start.col === colNumber
        })
        // 坐标系基于零，A1 的左上角将为 {col：0，row：0}，右下角为 {col：1，row：1}
        worksheet.addImage(imageId, {
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
        } as ImageRange)
        // 去除图片地址
        cell.value = cell.value.replace(targetRegex, "")
      }
    }
  }

  return workbook
}

export {
  fillTemplate,
  loadWorkbook,
  placeholderRange,
  saveWorkbook,
}
