# exceljs-xlsx-template

## 简介

基于 exceljs 库的 `.xlsx` 模板文件填充引擎。理论上支持 [exceljs](https://github.com/exceljs/exceljs) 库的所有 [api](https://github.com/exceljs/exceljs/blob/master/README_zh.md#目录)。

- 普通标签占位符格式：`{{xxx}}`、`{{xxx.xxx}}`
- 迭代标签占位符格式：`{{@@xxx.xxx}}`

## API

核心函数：

- `loadWorkbook(input: RenderInput) `：加载工作簿
- `fillTemplate(workbook: ExcelJS.Workbook, workbookData: RenderData, options: FillOptions = {})`：填充 Xlsx 模板
- `saveWorkbook(workbook: ExcelJS.Workbook, output: string) `：保存工作簿到文件
- `renderXlsxTemplate(input: RenderInput, workbookDat: RenderData, output: string, options: RenderOptions = {}) `：渲染 Xlsx 模板
- `placeholderRange(worksheet: ExcelJS.Worksheet, placeholder = "{{#placeholder}}", clearMatch = true) `：获取自定义占位符单元格范围
- `fetchUrlFile(urlStr: string)`：获取 url 文件
- `downloadBlob(blob: Blob, fileName: string)`：下载 blob 文件

## 示例

> 参考 test 目录下的 [test.html](https://github.com/cshaptx4869/exceljs-xlsx-template/blob/main/test/test.html) 或 [test.js](https://github.com/cshaptx4869/exceljs-xlsx-template/blob/main/test/test.js)

```vue
<script setup lang="ts">
import { BlobReader, BlobWriter, ZipWriter } from "@zip.js/zip.js"
import { downloadBlob, fetchUrlFile, fillTemplate, loadWorkbook, placeholderRange, renderXlsxTemplate } from "exceljs-xlsx-template"

function handleXlsxTemplate(isZip = false) {
  const xlsxFile = "https://raw.githubusercontent.com/cshaptx4869/exceljs-xlsx-template/refs/heads/main/test/assets/template.xlsx"
  const officialsealFile = "https://raw.githubusercontent.com/cshaptx4869/exceljs-xlsx-template/refs/heads/main/test/assets/officialseal.png"
  const imageUrl = officialsealFile
  const data = [
    {
      name: "John",
      items: [
        { no: "No.1", name: "JavaScript", quantity: 300 },
        { no: "No.2", name: "CSS", quantity: 400 },
        { no: "No.3", name: "HTML", quantity: 500 },
        { no: "No.4", name: "Node.js", quantity: 600 },
        { no: "No.5", name: "Three.js", quantity: 700 },
        { no: "No.6", name: "Vue", quantity: 800 },
        { no: "No.7", name: "React", quantity: 900 },
        { no: "No.8", name: "Angular", quantity: 1000 },
        { no: "No.9", name: "UniApp", quantity: 1100 }
      ],
      projects: [
        { name: "Project 1", description: "Description 1", image: imageUrl },
        { name: "Project 2", description: "Description 2", image: imageUrl },
        { name: "Project 3", description: "Description 3", image: imageUrl }
      ]
    },
    {
      invoice_number: "54548",
      user: {
        first_name: "John",
        last_name: "Chen"
      },
      phone: "00874****",
      invoice_date: "15/05/2008",
      items: [
        { name: "PHP", unit_price: 300 },
        { name: "HTML", unit_price: 400 }
      ],
      subtotal: 700,
      tax: 140,
      grand_total: 840
    }
  ]

  try {
    if (isZip !== true) {
      renderXlsxTemplate(xlsxFile, data, `${Date.now()}.xlsx`, {
        // 解析图片地址
        parseImage: true,
        // 保留迭代行静态内容
        keepDupRowStatic: true,
        // 保存前处理
        beforeSave: async (workbook) => {
          const worksheet = workbook.getWorksheet("新报关单")
          if (worksheet) {
            // 获取印章占位符位置信息
            const range = placeholderRange(worksheet, "{{#officialseal}}")
            if (range) {
              // 加载图片印章
              const officialsealArrayBuffer = await fetchUrlFile(officialsealFile)
              // 将图片添加到工作簿
              const imageId = workbook.addImage({
                buffer: officialsealArrayBuffer,
                extension: "png"
              })
              // 插入图片到表格中
              worksheet.addImage(imageId, {
                tl: { col: range.start.col, row: range.start.row - 4 },
                ext: { width: 200, height: 200 }
              })
            }
          }
        }
      })
    } else {
      loadWorkbook(xlsxFile).then(async (workbook) => {
        // 填充模板
        await fillTemplate(workbook, data, {
          parseImage: true,
          keepDupRowStatic: true
        })
        const worksheet = workbook.getWorksheet("新报关单")
        if (worksheet) {
          // 获取印章占位符位置信息
          const range = placeholderRange(worksheet, "{{#officialseal}}")
          if (range) {
            // 加载图片印章
            const officialsealArrayBuffer = await fetchUrlFile(officialsealFile)
            // 将图片添加到工作簿
            const imageId = workbook.addImage({
              buffer: officialsealArrayBuffer,
              extension: "png"
            })
            // 插入图片到表格中
            worksheet.addImage(imageId, {
              tl: { col: range.start.col, row: range.start.row - 4 },
              ext: { width: 200, height: 200 }
            })
          }
        }
        const buffer = await workbook.xlsx.writeBuffer()
        const xlsxBlob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })

        // 打包
        const zipWriter = new ZipWriter(new BlobWriter())
        // NOTE: 文件名不能包含 / 否则会当做目录分割符
        zipWriter.add(`${Date.now()}.xlsx`, new BlobReader(xlsxBlob))
        // 关闭压缩包
        const zipBlob = await zipWriter.close()

        // 下载
        downloadBlob(zipBlob, `${Date.now()}.zip`)
      })
    }
  } catch (error) {
    console.error("Error processing Excel file:", error)
  }
}
</script>

<template>
  <div>
    <button type="button" @click="handleXlsxTemplate()">
      渲染xlsx模板
    </button>
    <button type="button" @click="handleXlsxTemplate(true)">
      渲染xlsx模板并打包
    </button>
  </div>
</template>
```

![input](https://raw.githubusercontent.com/cshaptx4869/Stash/PicGo/425918263-31c05045-e3c1-49a6-ab7d-9f1d72b91710.png)

![output](https://raw.githubusercontent.com/cshaptx4869/Stash/PicGo/425918262-98853096-8674-4d09-bd88-e09bcc9547b2.png)
