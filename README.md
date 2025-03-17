# exceljs-xlsx-template

基于 [exceljs](https://github.com/exceljs/exceljs) 库的 .xlsx 模板文件填充引擎。理论上支持 exceljs 库的所有 [api](https://github.com/exceljs/exceljs/blob/master/README_zh.md#目录)。

- 单标签占位符格式：`{{xxx}}`
- 迭代标签占位符格式：`{{xxx.xxx}}`

## 接口

```typescript
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
```

## 示例

支持浏览器和 node.js 环境下使用。可参考 test 目录下的 test.html 或 test.js。

> vue3

```vue
<template>
  <button type="button" @click="handleXlsxTemplate">xlsx模板填充</button>
</template>

<script setup lang="ts">
import { fillTemplate, loadWorkbook, saveWorkbook, placeholderRange } from "exceljs-xlsx-template";

async function handleXlsxTemplate() {
  const imageUrl = "https://s2.loli.net/2025/03/07/ELZY594enrJwF7G.png";
  const data = [
    {
      name: "John",
      items: [
        { no: "No.1", name: "JavaScript" },
        { no: "No.2", name: "CSS" },
        { no: "No.3", name: "HTML" },
        { no: "No.4", name: "Node.js" },
        { no: "No.5", name: "Three.js" },
        { no: "No.6", name: "Vue" },
        { no: "No.7", name: "React" },
        { no: "No.8", name: "Angular" },
        { no: "No.9", name: "UniApp" },
      ],
      projects: [
        { name: "Project 1", description: "Description 1", image: imageUrl },
        { name: "Project 2", description: "Description 2", image: imageUrl },
        { name: "Project 3", description: "Description 3", image: imageUrl },
      ],
    },
  ];
  // 加载Excel文件
  const workbook = await loadWorkbook("http://example.com/test/assets/template.xlsx");
  // 填充模板
  await fillTemplate(workbook, data, true);
  // 加载印章图片
  const officialsealResponse = await fetch(imageUrl);
  if (!officialsealResponse.ok)
    throw new Error(`Failed to download image file, status code: ${officialsealResponse.status}`);
  const officialsealArrayBuffer = await officialsealResponse.arrayBuffer();
  // 遍历每个工作表
  workbook.eachSheet((worksheet, sheetId) => {
    if (sheetId === 1) {
      // 将图片添加到工作簿
      const imageId = workbook.addImage({
        buffer: officialsealArrayBuffer,
        extension: "png",
      });
      // 获取印章占位符位置信息
      const range = placeholderRange(worksheet, "{{#officialseal}}");
      if (range) {
        // 插入图片到表格中
        worksheet.addImage(imageId, {
          tl: { col: range.start.col, row: range.start.row - 4 },
          ext: { width: 200, height: 200 },
        });
      }
    }
  });
  // 保存为新的 Excel 文件
  await saveWorkbook(workbook, `${Date.now()}.xlsx`);
}
</script>
```

> node.js

```javascript
const path = require("path");
const fs = require("fs");
const { fillTemplate, loadWorkbook, saveWorkbook, placeholderRange } = require("exceljs-xlsx-template");

const xlsxFile = path.join(__dirname, "assets", "template.xlsx");
const officialsealFile = path.join(__dirname, "assets", "officialseal.png");
const imageUrl = "https://s2.loli.net/2025/03/07/ELZY594enrJwF7G.png";
const data = [
  {
    name: "John",
    items: [
      { no: "No.1", name: "JavaScript" },
      { no: "No.2", name: "CSS" },
      { no: "No.3", name: "HTML" },
      { no: "No.4", name: "Node.js" },
      { no: "No.5", name: "Three.js" },
      { no: "No.6", name: "Vue" },
      { no: "No.7", name: "React" },
      { no: "No.8", name: "Angular" },
      { no: "No.9", name: "UniApp" },
    ],
    projects: [
      { name: "Project 1", description: "Description 1", image: imageUrl },
      { name: "Project 2", description: "Description 2", image: imageUrl },
      { name: "Project 3", description: "Description 3", image: imageUrl },
    ],
  },
];

async function main() {
  // 加载Excel文件
  const workbook = await loadWorkbook(xlsxFile);
  // 填充模板
  await fillTemplate(workbook, data, true);
  // 遍历每个工作表
  workbook.eachSheet((worksheet, sheetId) => {
    if (sheetId === 1) {
      // 将图片添加到工作簿
      const imageId = workbook.addImage({
        filename: officialsealFile,
        extension: "png",
      });
      // 获取印章占位符位置信息
      const range = placeholderRange(worksheet, "{{#officialseal}}");
      if (range) {
        // 插入图片到表格中
        worksheet.addImage(imageId, {
          tl: { col: range.start.col, row: range.start.row - 4 },
          ext: { width: 200, height: 200 },
        });
      }
    }
  });
  // 保存为新的 Excel 文件
  const outputDir = path.join(__dirname, "output");
  !fs.existsSync(outputDir) && fs.mkdirSync(outputDir);
  const output = path.join(outputDir, `${Date.now()}.xlsx`);
  await saveWorkbook(workbook, output);
  return output;
}

main()
  .then((res) => {
    console.log("🚀 ~ output:", res);
  })
  .catch((error) => {
    console.error("Error processing Excel file:", error);
  });
```

---

![input](https://github.com/user-attachments/assets/72329b94-004b-4587-a7ab-5a38864bea17)

![output](https://github.com/user-attachments/assets/4aa5fa65-ca8a-4c6c-ba90-101137450ad7)
