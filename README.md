# exceljs-xlsx-template

基于 [exceljs](https://github.com/exceljs/exceljs) 库的 .xlsx 模板文件填充引擎。理论上支持 exceljs 库的所有 [api](https://github.com/exceljs/exceljs/blob/master/README_zh.md#目录)。

- 单标签占位符格式：`{{xxx}}`
- 迭代标签占位符格式：`{{xxx.xxx}}`

接口：

```typescript
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
 * @param {boolean} parseImage
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
 * @param {string} output
 * @returns {Promise<void>}
 */
declare function saveWorkbook(workbook: ExcelJS.Workbook, output: string): Promise<void>;
```

示例：

> 详见test目录下的test.js和test.html

```javascript
const path = require("path");
const fs = require("fs");
const { fillTemplate, loadWorkbook, saveWorkbook } = require("exceljs-xlsx-template");

const input = path.join(__dirname, "assets", "template.xlsx");
const officialseal = path.join(__dirname, "assets", "officialseal.png");
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
      { name: "Project 1", description: "Description 1" },
      { name: "Project 2", description: "Description 2" },
      { name: "Project 3", description: "Description 3" },
    ],
  },
];

async function main() {
  // 加载工作簿
  const workbook = await loadWorkbook(input);
  // 填充模板
  await fillTemplate(workbook, data);
  // 添加图片印章
  const imageId = workbook.addImage({
    filename: officialseal,
    extension: "png",
  });
  workbook.eachSheet((worksheet, sheetId) => {
    // 第1张sheet表添加印章
    if (sheetId === 1) {
      // 获取表格的最后一行最后一列
      const lastRow = worksheet.lastRow;
      const lastColumn = worksheet.lastColumn;
      // 插入图片到表格中
      worksheet.addImage(imageId, {
        // 左上角位置
        tl: { col: lastColumn.number / 2, row: lastRow.number - 8 },
        ext: { width: 200, height: 200 },
      });
    }
  });
  // 保存为新的 Excel 文件
  const output = path.join(__dirname, "output", `${Date.now()}.xlsx`);
  await saveWorkbook(workbook, output);
  return output;
}

main()
  .then((res) => {
    console.log("🚀 ~ file:", res);
  })
  .catch((error) => {
    console.error("Error processing Excel file:", error);
  });
```

