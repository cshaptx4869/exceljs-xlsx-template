# exceljs-xlsx-template

基于 [exceljs](https://github.com/exceljs/exceljs) 库的 .xlsx 模板文件填充引擎。理论上支持 exceljs 库的所有 [api](https://github.com/exceljs/exceljs/blob/master/README_zh.md#目录)。

- 普通标签占位符格式：`{{xxx}}`、`{{xxx.xxx}}`
- 迭代标签占位符格式：`{{@@xxx.xxx}}`

> 支持浏览器和 node.js 环境下使用。可参考 test 目录下的 [test.html](https://github.com/cshaptx4869/exceljs-xlsx-template/blob/main/test/test.html) 或 [test.js](https://github.com/cshaptx4869/exceljs-xlsx-template/blob/main/test/test.js)。

```vue
<template>
  <button type="button" @click="handleXlsxTemplate">渲染xlsx模板</button>
</template>

<script setup lang="ts">
import { renderXlsxTemplate, placeholderRange } from "exceljs-xlsx-template";

function handleXlsxTemplate() {
  const xlsxFile =
    "https://raw.githubusercontent.com/cshaptx4869/exceljs-xlsx-template/refs/heads/main/test/assets/template.xlsx";
  const officialsealFile =
    "https://raw.githubusercontent.com/cshaptx4869/exceljs-xlsx-template/refs/heads/main/test/assets/officialseal.png";
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
    {
      invoice_number: "54548",
      last_name: "John",
      first_name: "Doe",
      phone: "00874****",
      invoice_date: "15/05/2008",
      items: [
        { name: "description", unit_price: 300 },
        { name: "HTML", unit_price: 400 },
      ],
      subtotal: 700,
      tax: 140,
      grand_total: 840,
    },
  ];

  try {
    renderXlsxTemplate(xlsxFile, data, `${Date.now()}.xlsx`, {
      parseImage: true,
      async beforeSave(workbook) {
        // 获取工作表
        const worksheet = workbook.getWorksheet("新报关单");
        if (worksheet) {
          // 加载图片印章
          const officialsealRresponse = await fetch(officialsealFile);
          if (!officialsealRresponse.ok) {
            console.error(`Failed to download image file, status code: ${officialsealRresponse.status}`);
            return;
          }
          const officialsealArrayBuffer = await officialsealRresponse.arrayBuffer();
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
      },
    });
  } catch (error) {
    console.error("Error processing Excel file:", error);
  }
}
</script>
```

![input](https://github.com/user-attachments/assets/31c05045-e3c1-49a6-ab7d-9f1d72b91710)

![output](https://github.com/user-attachments/assets/98853096-8674-4d09-bd88-e09bcc9547b2)
