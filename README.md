# exceljs-xlsx-template

基于 [exceljs](https://github.com/exceljs/exceljs) 库的 .xlsx 模板文件填充引擎。理论上支持 exceljs 库的所有 [api](https://github.com/exceljs/exceljs/blob/master/README_zh.md#目录)。

- 普通标签占位符格式：`{{xxx}}`、`{{xxx.xxx}}`
- 迭代标签占位符格式：`{{@@xxx.xxx}}`

> 支持浏览器和 node.js 环境下使用。可参考 test 目录下的 test.html 或 test.js。

![input](https://github.com/user-attachments/assets/31c05045-e3c1-49a6-ab7d-9f1d72b91710)

![output](https://github.com/user-attachments/assets/98853096-8674-4d09-bd88-e09bcc9547b2)
