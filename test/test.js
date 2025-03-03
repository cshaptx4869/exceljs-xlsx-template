const path = require("path");
const fs = require("fs");
const { fillTemplate, loadWorkbook, saveWorkbook } = require("../src/index.js");

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
  // 方式一：本地文件
  // const workbook = await loadWorkbook(input);
  // 方式二：从文件读取二进制数据到缓冲区
  const buffer = fs.readFileSync(input);
  const workbook = await loadWorkbook(buffer);
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
