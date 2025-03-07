const path = require("path");
const fs = require("fs");
const { fillTemplate, loadWorkbook, saveWorkbook, placeholderRange } = require("../src/index.js");

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
