import path from "node:path"
import url from "node:url"
// eslint-disable-next-line antfu/no-import-dist
import { placeholderRange, renderXlsxTemplate } from "../dist/index.esm.js"

const __dirname = path.dirname(url.fileURLToPath(import.meta.url))

const xlsxFile = path.join(__dirname, "assets", "template.xlsx")
const officialsealFile = path.join(__dirname, "assets", "officialseal.png")
const imageUrl = "https://s2.loli.net/2025/03/07/ELZY594enrJwF7G.png"
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
    user: {
      last_name: "Doe",
      first_name: "John",
    },
    phone: "00874****",
    invoice_date: "15/05/2008",
    items: [
      {
        name: "description",
        unit_price: 300,
      },
      {
        name: "HTML",
        unit_price: 400,
      },
    ],
    subtotal: 700,
    tax: 140,
    grand_total: 840,
  },
]
const output = path.join(__dirname, "output", `${Date.now()}.xlsx`)

// 渲染Xlsx模板
renderXlsxTemplate(xlsxFile, data, output, {
  parseImage: true,
  beforeSave(workbook) {
    // 获取工作表
    const worksheet = workbook.getWorksheet("新报关单")
    if (worksheet) {
      // 获取印章占位符位置信息
      const range = placeholderRange(worksheet, "{{#officialseal}}")
      if (range) {
        // 将图片添加到工作簿
        const imageId = workbook.addImage({
          filename: officialsealFile,
          extension: "png",
        })
        // 插入图片到表格中
        worksheet.addImage(imageId, {
          tl: { col: range.start.col, row: range.start.row - 4 },
          ext: { width: 200, height: 200 },
        })
      }
    }
  },
})
  .then(() => {
    console.log("🚀 ~ output:", output)
  })
  .catch((error) => {
    console.error("Error processing Excel file:", error)
  })
