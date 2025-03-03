const path = require("path");
const TerserPlugin = require("terser-webpack-plugin");

module.exports = {
  entry: "./src/index.js",
  output: {
    filename: "bundle.js",
    path: path.resolve(__dirname, "dist"),
    library: "ExceljsXlsxTemplate", // 全局变量名（可选）
    libraryTarget: "umd", // 输出为 UMD 格式
    clean: true, // 清理输出目录
  },
  mode: "production", // 生产模式
  optimization: {
    minimize: true,
    minimizer: [
      new TerserPlugin({
        extractComments: false, // 禁用 LICENSE 文件生成
      }),
    ],
  },
};
