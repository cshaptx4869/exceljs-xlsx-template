import fs from "node:fs"
import commonjs from "@rollup/plugin-commonjs"
import nodeResolve from "@rollup/plugin-node-resolve"
import terser from "@rollup/plugin-terser"
import typescript from "@rollup/plugin-typescript"

const pkg = JSON.parse(fs.readFileSync("./package.json", "utf-8"))

export default {
  input: "src/index.ts",
  output: [
    // CommonJS 格式
    {
      file: pkg.main,
      format: "cjs",
      sourcemap: true,
      exports: "named",
    },
    // ES Module 格式
    {
      file: pkg.module,
      format: "esm",
      sourcemap: true,
    },
    // UMD 格式 (浏览器全局使用)
    {
      file: pkg.browser,
      format: "umd",
      sourcemap: true,
      exports: "named",
      name: "ExceljsXlsxTemplate", // 全局变量名
      globals: {
        // 如果有外部依赖，需要指定全局变量名
        exceljs: "ExcelJS",
      },
      plugins: [
        terser(), // 压缩代码
      ],
    },
  ],
  plugins: [
    nodeResolve(),
    commonjs(),
    typescript({
      tsconfig: "./tsconfig.json",
      declaration: true,
      declarationDir: "dist/types",
      emitDeclarationOnly: true,
    }),
  ],
  external: [...Object.keys(pkg.dependencies || {})], // 声明为外部依赖
}
