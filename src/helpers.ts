// 是否在浏览器环境
const isBrowser = typeof window !== "undefined" && typeof document !== "undefined"

/**
 * 获取url文件
 * @param urlStr - url地址
 * @returns Promise<ArrayBuffer>
 */
async function fetchUrlFile(urlStr: string) {
  const url = new URL(urlStr)

  if (typeof fetch !== "undefined") {
    // 优先使用 fetch
    const response = await fetch(urlStr)
    if (!response.ok) {
      throw new Error(`Failed to fetch ${urlStr}, status code: ${response.status}`)
    }
    return await response.arrayBuffer()
  }
  else {
    // Node.js 回退
    const http = url.protocol === "https:" ? await import("node:https") : await import("node:http")

    return new Promise<ArrayBuffer>((resolve, reject) => {
      const request = http.get(urlStr, (response) => {
        if (response.statusCode !== 200) {
          reject(new Error(`Failed to fetch ${urlStr}, status code: ${response.statusCode}`))
          // 消费数据防止内存泄漏
          response.resume()
          return
        }

        const chunks: Uint8Array[] = []
        response.on("data", (chunk: Uint8Array) => chunks.push(chunk))
        response.on("end", async () => {
          try {
            const { Buffer } = await import("node:buffer")
            // 合并所有 chunk 成一个 Buffer
            const buffer = Buffer.concat(chunks)
            // 转换为 ArrayBuffer（安全切片）
            const arrayBuffer = buffer.buffer.slice(
              buffer.byteOffset,
              buffer.byteOffset + buffer.byteLength,
            )
            resolve(arrayBuffer)
          }
          catch (err) {
            reject(err)
          }
        })
      })

      request.on("error", reject)
      request.on("timeout", () => request.destroy())
      request.setTimeout(30_000) // 30秒超时
    })
  }
}

export { fetchUrlFile, isBrowser }
