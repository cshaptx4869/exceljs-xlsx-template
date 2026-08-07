// 是否在浏览器环境
const isBrowser = typeof window !== "undefined" && typeof document !== "undefined"

// Node.js 模块惰性缓存（首次解析后缓存导出值，后续直接返回，跳过 import() 调用）
let _Buffer: typeof import("node:buffer").Buffer | null = null
let _Stream: typeof import("node:stream").Stream | null = null
let _Https: typeof import("node:https") | null = null
let _Http: typeof import("node:http") | null = null

/**
 * 获取 Node.js Buffer 构造函数
 */
async function getBuffer() {
  if (!_Buffer) {
    _Buffer = (await import("node:buffer")).Buffer
  }
  return _Buffer
}

/**
 * 获取 Node.js Stream 构造函数
 */
async function getStream() {
  if (!_Stream) {
    _Stream = (await import("node:stream")).Stream
  }
  return _Stream
}

/**
 * 获取 Node.js https 模块
 */
async function getHttps() {
  if (!_Https) {
    _Https = await import("node:https")
  }
  return _Https
}

/**
 * 获取 Node.js http 模块
 */
async function getHttp() {
  if (!_Http) {
    _Http = await import("node:http")
  }
  return _Http
}

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
    const http = url.protocol === "https:" ? await getHttps() : await getHttp()

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
            const Buffer = await getBuffer()
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
      request.on("timeout", () => {
        request.destroy()
        reject(new Error(`Request timeout for ${urlStr}`))
      })
      request.setTimeout(30_000) // 30秒超时
    })
  }
}

export {
  fetchUrlFile,
  getBuffer,
  getHttp,
  getHttps,
  getStream,
  isBrowser,
}
