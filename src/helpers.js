"use strict";

// 是否在浏览器环境
const isBrowser = typeof window !== "undefined" && typeof document !== "undefined";

/**
 * 获取url文件
 * @param {string} url
 * @returns {Promise<Blob|Buffer>}
 */
async function fetchUrlFile(url) {
  if (isBrowser) {
    try {
      const response = await fetch(url);
      if (!response.ok) {
        throw new Error(`Failed to fetch ${url}, status code: ${response.status}`);
      }
      return response.blob();
    } catch (error) {
      throw new Error(`Error fetching ${url}: ${error.message}`);
    }
  } else {
    const { get } = /^https:\/\//.test(url) ? require("https") : require("http");
    return new Promise((resolve, reject) => {
      get(url, (response) => {
        if (response.statusCode !== 200) {
          reject(new Error(`Failed to fetch ${url}, status code: ${response.statusCode}`));
          return;
        }
        const chunks = [];
        response.on("data", (chunk) => chunks.push(chunk));
        response.on("end", () => resolve(Buffer.concat(chunks)));
      }).on("error", (err) => reject(err));
    });
  }
}

module.exports = { isBrowser, fetchUrlFile };
