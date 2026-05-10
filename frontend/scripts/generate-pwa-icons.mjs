/**
 * 构建前生成占位 PWA PNG（无设计稿时可用）。替换品牌图：将自有 logo.png 放到 public/ 并同步更新 vite.config.js 中的 manifest.icons。
 */
import fs from 'fs'
import path from 'path'
import { fileURLToPath } from 'url'
import { PNG } from 'pngjs'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const publicDir = path.join(__dirname, '..', 'public')

function fillSolid(png, hex) {
  const r = Number.parseInt(hex.slice(1, 3), 16)
  const g = Number.parseInt(hex.slice(3, 5), 16)
  const b = Number.parseInt(hex.slice(5, 7), 16)
  for (let y = 0; y < png.height; y += 1) {
    for (let x = 0; x < png.width; x += 1) {
      const idx = (png.width * y + x) << 2
      png.data[idx] = r
      png.data[idx + 1] = g
      png.data[idx + 2] = b
      png.data[idx + 3] = 255
    }
  }
}

function writePng(name, w, h, hex) {
  const png = new PNG({ width: w, height: h })
  fillSolid(png, hex)
  const buf = PNG.sync.write(png)
  fs.mkdirSync(publicDir, { recursive: true })
  fs.writeFileSync(path.join(publicDir, name), buf)
}

writePng('pwa-192.png', 192, 192, '#1a5c4a')
writePng('pwa-512.png', 512, 512, '#1a5c4a')
writePng('pwa-512-maskable.png', 512, 512, '#0f3d32')

const logoPath = path.join(publicDir, 'logo.png')
if (!fs.existsSync(logoPath)) {
  fs.copyFileSync(path.join(publicDir, 'pwa-512.png'), logoPath)
}

console.log('[generate-pwa-icons] wrote pwa-192.png, pwa-512.png, pwa-512-maskable.png, logo.png (if missing)')
