import fs from 'node:fs';
import fsPromises from 'node:fs/promises';
import path from 'node:path';
import mammoth from 'mammoth';

/**
 * pdf-parse 包根 index.js 在 ESM 动态 import 下会误判为调试模式并读取 cwd 下幽灵路径；
 * 仅加载 lib 入口，且只传入 Buffer。
 */
const PDF_PARSE_LIB = 'pdf-parse/lib/pdf-parse.js';

export async function parsePdfBuffer(buffer) {
  if (!Buffer.isBuffer(buffer)) {
    throw new Error('PDF 解析需要 Buffer');
  }
  const { default: pdfParse } = await import(PDF_PARSE_LIB);
  return pdfParse(buffer);
}

/**
 * @param {Buffer} buffer
 * @param {string} ext 如 .pdf / .docx
 */
export async function extractPaperTextFromBuffer(buffer, ext) {
  const e = String(ext || '').toLowerCase();
  if (!Buffer.isBuffer(buffer) || buffer.length === 0) {
    throw new Error('文件为空或不可读');
  }
  if (e === '.pdf') {
    const data = await parsePdfBuffer(buffer);
    return String(data.text || '').replace(/\r\n/g, '\n');
  }
  if (e === '.docx') {
    const { value } = await mammoth.extractRawText({ buffer });
    return String(value || '').replace(/\r\n/g, '\n');
  }
  if (e === '.doc') {
    throw new Error('暂不支持旧版 Word（.doc），请在 Word 中另存为 .docx 后再上传');
  }
  throw new Error('不支持的文件类型');
}

/**
 * 按磁盘路径提取正文：PDF 读入 Buffer；Word 使用 mammoth.extractRawText({ path })（OOXML .docx）。
 */
export async function extractPaperTextFromPath(absPath, ext) {
  const normalized = path.resolve(path.normalize(String(absPath || '')));
  if (!normalized || !fs.existsSync(normalized)) {
    throw new Error('文件不存在或路径无效');
  }
  const e = String(ext || '').toLowerCase();
  if (e === '.pdf') {
    const buf = await fsPromises.readFile(normalized);
    return extractPaperTextFromBuffer(buf, '.pdf');
  }
  if (e === '.docx') {
    const { value } = await mammoth.extractRawText({ path: normalized });
    return String(value || '').replace(/\r\n/g, '\n');
  }
  if (e === '.doc') {
    throw new Error('暂不支持旧版 Word（.doc），请在 Word 中另存为 .docx 后再上传');
  }
  throw new Error('不支持的文件类型');
}
