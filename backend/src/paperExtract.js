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
  throw new Error('不支持的文件类型');
}
