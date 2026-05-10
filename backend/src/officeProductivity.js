import fs from 'fs';
import path from 'path';
import {
  postChatCompletion,
  extractTextFromResponse,
  withThinkingParams,
  normalizeModelId,
  transcribeAudioWithWhisper,
} from './services/openclawService.js';

/**
 * @param {import('./services/openclawService.js').ModelCallConfig} modelConfig
 * @param {string} system
 * @param {string} user
 * @param {number} [temperature]
 */
export async function llmCompletePlain(modelConfig, system, user, temperature = 0.35) {
  if (!modelConfig?.api_key?.trim()) {
    throw new Error('模型配置缺少 API Key');
  }
  const model = normalizeModelId(modelConfig.model_name) || 'qwen-plus';
  const body = withThinkingParams(
    {
      model,
      messages: [
        { role: 'system', content: system },
        { role: 'user', content: user },
      ],
      temperature,
      stream: false,
    },
    modelConfig.model_name,
  );
  const data = await postChatCompletion(modelConfig, body);
  const text = extractTextFromResponse(data);
  return String(text || '').trim();
}

/**
 * 使用独立 Whisper 客户端对本地音频文件做真实转写（与对话模型 base_url 分离）。
 * @param {import('./services/openclawService.js').ModelCallConfig} modelConfig
 * @param {string} absPath 磁盘绝对路径
 * @param {string} [originalFilename]
 * @param {{ whisperApiKey?: string, whisperBaseUrl?: string }} [whisperOpts] 来自前端 multipart 字段
 */
export async function transcribeMeetingAudioFile(modelConfig, absPath, originalFilename, whisperOpts = {}) {
  const buf = fs.readFileSync(absPath);
  const name = originalFilename || path.basename(absPath) || 'recording.webm';
  return transcribeAudioWithWhisper(modelConfig, buf, name, whisperOpts);
}

/**
 * @param {import('./services/openclawService.js').ModelCallConfig} modelConfig
 * @param {string} transcript
 */
export async function generateMeetingMinutesMarkdown(modelConfig, transcript) {
  const system = `你是一个金牌会议秘书。请将用户给出的会议录音转写文本，严格按照以下三个模块进行结构化输出（使用 Markdown，二级标题必须一字不差）：

## 议题背景
## 核心结论
## 待办事项（负责人+时间）

要求：
- 「议题背景」用 2–5 条要点概括上下文与讨论范围。
- 「核心结论」用条目列出已对齐的结论；不得编造逐字稿中未出现的硬事实；不确定处标注「待确认」。
- 「待办事项（负责人+时间）」每条使用 Markdown 任务列表「- [ ] 」，同一行内写清责任人与时间点（若逐字稿中缺失则写「待指派｜时间待确认」）。
- 不要输出与以上三模块无关的大段开场白。`;
  const user = `【会议转写文本】\n${String(transcript || '').trim().slice(0, 24000)}`;
  return llmCompletePlain(modelConfig, system, user, 0.25);
}

function stripJsonFence(raw) {
  let s = String(raw || '').trim();
  const fence = s.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fence) {
    s = fence[1].trim();
  }
  return s;
}

/**
 * @param {import('./services/openclawService.js').ModelCallConfig} modelConfig
 * @param {string} combinedText
 * @param {string} nowIso 当前时间 ISO 字符串（注入模型推理相对日期）
 */
export async function extractTodosJson(modelConfig, combinedText, nowIso) {
  const system = `你是待办提取器。从对话或纪要中抽取「含时间信息或可转为截止时间的行动项」。
仅输出 **JSON 数组**，元素形状：
[{"title":"string","remind_at":"ISO8601","notes":"string"}]
规则：
- remind_at 尽量具体到小时；若只有「明天下午三点」等中文，请结合用户给出的「当前时间」换算为中国本地时间的 ISO8601（不加 Z，用 +08:00 或本地 naive 均可，保持一致）。
- 若无时间线索但有明确交付项，可将 remind_at 设为当日 18:00 或下一工作日 10:00，并在 notes 标明「时间系推断」。
- 最多 12 条；若无合格项输出 []。
- 不要输出任何 markdown、解释或前后缀。`;
  const user = `当前时间参考：${nowIso}

【待分析文本】
${String(combinedText || '').trim().slice(0, 12000)}`;
  const raw = await llmCompletePlain(modelConfig, system, user, 0.1);
  const inner = stripJsonFence(raw);
  try {
    const arr = JSON.parse(inner);
    return Array.isArray(arr) ? arr : [];
  } catch {
    const brace = inner.match(/\[[\s\S]*\]/);
    if (brace) {
      try {
        const arr = JSON.parse(brace[0]);
        return Array.isArray(arr) ? arr : [];
      } catch {
        return [];
      }
    }
    return [];
  }
}

/**
 * @param {import('node:sqlite').DatabaseSync} db
 * @param {string} userEmail
 * @param {{ title?: string, remind_at?: string, notes?: string }[]} items
 * @param {string} source
 */
export function insertTodoRows(db, userEmail, items, source) {
  const ue = String(userEmail || '').trim();
  if (!ue || !Array.isArray(items)) {
    return 0;
  }
  const stmt = db.prepare(
    `INSERT INTO todos (user_email, title, remind_at, notes, source, done)
     VALUES (?, ?, ?, ?, ?, 0)`,
  );
  let n = 0;
  for (const it of items) {
    const title = String(it?.title || '').trim().slice(0, 500);
    let remindAt = String(it?.remind_at || '').trim().slice(0, 40);
    const notes = String(it?.notes || '').trim().slice(0, 2000);
    if (!title) {
      continue;
    }
    if (!remindAt) {
      remindAt = new Date().toISOString().slice(0, 16);
    }
    try {
      stmt.run(ue, title, remindAt, notes, source);
      n += 1;
    } catch (e) {
      console.warn('[todos insert]', e?.message || e);
    }
  }
  return n;
}

/**
 * @param {import('./services/openclawService.js').ModelCallConfig} modelConfig
 */
export async function generatePptOutlineMarkdown(modelConfig, topic, pages, style) {
  const system = `你是演示文稿策划。输出 **Markdown 幻灯片提纲**，便于导入生成 PPTX。
格式要求：
- 第一行：# 总标题（简短）
- 随后每一页用二级标题 ## 第 n 页：短标题
- 每页下面用 - 开头的 2–5 条要点，不要长篇段落。
- 共约 ${pages} 页（含封面信息页）。风格：${style}。`;
  const user = `主题：${String(topic || '').trim().slice(0, 500)}`;
  return llmCompletePlain(modelConfig, system, user, 0.45);
}

/**
 * 清理会议上传临时文件（失败不影响响应）
 * @param {string} absPath
 */
export function safeUnlink(absPath) {
  try {
    if (absPath && fs.existsSync(absPath)) {
      fs.unlinkSync(absPath);
    }
  } catch {
    /* ignore */
  }
}

export function getMulterMeetingAbsPath(file) {
  if (!file) {
    return '';
  }
  if (file.path && path.isAbsolute(file.path)) {
    return file.path;
  }
  const dest = file.destination || '';
  const fn = file.filename || '';
  if (dest && fn) {
    return path.join(dest, fn);
  }
  return '';
}
