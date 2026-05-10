import path from 'path';
import fs from 'fs';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { extractPaperTextFromBuffer } from './src/paperExtract.js';

/** 单次对话随请求传入的上下文（由前端提供，用于 parse / Word 等） */
export const MAX_AGENT_DOCUMENT_BYTES = 10 * 1024 * 1024;

/**
 * HTTP 兼容 tools 定义（JSON Schema）
 * @returns {object[]}
 */
export function getOpenClawChatTools() {
  return [
    {
      type: 'function',
      function: {
        name: 'request_clarification',
        description:
          '向用户澄清模糊意图。最多只能包含 3 个核心问题。只能在任务开始前调用一次。',
        parameters: {
          type: 'object',
          properties: {
            missing_info: {
              type: 'array',
              items: { type: 'string' },
              maxItems: 3,
              description:
                '最多 3 条短问题，每条一句；只收录最影响交付的核心缺口（如受众、篇幅、截止时间）。',
            },
          },
          required: ['missing_info'],
        },
      },
    },
    {
      type: 'function',
      function: {
        name: 'parse_document',
        description:
          '获取文献正文：若本轮请求已在 agentContext 中附带「已解析全文」parsedPaperText，则直接返回该文本摘要，无需重新上传；否则使用 documentBase64 解析 PDF/Word。仅在用户明确要求提取/解析正文时调用。',
        parameters: {
          type: 'object',
          properties: {
            reason: { type: 'string', description: '简短说明调用原因（内部审计用）' },
          },
          required: [],
        },
      },
    },
    {
      type: 'function',
      function: {
        name: 'generate_word_report',
        description:
          '根据已上传的 Word 模板与占位变量生成填充后的 .docx（Base64）。需已知模板 ID 与变量键值；不得编造模板 ID。',
        parameters: {
          type: 'object',
          properties: {
            template_id: { type: 'integer', description: '模板在系统中的数字 ID' },
            variables: {
              type: 'object',
              description: '占位符名 → 替换文本（与模板中 {变量名} 一致）',
              additionalProperties: { type: 'string' },
            },
          },
          required: ['template_id', 'variables'],
        },
      },
    },
    {
      type: 'function',
      function: {
        name: 'switch_theme',
        description:
          '建议切换工作台界面模式或跳转 Tab（仅返回导航意图，由前端执行）。例如用户想去情感陪伴或文献页。',
        parameters: {
          type: 'object',
          properties: {
            target: {
              type: 'string',
              enum: ['default', 'companion', 'paper', 'templates', 'apiConfig', 'chat'],
              description: '目标：default=主工作台对话；companion=情感陪伴；paper=文献摘要；templates=模板引擎；apiConfig=模型配置；chat=对话助手',
            },
          },
          required: ['target'],
        },
      },
    },
    {
      type: 'function',
      function: {
        name: 'update_model_config',
        description:
          '在已存在的模型配置上调整启用状态或优先级。禁止通过此工具写入 API Key；密钥须由用户在配置页手动填写。',
        parameters: {
          type: 'object',
          properties: {
            action: {
              type: 'string',
              enum: ['set_enabled', 'set_priority'],
              description: 'set_enabled：启用/禁用；set_priority：仅改优先级数字',
            },
            id: { type: 'integer', description: 'model_configs 表主键 id' },
            is_enabled: { type: 'boolean', description: 'action=set_enabled 时必填' },
            priority: { type: 'integer', description: 'action=set_priority 时必填' },
          },
          required: ['action', 'id'],
        },
      },
    },
  ];
}

function buildFilledDocxBase64(buffer, values) {
  const zip = new PizZip(buffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: { start: '{', end: '}' },
    nullGetter: (part) => `{${part.value}}`,
  });
  doc.render(values || {});
  const output = doc.getZip().generate({ type: 'nodebuffer' });
  return output.toString('base64');
}

/**
 * @param {{ db: import('node:sqlite').DatabaseSync, uploadsDir: string, tempDir: string, agentContext?: object }} deps
 * @returns {(name: string, rawArgs: string|object) => Promise<string>}
 */
export function createToolExecutor(deps) {
  const { db, uploadsDir, agentContext } = deps;

  return async function executeOpenClawTool(name, rawArgs) {
    let args = {};
    try {
      args = typeof rawArgs === 'string' && rawArgs.trim() ? JSON.parse(rawArgs) : rawArgs && typeof rawArgs === 'object' ? rawArgs : {};
    } catch {
      return JSON.stringify({ ok: false, error: '工具参数 JSON 无效' });
    }

    try {
      if (name === 'request_clarification') {
        const arr = Array.isArray(args.missing_info)
          ? args.missing_info.map((x) => String(x).trim()).filter(Boolean).slice(0, 3)
          : [];
        return JSON.stringify({
          ok: true,
          pending_ui_clarification: true,
          missing_info: arr,
          message: '已记录澄清请求；若未走 Agent 流拦截，请用户在前端澄清卡片中填写并提交。',
        });
      }

      if (name === 'parse_document') {
        const cached = String(agentContext?.parsedPaperText || agentContext?.paperFullText || '').trim();
        if (cached.length > 0) {
          const fname = String(agentContext?.paperFileName || agentContext?.documentName || '已解析文献').trim() || '已解析文献';
          const text = cached;
          const maxOut = 120_000;
          const truncated = text.length > maxOut;
          const out = truncated ? text.slice(0, maxOut) : text;
          const modelPreviewLen = 12_000;
          const modelText =
            out.length > modelPreviewLen ? `${out.slice(0, modelPreviewLen)}\n…（截断展示，完整字数见 charCount）` : out;
          return JSON.stringify({
            ok: true,
            filename: fname,
            charCount: text.length,
            truncated,
            text: modelText,
            source: 'parsed_session',
            navigate_tab: 'paper',
            ui_hint: '文献模块',
          });
        }

        const b64 = String(agentContext?.documentBase64 || '').trim();
        const fname = String(agentContext?.documentName || 'document.bin').trim() || 'document.bin';
        if (!b64) {
          return JSON.stringify({
            ok: false,
            error:
              '当前请求未附带已解析正文或文件数据。请用户在「文献论文摘要」页先解析文档，或使用「指挥 OpenClaw」注入上下文后再试。',
          });
        }
        let buf;
        try {
          buf = Buffer.from(b64, 'base64');
        } catch {
          return JSON.stringify({ ok: false, error: '文件 Base64 无效' });
        }
        if (!buf.length || buf.length > MAX_AGENT_DOCUMENT_BYTES) {
          return JSON.stringify({ ok: false, error: '文件为空或超过大小上限（10MB）' });
        }
        const ext = path.extname(fname).toLowerCase();
        if (ext !== '.pdf' && ext !== '.docx') {
          return JSON.stringify({ ok: false, error: '仅支持 .pdf 或 .docx' });
        }
        const text = await extractPaperTextFromBuffer(buf, ext);
        const maxOut = 120_000;
        const truncated = text.length > maxOut;
        const out = truncated ? text.slice(0, maxOut) : text;
        const modelPreviewLen = 12_000;
        const modelText = out.length > modelPreviewLen ? `${out.slice(0, modelPreviewLen)}\n…（截断展示，完整字数见 charCount）` : out;
        return JSON.stringify({
          ok: true,
          filename: fname,
          charCount: text.length,
          truncated,
          text: modelText,
          navigate_tab: 'paper',
          ui_hint: '文献模块',
        });
      }

      if (name === 'generate_word_report') {
        const templateId = Number.parseInt(String(args.template_id), 10);
        if (!Number.isFinite(templateId) || templateId < 1) {
          return JSON.stringify({ ok: false, error: '无效的 template_id' });
        }
        const variables = args.variables && typeof args.variables === 'object' && !Array.isArray(args.variables) ? args.variables : {};
        const stmt = db.prepare('SELECT * FROM templates WHERE id = ?');
        const template = stmt.get(templateId);
        if (!template) {
          return JSON.stringify({ ok: false, error: '模板不存在' });
        }
        const filePath = path.join(uploadsDir, template.filename);
        if (!fs.existsSync(filePath)) {
          return JSON.stringify({ ok: false, error: '模板文件在服务器上不存在' });
        }
        const buffer = await fs.promises.readFile(filePath);
        buildFilledDocxBase64(buffer, variables);
        const outName = `${String(template.original_name || 'report').replace(/\.docx$/i, '')}-filled.docx`;
        return JSON.stringify({
          ok: true,
          template_id: templateId,
          download_filename: outName,
          message:
            '已在服务端用给定变量完成 Word 填充校验（模板与变量匹配）。请在「智能模板引擎」Tab 使用「导出 Word」或预览；为控制上下文体积，此处不返回整文件 Base64。',
          navigate_tab: 'templates',
          ui_hint: 'Word 引擎',
        });
      }

      if (name === 'switch_theme') {
        const target = String(args.target || '').trim();
        const allowed = new Set(['default', 'companion', 'paper', 'templates', 'apiConfig', 'chat']);
        if (!allowed.has(target)) {
          return JSON.stringify({ ok: false, error: '无效 target' });
        }
        return JSON.stringify({
          ok: true,
          navigate_tab: target === 'default' ? 'chat' : target,
          theme_mode: target === 'companion' ? 'companion' : 'default',
          ui_hint: target === 'companion' ? '情感陪伴' : '工作台',
        });
      }

      if (name === 'update_model_config') {
        const action = String(args.action || '');
        const id = Number.parseInt(String(args.id), 10);
        if (!Number.isFinite(id) || id < 1) {
          return JSON.stringify({ ok: false, error: '无效的 id' });
        }
        const existing = db.prepare('SELECT * FROM model_configs WHERE id = ?').get(id);
        if (!existing) {
          return JSON.stringify({ ok: false, error: '配置不存在' });
        }
        if (action === 'set_enabled') {
          if (typeof args.is_enabled !== 'boolean') {
            return JSON.stringify({ ok: false, error: 'set_enabled 需要布尔字段 is_enabled' });
          }
          const enabled = args.is_enabled ? 1 : 0;
          db.prepare('UPDATE model_configs SET is_enabled = ? WHERE id = ?').run(enabled, id);
          return JSON.stringify({
            ok: true,
            message: `已${args.is_enabled ? '启用' : '禁用'}配置 id=${id}`,
            navigate_tab: 'apiConfig',
            ui_hint: 'API 配置',
          });
        }
        if (action === 'set_priority') {
          const pri = Number.parseInt(String(args.priority), 10);
          if (!Number.isFinite(pri)) {
            return JSON.stringify({ ok: false, error: 'set_priority 需要整数 priority' });
          }
          db.prepare('UPDATE model_configs SET priority = ? WHERE id = ?').run(pri, id);
          return JSON.stringify({
            ok: true,
            message: `已将配置 id=${id} 的优先级设为 ${pri}`,
            navigate_tab: 'apiConfig',
            ui_hint: 'API 配置',
          });
        }
        return JSON.stringify({ ok: false, error: '未知 action' });
      }

      return JSON.stringify({ ok: false, error: `未知工具: ${name}` });
    } catch (e) {
      return JSON.stringify({ ok: false, error: String(e?.message || e || '工具执行失败') });
    }
  };
}

/**
 * 供 system 追加的护栏文案（减少误调用）
 */
export const TOOL_GUARD_SYSTEM_ADDON = `【工具调用纪律】
你拥有调用工作台工具的权限，但必须遵守：
1) 仅在用户明确请求执行某项功能、或已确认要生成/解析/跳转时调用工具；闲聊、泛泛咨询不要调用工具。
2) 不得猜测或捏造 template_id、文件内容；缺少参数时应先向用户追问，必要时调用 request_clarification。
3) update_model_config 仅允许调整启用状态与优先级，不得尝试写入或替换 API Key。
4) parse_document：优先使用 agentContext 中的已解析全文 parsedPaperText；仅在没有缓存正文时才依赖 documentBase64。
5) request_clarification：边界不清时用 missing_info 发问（≤3 条、同一任务仅一轮）；禁止在普通正文里追问；用户补充后禁止第二轮澄清，缺口用占位符写入交付物。`;
