import {
  postChatCompletion,
  buildOpenClawMergedSystemContent,
  resolveChatTemperature,
  withThinkingParams,
  normalizeModelId,
  parseSsePayloads,
  sendChatStreamWithBody,
} from './services/openclawService.js';
import { getOpenClawChatTools, TOOL_GUARD_SYSTEM_ADDON } from '../tools.js';
import logger from './utils/logger.js';

const MAX_TOOL_ROUNDS = 4;

const TOOL_STATUS_LABELS = {
  request_clarification: '正在收集关键约束…',
  parse_document: '正在指挥 [智能长文速读] 提取正文…',
  generate_word_report: '正在调用 [Word 引擎] 生成文档…',
  switch_theme: '正在切换界面 / 主题…',
  update_model_config: '正在更新 [API 模型配置]…',
};

function writeSseJson(res, flushSse, obj) {
  try {
    if (res.writableEnded) {
      return;
    }
    res.write(`data: ${JSON.stringify(obj)}\n\n`);
    flushSse();
  } catch (e) {
    console.warn('[agent] SSE write failed (client reset or closed):', e?.code || e?.message || e);
  }
}

function assistantMessageText(msg) {
  if (!msg || msg.content == null) {
    return '';
  }
  if (typeof msg.content === 'string') {
    return msg.content;
  }
  if (Array.isArray(msg.content)) {
    return msg.content
      .map((p) => (typeof p?.text === 'string' ? p.text : ''))
      .filter(Boolean)
      .join('\n');
  }
  return '';
}

/**
 * 规范化各厂商返回的 tool_calls 元素，避免访问 undefined 触发 TypeError。
 * @param {unknown[]} toolCalls
 */
function normalizeToolCalls(toolCalls) {
  if (!Array.isArray(toolCalls)) {
    return [];
  }
  return toolCalls
    .filter((tc) => tc != null && typeof tc === 'object')
    .map((tc) => {
      const fn = tc.function && typeof tc.function === 'object' ? tc.function : {};
      const name = String(fn.name ?? tc.name ?? '').trim();
      const args = String(fn.arguments ?? tc.arguments ?? '{}');
      return {
        id: String(tc.id ?? ''),
        type: 'function',
        function: { name, arguments: args },
      };
    })
    .filter((tc) => tc.function.name.length > 0);
}

/**
 * Agent：非流式多轮 tool_calls → 后端执行 → 最后流式输出自然语言（SSE 不断连）。
 * @param {import('express').Response} res
 * @param {{ userMessage: string, modelConfig: object, streamOptions: object, executeTool: (name:string, args:string)=>Promise<string>, flushSse: ()=>void }} deps
 */
export async function runAgentChatToSse(res, deps) {
  const { userMessage, modelConfig, streamOptions, executeTool, flushSse } = deps;
  const clarificationFollowUp = Boolean(streamOptions?.clarificationFollowUp);

  try {
    const model = normalizeModelId(modelConfig.model_name) || 'deepseek-v3';
    const systemContent = `${buildOpenClawMergedSystemContent(streamOptions)}\n\n${TOOL_GUARD_SYSTEM_ADDON}`;
    /** @type {object[]} */
    const messages = [
      { role: 'system', content: systemContent },
      { role: 'user', content: String(userMessage || '').trim() },
    ];

    let pendingNavTab = '';

    for (let round = 0; round < MAX_TOOL_ROUNDS; round += 1) {
      const bodyCore = {
        model,
        messages,
        temperature: resolveChatTemperature(streamOptions),
        stream: false,
      };
      const body = withThinkingParams(
        clarificationFollowUp
          ? bodyCore
          : {
              ...bodyCore,
              tools: getOpenClawChatTools(),
              tool_choice: 'auto',
            },
        modelConfig.model_name,
      );

      let data;
      try {
        data = await postChatCompletion(modelConfig, body);
      } catch (e) {
        logger.error('[agent] postChatCompletion failed', e);
        writeSseJson(res, flushSse, { type: 'delta', text: `请求模型失败：${String(e?.message || e)}` });
        writeSseJson(res, flushSse, { type: 'done' });
        return;
      }

      const choice = Array.isArray(data?.choices) ? data.choices[0] : undefined;
      const msg = choice && typeof choice.message === 'object' ? choice.message : null;

      if (!msg) {
        writeSseJson(res, flushSse, { type: 'delta', text: '模型未返回有效内容。' });
        writeSseJson(res, flushSse, { type: 'done' });
        return;
      }

      const toolCalls = normalizeToolCalls(msg.tool_calls);
      if (toolCalls.length > 0) {
        const clarifyTc = toolCalls.find((tc) => tc.function.name === 'request_clarification');
        if (clarifyTc) {
          let missing = [];
          try {
            const rawArgs = clarifyTc.function.arguments || '{}';
            const args =
              typeof rawArgs === 'string' ? JSON.parse(rawArgs || '{}') : rawArgs && typeof rawArgs === 'object' ? rawArgs : {};
            if (args && Array.isArray(args.missing_info)) {
              missing = args.missing_info.map((x) => String(x).trim()).filter(Boolean);
            }
          } catch {
            missing = [];
          }
          missing = missing.slice(0, 3);
          if (missing.length === 0) {
            missing = ['请补充：具体场景、受众或交付格式是什么？'];
          }
          writeSseJson(res, flushSse, { type: 'clarification', missing_info: missing });
          writeSseJson(res, flushSse, { type: 'done' });
          return;
        }

        const assistantRecord = {
          role: 'assistant',
          content: assistantMessageText(msg) || null,
          tool_calls: toolCalls.map((tc) => ({
            id: tc.id,
            type: 'function',
            function: {
              name: tc.function.name,
              arguments: tc.function.arguments,
            },
          })),
        };
        messages.push(assistantRecord);

        for (const tc of toolCalls) {
          const name = tc.function.name;
          const argsStr = tc.function.arguments;
          const label = TOOL_STATUS_LABELS[name] || `正在执行工具：${name}…`;
          writeSseJson(res, flushSse, { type: 'agent_status', text: label });
          let toolContent;
          try {
            toolContent = await executeTool(name, argsStr);
          } catch (toolErr) {
            logger.error('[agent] executeTool failed', name, toolErr);
            toolContent = JSON.stringify({
              ok: false,
              error: String(toolErr?.message || toolErr || '工具执行异常'),
            });
          }
          messages.push({
            role: 'tool',
            tool_call_id: tc.id || `call_${round}_${name}`,
            content: toolContent,
          });
          try {
            const parsed = JSON.parse(toolContent);
            if (parsed && typeof parsed.navigate_tab === 'string' && parsed.navigate_tab) {
              pendingNavTab = parsed.navigate_tab;
            }
          } catch {
            /* 非 JSON 工具结果忽略 */
          }
        }
        continue;
      }

      const plain = assistantMessageText(msg);
      if (plain) {
        writeSseJson(res, flushSse, { type: 'agent_status', text: '正在整理回复…' });
        messages.push({ role: 'assistant', content: plain });
        writeSseJson(res, flushSse, { type: 'delta', text: plain });
      } else {
        const streamBody = {
          model,
          messages,
          temperature: resolveChatTemperature(streamOptions),
          stream: true,
        };
        let reader;
        let response;
        try {
          const pair = await sendChatStreamWithBody(modelConfig, streamBody);
          reader = pair.reader;
          response = pair.response;
        } catch (streamErr) {
          logger.error('[agent] sendChatStreamWithBody threw', streamErr);
          writeSseJson(res, flushSse, { type: 'delta', text: `流式请求失败：${String(streamErr?.message || streamErr)}` });
          writeSseJson(res, flushSse, { type: 'done' });
          return;
        }
        if (!response.ok || !reader) {
          let errText = '';
          try {
            errText = await response.text();
          } catch {
            errText = '';
          }
          logger.error('[agent] final stream failed', response.status, errText);
          writeSseJson(res, flushSse, { type: 'delta', text: '生成回复失败，请稍后再试。' });
          writeSseJson(res, flushSse, { type: 'done' });
          return;
        }
        writeSseJson(res, flushSse, { type: 'agent_status', text: '正在生成回复…' });
        const decoder = new TextDecoder('utf-8');
        let buf = '';
        try {
          while (true) {
            const { value, done } = await reader.read();
            if (value?.byteLength) {
              buf += decoder.decode(value, { stream: !done });
            }
            const lines = buf.split('\n');
            buf = lines.pop() || '';
            for (const line of lines) {
              const trimmed = line.trim();
              if (!trimmed.startsWith('data:')) {
                continue;
              }
              const payload = trimmed.replace(/^data:\s*/i, '').trim();
              if (!payload || payload === '[DONE]') {
                continue;
              }
              let payloads = [];
              try {
                payloads = parseSsePayloads(payload);
              } catch {
                continue;
              }
              for (const json of payloads) {
                try {
                  const choices = json?.choices;
                  const first = Array.isArray(choices) ? choices[0] : undefined;
                  const delta = first && typeof first.delta === 'object' ? first.delta : null;
                  if (!delta) {
                    continue;
                  }
                  const reasoning = delta.reasoning_content;
                  if (typeof reasoning === 'string' && reasoning.length > 0) {
                    writeSseJson(res, flushSse, { type: 'thought', text: reasoning });
                  }
                  const contentPiece = delta.content;
                  if (typeof contentPiece === 'string' && contentPiece.length > 0) {
                    writeSseJson(res, flushSse, { type: 'delta', text: contentPiece });
                  }
                } catch {
                  /* 单帧解析失败不中断流 */
                }
              }
            }
            if (done) {
              break;
            }
          }
        } catch (readErr) {
          logger.error('[agent] stream read error', readErr);
          writeSseJson(res, flushSse, { type: 'delta', text: '\n（流式读取中断，请重试）' });
        }
      }

      if (pendingNavTab) {
        writeSseJson(res, flushSse, { type: 'agent_nav', tab: pendingNavTab });
      }
      writeSseJson(res, flushSse, { type: 'done' });
      return;
    }

    writeSseJson(res, flushSse, {
      type: 'delta',
      text: '工具调用轮数过多，已中止。请简化需求后重试。',
    });
    writeSseJson(res, flushSse, { type: 'done' });
  } catch (fatal) {
    logger.error('【致命错误】[agent] runAgentChatToSse:', fatal);
    if (fatal && typeof fatal === 'object' && fatal.stack) {
      logger.error(fatal.stack);
    }
    writeSseJson(res, flushSse, { type: 'delta', text: `助手执行异常：${String(fatal?.message || fatal)}` });
    writeSseJson(res, flushSse, { type: 'done' });
  }
}
