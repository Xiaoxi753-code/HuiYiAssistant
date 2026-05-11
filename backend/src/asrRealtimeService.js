import fs from 'node:fs';
import WebSocket from 'ws';
import ffmpeg from 'fluent-ffmpeg';
import ffmpegInstaller from '@ffmpeg-installer/ffmpeg';
import logger from './utils/logger.js';

ffmpeg.setFfmpegPath(ffmpegInstaller.path);

const CHUNK_BYTES = 3200;
/** 约对应 16kHz / 16bit / mono 下 ~100ms 数据量，用于节流避免瞬时打满 WS 缓冲 */
const CHUNK_PACE_MS = 100;
const DEFAULT_REALTIME_WS =
  'wss://dashscope.aliyuncs.com/api-ws/v1/realtime?model=qwen3-asr-flash-realtime-2026-02-10';

/**
 * 仅允许 DashScope / 阿里云 wss，降低 SSRF 风险
 * @param {string} urlStr
 */
export function assertSafeDashscopeRealtimeUrl(urlStr) {
  const s = String(urlStr || '').trim();
  if (!s || s.length > 2048) {
    throw new Error('无效的 WebSocket URL');
  }
  let u;
  try {
    u = new URL(s);
  } catch {
    throw new Error('无效的 WebSocket URL');
  }
  if (u.protocol !== 'wss:') {
    throw new Error('仅支持 wss 协议');
  }
  const host = u.hostname.toLowerCase();
  if (!host.endsWith('.aliyuncs.com') && host !== 'aliyuncs.com') {
    throw new Error('WebSocket 主机名须为 *.aliyuncs.com（DashScope 实时链路）');
  }
  return s;
}

export { DEFAULT_REALTIME_WS };

function flushSse(res) {
  if (typeof res.flush === 'function') {
    try {
      res.flush();
    } catch {
      /* compression 未启用时无 flush */
    }
  }
}

function sseWrite(res, obj) {
  res.write(`data: ${JSON.stringify(obj)}\n\n`);
  flushSse(res);
}

/**
 * 将上传文件（webm/mp3 等）经 FFmpeg 实时解码为 16kHz 单声道 s16le 裸 PCM 可读流。
 * 在 pipe 之前注册 `error`，避免 spawn 失败时未捕获异常拖垮进程。
 * @param {string} inputPath
 * @param {(err: Error) => void} [onError]
 * @returns {{ command: { kill: (signal?: string) => void }; stream: import('stream').Readable }}
 */
function createFfmpegPcm16Mono16kStream(inputPath, onError) {
  const command = ffmpeg(inputPath)
    .toFormat('s16le')
    .audioChannels(1)
    .audioFrequency(16000)
    .audioCodec('pcm_s16le')
    .noVideo()
    .on('error', (err) => {
      logger.error('[asr-realtime] FFmpeg 处理发生错误', err?.message || err);
      if (typeof onError === 'function') {
        try {
          onError(err);
        } catch (e) {
          logger.error('[asr-realtime] FFmpeg onError 回调异常', e?.message || e);
        }
      }
    });

  let stream;
  try {
    stream = command.pipe({ end: true });
  } catch (err) {
    logger.error('[asr-realtime] FFmpeg pipe/spawn 同步失败', err?.message || err);
    try {
      command.kill('SIGKILL');
    } catch {
      /* ignore */
    }
    if (typeof onError === 'function') {
      try {
        onError(err);
      } catch (e) {
        logger.error('[asr-realtime] FFmpeg onError 回调异常', e?.message || e);
      }
    }
    return { command, stream: null };
  }

  stream.on('error', (err) => {
    logger.error('[asr-realtime] FFmpeg PCM 可读流错误', err?.message || err);
    if (typeof onError === 'function') {
      try {
        onError(err);
      } catch (e) {
        logger.error('[asr-realtime] FFmpeg PCM onError 回调异常', e?.message || e);
      }
    }
  });

  return { command, stream };
}

/**
 * 从 Realtime / 兼容事件中抽取文本增量或整段
 * @returns {{ kind: 'append' | 'replace', text: string } | null}
 */
export function extractTranscriptPatch(parsed) {
  if (!parsed || typeof parsed !== 'object') {
    return null;
  }
  const t = String(parsed.type || '');
  if (t === 'response.audio_transcript.delta' && parsed.delta != null) {
    return { kind: 'append', text: String(parsed.delta) };
  }
  if (t === 'response.output_text.delta' && parsed.delta != null) {
    return { kind: 'append', text: String(parsed.delta) };
  }
  if (t === 'conversation.item.input_audio_transcription.delta' && parsed.delta != null) {
    return { kind: 'append', text: String(parsed.delta) };
  }
  if (t === 'response.text.delta' && parsed.delta != null) {
    return { kind: 'append', text: String(parsed.delta) };
  }
  if (parsed.delta != null && typeof parsed.delta === 'string' && t.includes('transcript')) {
    return { kind: 'append', text: parsed.delta };
  }
  if (t === 'response.audio_transcript.done' && typeof parsed.transcript === 'string') {
    return { kind: 'replace', text: parsed.transcript };
  }
  if (t === 'conversation.item.input_audio_transcription.completed' && typeof parsed.transcript === 'string') {
    return { kind: 'replace', text: parsed.transcript };
  }
  if (typeof parsed.text === 'string' && t.endsWith('.delta')) {
    return { kind: 'append', text: parsed.text };
  }
  return null;
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

/**
 * 建立 WebSocket、流式推送音频块（FFmpeg 转 PCM + 3200 字节分片 + 节流）、将识别增量以 SSE 写给客户端
 * @param {import('express').Response} res 已设置 text/event-stream
 * @param {{ absPath: string, wsUrl: string, apiKey: string, signal?: AbortSignal }} opts
 * @returns {Promise<string>} 最终完整转写文本
 */
export async function streamDashscopeRealtimeAsrToSse(res, opts) {
  const { absPath, apiKey, signal } = opts;
  let wsUrl = String(opts.wsUrl || '').trim() || DEFAULT_REALTIME_WS;
  wsUrl = assertSafeDashscopeRealtimeUrl(wsUrl);
  const key = String(apiKey || '').trim();
  if (!key) {
    throw new Error('缺少 ASR API Key（百炼）');
  }
  if (!absPath || !fs.existsSync(absPath)) {
    throw new Error('音频文件不存在');
  }

  return new Promise((resolve, reject) => {
    let accumulated = '';
    let settled = false;
    let ws = null;
    let idleTimer = null;
    /** @type {() => void} */
    let killFfmpeg = () => {};

    const connectTimer = setTimeout(() => {
      if (!settled) {
        finishErr('连接实时 ASR 服务超时');
      }
    }, 30000);

    const clearIdle = () => {
      if (idleTimer) {
        clearTimeout(idleTimer);
        idleTimer = null;
      }
    };

    const bumpIdle = () => {
      clearIdle();
      idleTimer = setTimeout(() => {
        if (!settled) {
          finishOk();
        }
      }, 180000);
    };

    const finishOk = () => {
      if (settled) {
        return;
      }
      settled = true;
      killFfmpeg();
      clearTimeout(connectTimer);
      clearIdle();
      const text = String(accumulated || '').trim();
      try {
        sseWrite(res, { type: 'done', transcript: text });
      } catch {
        /* client gone */
      }
      try {
        ws?.close();
      } catch {
        /* ignore */
      }
      resolve(text);
    };

    const finishErr = (msg) => {
      if (settled) {
        return;
      }
      settled = true;
      killFfmpeg();
      clearTimeout(connectTimer);
      clearIdle();
      const m = String(msg || '实时转写失败');
      try {
        sseWrite(res, { type: 'error', message: m });
      } catch {
        /* ignore */
      }
      try {
        ws?.close();
      } catch {
        /* ignore */
      }
      reject(new Error(m));
    };

    const onMessage = (data) => {
      bumpIdle();
      let raw = '';
      try {
        raw = Buffer.isBuffer(data) ? data.toString('utf8') : String(data);
      } catch {
        return;
      }
      let parsed;
      try {
        parsed = JSON.parse(raw);
      } catch {
        return;
      }
      try {
        if (parsed.type === 'error') {
          const em = parsed.error?.message || parsed.message || JSON.stringify(parsed);
          finishErr(em);
          return;
        }
        const patch = extractTranscriptPatch(parsed);
        if (patch && patch.text) {
          if (patch.kind === 'replace') {
            accumulated = patch.text;
            sseWrite(res, { type: 'set', text: accumulated });
          } else {
            accumulated += patch.text;
            sseWrite(res, { type: 'delta', text: patch.text });
          }
        }
        const t = String(parsed.type || '');
        if (t === 'response.done' || t === 'response.completed') {
          finishOk();
        }
      } catch (e) {
        logger.warn('[asr-realtime] message handler', e?.message || e);
      }
    };

    try {
      ws = new WebSocket(wsUrl, {
        headers: {
          Authorization: `Bearer ${key}`,
          'OpenAI-Beta': 'realtime=v1',
        },
      });
    } catch (e) {
      finishErr(e?.message || e);
      return;
    }

    ws.on('message', onMessage);

    ws.on('error', (err) => {
      logger.error('[asr-realtime] ws error', err);
      if (!settled) {
        finishErr(err?.message || 'WebSocket 错误');
      }
    });

    ws.on('close', () => {
      if (!settled) {
        finishOk();
      }
    });

    if (signal && typeof signal.addEventListener === 'function') {
      signal.addEventListener(
        'abort',
        () => {
          if (!settled) {
            finishErr('客户端已中断');
          }
        },
        { once: true },
      );
    }

    ws.on('open', () => {
      clearTimeout(connectTimer);
      bumpIdle();

      let ffmpegStopped = false;

      const { command, stream: audioStream } = createFfmpegPcm16Mono16kStream(absPath, (err) => {
        if (settled) {
          return;
        }
        killFfmpeg();
        finishErr(err?.message || '音频转码失败');
      });

      if (!audioStream) {
        return;
      }

      killFfmpeg = () => {
        if (ffmpegStopped) {
          return;
        }
        ffmpegStopped = true;
        try {
          command.kill('SIGKILL');
        } catch {
          /* ignore */
        }
        try {
          audioStream.destroy();
        } catch {
          /* ignore */
        }
      };

      void (async () => {
        try {
          const sessionUpdate = {
            type: 'session.update',
            session: {
              modalities: ['text'],
              turn_detection: { type: 'server_vad' },
              input_audio_format: 'pcm16',
              input_audio_sample_rate: 16000,
            },
          };
          ws.send(JSON.stringify(sessionUpdate));

          let carry = Buffer.alloc(0);
          let totalPcmBytes = 0;

          for await (const chunk of audioStream) {
            if (signal?.aborted) {
              killFfmpeg();
              finishErr('已取消');
              return;
            }
            if (settled) {
              killFfmpeg();
              return;
            }
            if (ws.readyState !== WebSocket.OPEN) {
              killFfmpeg();
              return;
            }

            const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
            if (buf.length === 0) {
              continue;
            }
            carry = Buffer.concat([carry, buf]);

            while (carry.length >= CHUNK_BYTES && !settled && ws.readyState === WebSocket.OPEN) {
              const packet = carry.subarray(0, CHUNK_BYTES);
              carry = carry.subarray(CHUNK_BYTES);
              ws.send(
                JSON.stringify({
                  type: 'input_audio_buffer.append',
                  audio: packet.toString('base64'),
                }),
              );
              totalPcmBytes += CHUNK_BYTES;
              bumpIdle();
              await sleep(CHUNK_PACE_MS);
            }
          }

          if (settled) {
            return;
          }
          if (ws.readyState !== WebSocket.OPEN) {
            return;
          }

          if (carry.length > 0) {
            ws.send(
              JSON.stringify({
                type: 'input_audio_buffer.append',
                audio: carry.toString('base64'),
              }),
            );
            totalPcmBytes += carry.length;
            carry = Buffer.alloc(0);
            bumpIdle();
            await sleep(CHUNK_PACE_MS);
          }

          if (totalPcmBytes < 64) {
            finishErr('转码后音频过短或为空，请检查源文件是否有效');
            return;
          }

          ws.send(JSON.stringify({ type: 'input_audio_buffer.commit' }));
          bumpIdle();

          if (String(process.env.REALTIME_ASR_RESPONSE_CREATE || '1').trim() !== '0') {
            try {
              ws.send(
                JSON.stringify({
                  type: 'response.create',
                  response: { modalities: ['text'] },
                }),
              );
            } catch (e) {
              logger.warn('[asr-realtime] response.create skipped', e?.message || e);
            }
          }
        } catch (e) {
          killFfmpeg();
          if (!settled) {
            finishErr(e?.message || e);
          }
        }
      })();
    });
  });
}
