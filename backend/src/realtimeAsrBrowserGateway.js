import { PassThrough } from 'node:stream';
import { WebSocketServer } from 'ws';
import WebSocket from 'ws';
import ffmpeg from 'fluent-ffmpeg';
import logger from './utils/logger.js';
import {
  DEFAULT_REALTIME_WS,
  assertSafeDashscopeRealtimeUrl,
  extractTranscriptPatch,
} from './asrRealtimeService.js';

const CHUNK_BYTES = 3200;
const CHUNK_PACE_MS = 100;

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

/**
 * 将浏览器 WebM 块顺序写入 Writable；高水位 + drain 避免反压丢失。
 * @param {import('stream').Writable} stream
 * @param {Buffer} buf
 */
function writeWebmChunk(stream, buf) {
  return new Promise((resolve, reject) => {
    if (stream.destroyed) {
      reject(new Error('WebM 输入流已关闭'));
      return;
    }
    const onError = (err) => {
      stream.off('drain', onDrain);
      reject(err);
    };
    const onDrain = () => {
      stream.off('error', onError);
      resolve();
    };
    stream.once('error', onError);
    try {
      const ok = stream.write(buf);
      if (ok) {
        stream.off('error', onError);
        resolve();
      } else {
        stream.once('drain', onDrain);
      }
    } catch (e) {
      stream.off('error', onError);
      reject(e);
    }
  });
}

/**
 * 在已有 HTTP(S) server 上挂载浏览器 ⇄ DashScope 实时 ASR 代理（全双工）。
 * @param {import('http').Server} httpServer
 */
export function attachRealtimeAsrWss(httpServer) {
  const wss = new WebSocketServer({ server: httpServer, path: '/api/ws/realtime-asr' });

  wss.on('connection', (browserWs) => {
    let settled = false;
    let helloReceived = false;
    let finalized = false;
    /** @type {import('stream').PassThrough | null} */
    let webmIn = null;
    /** @type {import('stream').Readable | null} */
    let pcmOut = null;
    /** @type {{ kill: (s?: string) => void } | null} */
    let ffmpegCmd = null;
    /** @type {WebSocket | null} */
    let upstream = null;
    let accumulated = '';
    /** @type {Promise<void>} */
    let writeQueue = Promise.resolve();

    const safeSendBrowser = (obj) => {
      if (browserWs.readyState !== WebSocket.OPEN) {
        return;
      }
      try {
        browserWs.send(JSON.stringify(obj));
      } catch {
        /* ignore */
      }
    };

    const cleanup = () => {
      if (settled) {
        return;
      }
      settled = true;
      try {
        ffmpegCmd?.kill?.('SIGKILL');
      } catch {
        /* ignore */
      }
      ffmpegCmd = null;
      try {
        webmIn?.destroy();
      } catch {
        /* ignore */
      }
      webmIn = null;
      try {
        pcmOut?.destroy();
      } catch {
        /* ignore */
      }
      pcmOut = null;
      try {
        upstream?.close();
      } catch {
        /* ignore */
      }
      upstream = null;
    };

    const failBrowser = (message) => {
      safeSendBrowser({ type: 'error', message: String(message || '实时转写失败') });
      try {
        browserWs.close();
      } catch {
        /* ignore */
      }
      cleanup();
    };

    browserWs.on('close', cleanup);
    browserWs.on('error', () => {
      cleanup();
    });

    /**
     * 顺序消费 FFmpeg 输出的 PCM：满 3200 字节再 pacing，结束时 commit。
     */
    async function pumpPcmToUpstream() {
      if (!pcmOut || !upstream) {
        return;
      }
      let carry = Buffer.alloc(0);
      try {
        for await (const chunk of pcmOut) {
          if (settled || upstream.readyState !== WebSocket.OPEN) {
            break;
          }
          const buf = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
          if (buf.length === 0) {
            continue;
          }
          carry = Buffer.concat([carry, buf]);
          while (carry.length >= CHUNK_BYTES && upstream.readyState === WebSocket.OPEN && !settled) {
            const packet = carry.subarray(0, CHUNK_BYTES);
            carry = carry.subarray(CHUNK_BYTES);
            upstream.send(
              JSON.stringify({
                type: 'input_audio_buffer.append',
                audio: packet.toString('base64'),
              }),
            );
            await sleep(CHUNK_PACE_MS);
          }
        }

        if (settled || upstream.readyState !== WebSocket.OPEN) {
          return;
        }

        if (carry.length > 0) {
          upstream.send(
            JSON.stringify({
              type: 'input_audio_buffer.append',
              audio: carry.toString('base64'),
            }),
          );
          carry = Buffer.alloc(0);
          await sleep(CHUNK_PACE_MS);
        }

        upstream.send(JSON.stringify({ type: 'input_audio_buffer.commit' }));

        if (String(process.env.REALTIME_ASR_RESPONSE_CREATE || '1').trim() !== '0') {
          try {
            upstream.send(
              JSON.stringify({
                type: 'response.create',
                response: { modalities: ['text'] },
              }),
            );
          } catch (e) {
            logger.warn('[realtime-asr-ws] response.create skipped', e?.message || e);
          }
        }
      } catch (e) {
        if (!settled) {
          logger.error('[realtime-asr-ws] pcm pump', e);
          failBrowser(e?.message || 'PCM 转发失败');
        }
      }
    }

    const onUpstreamMessage = (data) => {
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
      if (parsed.type === 'error') {
        const em = parsed.error?.message || parsed.message || JSON.stringify(parsed);
        failBrowser(em);
        return;
      }
      const patch = extractTranscriptPatch(parsed);
      if (patch?.text) {
        if (patch.kind === 'replace') {
          accumulated = patch.text;
          safeSendBrowser({ type: 'transcript', op: 'set', text: patch.text });
        } else {
          accumulated += patch.text;
          safeSendBrowser({ type: 'transcript', op: 'append', text: patch.text });
        }
      }
      const t = String(parsed.type || '');
      if (t === 'response.done' || t === 'response.completed') {
        safeSendBrowser({ type: 'done', transcript: String(accumulated || '').trim() });
        cleanup();
        try {
          browserWs.close();
        } catch {
          /* ignore */
        }
      }
    };

    browserWs.on('message', (data, isBinary) => {
      if (settled) {
        return;
      }
      if (!isBinary) {
        let msg;
        try {
          msg = JSON.parse(String(data));
        } catch {
          failBrowser('无效的 JSON 控制帧');
          return;
        }
        const t = String(msg.type || '');
        if (t === 'hello') {
          if (helloReceived) {
            return;
          }
          helloReceived = true;
          let wsUrl = String(msg.asrWsUrl || msg.realtimeAsrWsUrl || '').trim() || DEFAULT_REALTIME_WS;
          const apiKey = String(msg.asrApiKey || msg.realtimeAsrApiKey || '').trim();
          if (!apiKey) {
            failBrowser('缺少 asrApiKey（百炼）');
            return;
          }
          try {
            wsUrl = assertSafeDashscopeRealtimeUrl(wsUrl);
          } catch (e) {
            failBrowser(e?.message || 'WebSocket URL 无效');
            return;
          }

          webmIn = new PassThrough({ highWaterMark: 16 * 1024 * 1024 });
          const cmd = ffmpeg(webmIn)
            .inputOptions(['-f', 'webm', '-fflags', '+genpts+discardcorrupt', '-thread_queue_size', '512'])
            .noVideo()
            .audioCodec('pcm_s16le')
            .audioChannels(1)
            .audioFrequency(16000)
            .format('s16le')
            .on('error', (err) => {
              if (!settled) {
                logger.error('[realtime-asr-ws] FFmpeg 处理发生错误', err?.message || err);
                failBrowser(err?.message || 'FFmpeg 转码失败');
              }
            });

          try {
            pcmOut = cmd.pipe({ end: true });
          } catch (err) {
            logger.error('[realtime-asr-ws] FFmpeg pipe/spawn 同步失败', err?.message || err);
            failBrowser(err?.message || 'FFmpeg 启动失败');
            return;
          }
          ffmpegCmd = cmd;

          pcmOut.on('error', (err) => {
            if (!settled) {
              logger.error('[realtime-asr-ws] pcm stream', err);
              failBrowser(err?.message || 'PCM 流错误');
            }
          });

          try {
            upstream = new WebSocket(wsUrl, {
              headers: {
                Authorization: `Bearer ${apiKey}`,
                'OpenAI-Beta': 'realtime=v1',
              },
            });
          } catch (e) {
            failBrowser(e?.message || '无法连接 DashScope');
            return;
          }

          upstream.on('message', onUpstreamMessage);
          upstream.on('error', (err) => {
            if (!settled) {
              logger.error('[realtime-asr-ws] upstream ws', err);
              failBrowser(err?.message || 'DashScope WebSocket 错误');
            }
          });
          upstream.on('close', () => {
            if (settled) {
              return;
            }
            safeSendBrowser({ type: 'done', transcript: String(accumulated || '').trim() });
            cleanup();
            try {
              browserWs.close();
            } catch {
              /* ignore */
            }
          });

          upstream.on('open', () => {
            if (settled) {
              return;
            }
            try {
              upstream.send(
                JSON.stringify({
                  type: 'session.update',
                  session: {
                    modalities: ['text'],
                    turn_detection: { type: 'server_vad' },
                    input_audio_format: 'pcm16',
                    input_audio_sample_rate: 16000,
                  },
                }),
              );
            } catch (e) {
              failBrowser(e?.message || 'session.update 发送失败');
              return;
            }
            safeSendBrowser({ type: 'ready' });
            void pumpPcmToUpstream();
          });

          return;
        }

        if (t === 'finalize') {
          if (!helloReceived || !webmIn || finalized) {
            return;
          }
          finalized = true;
          writeQueue = writeQueue
            .then(async () => {
              try {
                webmIn.end();
              } catch (e) {
                logger.warn('[realtime-asr-ws] webmIn.end', e?.message || e);
              }
            })
            .catch((e) => {
              if (!settled) {
                failBrowser(e?.message || e);
              }
            });
          return;
        }

        failBrowser(`未知控制类型: ${t}`);
        return;
      }

      if (!helloReceived || !webmIn || settled) {
        return;
      }
      const buf = Buffer.isBuffer(data) ? data : Buffer.from(data);
      if (buf.length === 0) {
        return;
      }
      writeQueue = writeQueue
        .then(() => writeWebmChunk(webmIn, buf))
        .catch((e) => {
          if (!settled) {
            failBrowser(e?.message || '写入 WebM 失败');
          }
        });
    });
  });

  logger.info('[realtime-asr-ws] WebSocket 网关已挂载 path=/api/ws/realtime-asr');
}
