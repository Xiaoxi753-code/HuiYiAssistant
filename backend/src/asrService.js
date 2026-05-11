import { streamDashscopeRealtimeAsrToSse } from './asrRealtimeService.js';

/** 不落盘 SSE：仅消费 DashScope 转写结果，供 HTTP 文件转写等非浏览器场景复用 */
const noopSseRes = {
  flush() {},
  write() {},
};

/**
 * 将磁盘上的音频文件（mp3/wav/m4a/webm 等）经 FFmpeg→PCM 后走 DashScope 实时 ASR，返回全文转写。
 * 全程流式读文件，不把整文件读入内存。
 * @param {string} absPath 绝对路径
 * @param {{ apiKey: string, wsUrl?: string, signal?: AbortSignal }} opts
 * @returns {Promise<string>}
 */
export async function transcribeUploadedAudioFile(absPath, opts) {
  return streamDashscopeRealtimeAsrToSse(noopSseRes, {
    absPath,
    apiKey: opts.apiKey,
    wsUrl: opts.wsUrl,
    signal: opts.signal,
  });
}
