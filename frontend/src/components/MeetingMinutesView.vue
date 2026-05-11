<script setup>
import { ref, computed, onBeforeUnmount, nextTick, watch } from 'vue'
import { ElMessage, ElIcon } from 'element-plus'
import 'element-plus/es/components/icon/style/css'
import { Folder, VideoCamera, VideoPause, VideoPlay, SwitchButton } from '@element-plus/icons-vue'
import { marked } from 'marked'
import DOMPurify from 'dompurify'
import { useUserStore } from '../store/user.js'

marked.setOptions({
    gfm: true,
    breaks: true,
})

const emit = defineEmits(['processed'])

const userStore = useUserStore()

const MEETING_PROCESS_URL = '/api/meeting/process'
const FILE_TRANSCRIBE_URL = '/api/meeting/file-transcribe'
const WHISPER_ASR_STORAGE_KEY = 'openclaw_whisper_asr_v1'
const DEFAULT_REALTIME_ASR_WS =
    'wss://dashscope.aliyuncs.com/api-ws/v1/realtime?model=qwen3-asr-flash-realtime-2026-02-10'
const FETCH_TIMEOUT_MS = 900000

function readRealtimeAsrConfig() {
    if (typeof localStorage === 'undefined') {
        return { realtimeAsrWsUrl: DEFAULT_REALTIME_ASR_WS, realtimeAsrApiKey: '' }
    }
    try {
        const raw = localStorage.getItem(WHISPER_ASR_STORAGE_KEY)
        if (!raw) {
            return { realtimeAsrWsUrl: DEFAULT_REALTIME_ASR_WS, realtimeAsrApiKey: '' }
        }
        const o = JSON.parse(raw)
        return {
            realtimeAsrWsUrl: String(o.realtimeAsrWsUrl || o.asrWsUrl || '').trim() || DEFAULT_REALTIME_ASR_WS,
            realtimeAsrApiKey: String(o.realtimeAsrApiKey ?? o.asrApiKey ?? '').trim(),
        }
    } catch {
        return { realtimeAsrWsUrl: DEFAULT_REALTIME_ASR_WS, realtimeAsrApiKey: '' }
    }
}

function realtimeAsrBrowserWsUrl() {
    if (typeof window === 'undefined') {
        return ''
    }
    const proto = window.location.protocol === 'https:' ? 'wss' : 'ws'
    return `${proto}://${window.location.host}/api/ws/realtime-asr`
}

const phase = ref('idle')
/** 实时字幕全文 */
const transcript = ref('')
const minutesMd = ref('')
const todosInserted = ref(0)
const liveTranscriptEl = ref(null)
const fileInputRef = ref(null)
const uploadProgress = ref(0)

let asrWs = null
let mediaRecorder = null
let recordStream = null
let recordIntervalId = null
let doneResolver = null
let donePromise = null

const minutesHtml = computed(() => {
    try {
        const raw = marked.parse(String(minutesMd.value || ''))
        return DOMPurify.sanitize(raw)
    } catch {
        return DOMPurify.sanitize(String(minutesMd.value || ''))
    }
})

const formatRecSeconds = computed(() => {
    const s = recordSeconds.value
    const m = Math.floor(s / 60)
    const sec = s % 60
    return `${m}:${String(sec).padStart(2, '0')}`
})

const recordSeconds = ref(0)

const isMeetingBusy = computed(() => ['connecting', 'summarizing', 'fileTranscribing'].includes(phase.value))

watch(transcript, async () => {
    await nextTick()
    const el = liveTranscriptEl.value
    if (el) {
        el.scrollTop = el.scrollHeight
    }
})

function clearRecordTimer() {
    if (recordIntervalId != null) {
        clearInterval(recordIntervalId)
        recordIntervalId = null
    }
}

function stopMicTracks() {
    try {
        recordStream?.getTracks()?.forEach((t) => t.stop())
    } catch {
        /* ignore */
    }
    recordStream = null
}

function pickWebmMimeType() {
    const cands = ['audio/webm;codecs=opus', 'audio/webm']
    for (const c of cands) {
        try {
            if (typeof MediaRecorder !== 'undefined' && MediaRecorder.isTypeSupported(c)) {
                return c
            }
        } catch {
            /* ignore */
        }
    }
    return ''
}

function resetFlow() {
    phase.value = 'idle'
    transcript.value = ''
    minutesMd.value = ''
    todosInserted.value = 0
    recordSeconds.value = 0
    uploadProgress.value = 0
}

function closeAsrWs() {
    try {
        asrWs?.close()
    } catch {
        /* ignore */
    }
    asrWs = null
}

function killRecordingPipeline() {
    clearRecordTimer()
    try {
        if (mediaRecorder && mediaRecorder.state !== 'inactive') {
            mediaRecorder.stop()
        }
    } catch {
        /* ignore */
    }
    mediaRecorder = null
    stopMicTracks()
    closeAsrWs()
    doneResolver = null
    donePromise = null
    if (phase.value === 'recording' || phase.value === 'paused' || phase.value === 'connecting') {
        phase.value = 'idle'
    }
}

function handleAsrParsed(msg) {
    const t = msg.type
    if (t === 'transcript') {
        if (msg.op === 'set' && typeof msg.text === 'string') {
            transcript.value = msg.text
        } else if (typeof msg.text === 'string') {
            transcript.value += msg.text
        }
        return
    }
    if (t === 'ready') {
        return
    }
    if (t === 'error') {
        ElMessage.error(String(msg.message || '实时转写错误'))
        killRecordingPipeline()
        return
    }
    if (t === 'done') {
        if (typeof msg.transcript === 'string' && msg.transcript.trim()) {
            transcript.value = msg.transcript.trim()
        }
        try {
            doneResolver?.()
        } catch {
            /* ignore */
        }
        doneResolver = null
    }
}

async function startRecording() {
    if (!userStore.user?.email) {
        ElMessage.warning('请先登录后再使用现场录音')
        return
    }
    if (isMeetingBusy.value || phase.value === 'recording' || phase.value === 'paused') {
        return
    }
    const cfg = readRealtimeAsrConfig()
    if (!cfg.realtimeAsrApiKey) {
        ElMessage.warning('请先在「API & 模型配置」→ 语音识别中填写「ASR API Key（实时 · 百炼）」并保存到本机')
        return
    }
    if (typeof navigator === 'undefined' || !navigator.mediaDevices?.getUserMedia) {
        ElMessage.error('当前浏览器不支持现场录音')
        return
    }

    phase.value = 'connecting'
    transcript.value = ''
    minutesMd.value = ''
    todosInserted.value = 0
    recordSeconds.value = 0

    let stream
    try {
        stream = await navigator.mediaDevices.getUserMedia({ audio: true, video: false })
    } catch (e) {
        phase.value = 'idle'
        const name = e?.name || ''
        if (name === 'NotAllowedError' || name === 'PermissionDeniedError') {
            ElMessage.error('麦克风权限被拒绝：请在浏览器设置中允许本站使用麦克风后重试')
        } else if (name === 'NotFoundError' || name === 'DevicesNotFoundError') {
            ElMessage.error('未检测到可用麦克风设备')
        } else {
            ElMessage.error(String(e?.message || e || '无法开启麦克风'))
        }
        return
    }
    recordStream = stream

    const url = realtimeAsrBrowserWsUrl()
    const ws = new WebSocket(url)
    asrWs = ws
    ws.binaryType = 'arraybuffer'

    donePromise = new Promise((resolve) => {
        doneResolver = resolve
    })

    let wsPhase = 'handshake'
    let handshakeResolve = null
    let handshakeReject = null
    const handshakePromise = new Promise((resolve, reject) => {
        handshakeResolve = resolve
        handshakeReject = reject
    })
    const handshakeTimer = setTimeout(() => {
        if (wsPhase === 'handshake') {
            handshakeReject?.(new Error('连接实时服务超时'))
        }
    }, 20000)

    ws.onmessage = (ev) => {
        if (typeof ev.data !== 'string') {
            return
        }
        let msg
        try {
            msg = JSON.parse(ev.data)
        } catch {
            return
        }
        if (wsPhase === 'handshake') {
            if (msg.type === 'ready') {
                clearTimeout(handshakeTimer)
                wsPhase = 'open'
                handshakeResolve?.()
                return
            }
            if (msg.type === 'error') {
                clearTimeout(handshakeTimer)
                handshakeReject?.(new Error(msg.message || '握手失败'))
                return
            }
            return
        }
        handleAsrParsed(msg)
    }
    ws.onerror = () => {
        clearTimeout(handshakeTimer)
        if (wsPhase === 'handshake') {
            try {
                handshakeReject?.(new Error('WebSocket 错误'))
            } catch {
                /* ignore */
            }
        }
        ElMessage.error('与实时转写服务连接失败')
        killRecordingPipeline()
    }
    ws.onclose = () => {
        clearTimeout(handshakeTimer)
        if (wsPhase === 'handshake') {
            try {
                handshakeReject?.(new Error('连接已断开'))
            } catch {
                /* ignore */
            }
        }
        asrWs = null
        try {
            doneResolver?.()
        } catch {
            /* ignore */
        }
    }

    ws.onopen = () => {
        ws.send(
            JSON.stringify({
                type: 'hello',
                asrWsUrl: cfg.realtimeAsrWsUrl || DEFAULT_REALTIME_ASR_WS,
                asrApiKey: cfg.realtimeAsrApiKey,
            }),
        )
    }

    try {
        await handshakePromise
    } catch (e) {
        ElMessage.error(String(e?.message || e))
        killRecordingPipeline()
        return
    }

    const mime = pickWebmMimeType()
    const mr = mime ? new MediaRecorder(stream, { mimeType: mime }) : new MediaRecorder(stream)
    mediaRecorder = mr

    mr.ondataavailable = async (ev) => {
        if (!ev.data || ev.data.size === 0 || !asrWs || asrWs.readyState !== WebSocket.OPEN) {
            return
        }
        try {
            const buf = await ev.data.arrayBuffer()
            asrWs.send(buf)
        } catch {
            /* ignore */
        }
    }
    mr.onerror = () => {
        ElMessage.error('录音过程出错，请重试')
        killRecordingPipeline()
    }

    mr.start(250)
    phase.value = 'recording'
    clearRecordTimer()
    recordIntervalId = window.setInterval(() => {
        recordSeconds.value += 1
    }, 1000)
}

function pauseRecording() {
    if (!mediaRecorder || phase.value !== 'recording') {
        return
    }
    try {
        mediaRecorder.pause()
        phase.value = 'paused'
        clearRecordTimer()
    } catch {
        /* ignore */
    }
}

function resumeRecording() {
    if (!mediaRecorder || phase.value !== 'paused') {
        return
    }
    try {
        mediaRecorder.resume()
        phase.value = 'recording'
        clearRecordTimer()
        recordIntervalId = window.setInterval(() => {
            recordSeconds.value += 1
        }, 1000)
    } catch {
        /* ignore */
    }
}

async function stopAndSummarize() {
    if (!mediaRecorder || (phase.value !== 'recording' && phase.value !== 'paused')) {
        return
    }
    clearRecordTimer()
    const mr = mediaRecorder
    mediaRecorder = null

    await new Promise((resolve) => {
        mr.addEventListener(
            'stop',
            () => {
                resolve()
            },
            { once: true },
        )
        try {
            if (mr.state !== 'inactive') {
                mr.stop()
            } else {
                resolve()
            }
        } catch {
            resolve()
        }
    })

    stopMicTracks()

    const ws = asrWs
    if (ws && ws.readyState === WebSocket.OPEN) {
        ws.send(JSON.stringify({ type: 'finalize' }))
    }

    const email = userStore.user?.email
    const wait = donePromise || Promise.resolve()
    try {
        await Promise.race([wait, new Promise((r) => setTimeout(r, 120000))])
    } catch {
        /* ignore */
    }
    closeAsrWs()

    const finalText = String(transcript.value || '').trim()
    if (!finalText) {
        ElMessage.warning('未收到有效转写文本，请检查实时 ASR 配置与麦克风')
        phase.value = 'idle'
        return
    }

    if (!email) {
        phase.value = 'idle'
        return
    }

    phase.value = 'summarizing'
    const ac = new AbortController()
    const timer = setTimeout(() => ac.abort(), FETCH_TIMEOUT_MS)
    try {
        const r = await fetch(MEETING_PROCESS_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                userEmail: email,
                transcript: finalText,
            }),
            signal: ac.signal,
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        transcript.value = d.transcript || finalText
        minutesMd.value = d.minutesMarkdown || ''
        todosInserted.value = Number(d.todosInserted || 0)
        phase.value = 'result'
        emit('processed', { todosInserted: todosInserted.value })
        if (todosInserted.value > 0) {
            ElMessage.success(
                `已为您自动提取并创建了 ${todosInserted.value} 条待办事项，请在「我的待办」中查看`,
            )
        } else {
            ElMessage.success('实时转写与纪要已生成')
        }
    } catch (e) {
        const msg =
            e?.name === 'AbortError' ? '请求超时，请稍后重试' : String(e?.message || e || '纪要生成失败')
        ElMessage.error(msg)
        phase.value = 'idle'
    } finally {
        clearTimeout(timer)
    }
}

function openMeetingFilePicker() {
    if (!userStore.user?.email) {
        ElMessage.warning('请先登录')
        return
    }
    if (isMeetingBusy.value || phase.value === 'recording' || phase.value === 'paused' || phase.value === 'connecting') {
        return
    }
    const cfg = readRealtimeAsrConfig()
    if (!cfg.realtimeAsrApiKey) {
        ElMessage.warning('请先在「API & 模型配置」中填写并保存「ASR API Key（实时 · 百炼）」')
        return
    }
    fileInputRef.value?.click()
}

/**
 * @param {File} file
 * @param {FormData} formData
 * @param {{ signal?: AbortSignal }} [opts]
 * @returns {Promise<Record<string, unknown>>}
 */
function postFileTranscribeWithProgress(file, formData, opts = {}) {
    return new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest()
        xhr.open('POST', FILE_TRANSCRIBE_URL)
        xhr.responseType = 'json'
        xhr.timeout = FETCH_TIMEOUT_MS
        xhr.upload.onprogress = (e) => {
            if (e.lengthComputable) {
                uploadProgress.value = Math.min(99, Math.round((e.loaded / e.total) * 100))
            }
        }
        xhr.onload = () => {
            uploadProgress.value = 100
            let body = xhr.response
            if (typeof body === 'string') {
                try {
                    body = JSON.parse(body)
                } catch {
                    body = null
                }
            }
            if (xhr.status >= 200 && xhr.status < 300 && body && typeof body === 'object' && body.success) {
                resolve(body)
                return
            }
            const msg =
                (body && typeof body === 'object' && (body.message || body.error)) ||
                (typeof xhr.responseText === 'string' && xhr.responseText.slice(0, 200)) ||
                `HTTP ${xhr.status}`
            reject(new Error(String(msg)))
        }
        xhr.onerror = () => reject(new Error('网络错误，上传失败'))
        xhr.onabort = () => reject(new DOMException('Aborted', 'AbortError'))
        xhr.ontimeout = () => reject(new Error('请求超时'))
        if (opts.signal) {
            const onAbort = () => {
                xhr.abort()
                reject(new DOMException('Aborted', 'AbortError'))
            }
            if (opts.signal.aborted) {
                onAbort()
                return
            }
            opts.signal.addEventListener('abort', onAbort, { once: true })
        }
        xhr.send(formData)
    })
}

async function onMeetingFileSelected(ev) {
    const input = ev.target
    const file = input?.files?.[0]
    if (input) {
        input.value = ''
    }
    if (!file) {
        return
    }
    const name = String(file.name || '').toLowerCase()
    if (!name.endsWith('.mp3') && !name.endsWith('.wav') && !name.endsWith('.m4a')) {
        ElMessage.error('仅支持 .mp3 / .wav / .m4a')
        return
    }
    const maxBytes = 52 * 1024 * 1024
    if (file.size > maxBytes) {
        ElMessage.error('文件需小于 52MB')
        return
    }
    const email = userStore.user?.email
    if (!email) {
        ElMessage.warning('请先登录')
        return
    }
    const cfg = readRealtimeAsrConfig()
    if (!cfg.realtimeAsrApiKey) {
        ElMessage.warning('请配置实时 ASR API Key')
        return
    }

    phase.value = 'fileTranscribing'
    transcript.value = ''
    minutesMd.value = ''
    todosInserted.value = 0
    uploadProgress.value = 0

    const fd = new FormData()
    fd.append('file', file)
    fd.append('userEmail', email)
    fd.append('asrWsUrl', cfg.realtimeAsrWsUrl || DEFAULT_REALTIME_ASR_WS)
    fd.append('asrApiKey', cfg.realtimeAsrApiKey)

    const ac = new AbortController()
    const timer = setTimeout(() => ac.abort(), FETCH_TIMEOUT_MS)
    try {
        const d = await postFileTranscribeWithProgress(file, fd, { signal: ac.signal })
        transcript.value = String(d.transcript || '').trim()
        minutesMd.value = String(d.minutesMarkdown || '')
        todosInserted.value = Number(d.todosInserted || 0)
        phase.value = 'result'
        emit('processed', { todosInserted: todosInserted.value })
        if (todosInserted.value > 0) {
            ElMessage.success(
                `已为您自动提取并创建了 ${todosInserted.value} 条待办事项，请在「我的待办」中查看`,
            )
        } else {
            ElMessage.success('上传转写与纪要已生成')
        }
    } catch (e) {
        const msg =
            e?.name === 'AbortError' ? '请求超时，请稍后重试或缩短音频' : String(e?.message || e || '处理失败')
        ElMessage.error(msg)
        phase.value = 'idle'
    } finally {
        clearTimeout(timer)
        uploadProgress.value = 0
    }
}

async function copyMinutes() {
    const t = minutesMd.value.trim()
    if (!t) {
        ElMessage.info('暂无纪要可复制')
        return
    }
    try {
        await navigator.clipboard.writeText(t)
        ElMessage.success('已复制到剪贴板')
    } catch {
        ElMessage.warning('复制失败，请手动选择文本')
    }
}

onBeforeUnmount(() => {
    killRecordingPipeline()
})
</script>

<template>
    <div class="meeting-minutes-view">
        <div v-if="phase === 'result'" class="meeting-result-split">
            <div class="meeting-result-toolbar">
                <span class="meeting-result-done">处理完成（实时转写 / 上传转写 + 纪要）</span>
                <button type="button" class="btn btn-outline btn-sm" @click="resetFlow">再录一段</button>
            </div>
            <div class="meeting-result-cols">
                <section class="meeting-result-col meeting-result-col--left">
                    <h3 class="meeting-result-col-title">转写原文（ASR）</h3>
                    <div class="meeting-result-scroll">
                        <pre class="meeting-result-pre">{{ transcript }}</pre>
                    </div>
                </section>
                <section class="meeting-result-col meeting-result-col--right">
                    <div class="meeting-result-col-head">
                        <h3 class="meeting-result-col-title">智能纪要</h3>
                        <button type="button" class="btn btn-outline btn-xs" @click="copyMinutes">一键复制</button>
                    </div>
                    <div class="meeting-result-scroll meeting-result-scroll--md">
                        <div class="meeting-mv-md markdown-body" v-html="minutesHtml"></div>
                    </div>
                </section>
            </div>
        </div>

        <div v-else-if="phase === 'summarizing'" class="meeting-stream-dual">
            <section class="meeting-stream-col meeting-stream-col--left">
                <h3 class="meeting-stream-title">转写原文</h3>
                <div class="meeting-stream-pre-wrap">
                    <pre class="meeting-stream-pre" role="log" aria-live="polite">{{ transcript }}</pre>
                </div>
            </section>
            <section class="meeting-stream-col meeting-stream-col--right">
                <h3 class="meeting-stream-title">状态</h3>
                <div class="meeting-stream-status">
                    <span class="meeting-mv-spinner meeting-mv-spinner--inline" aria-hidden="true"></span>
                    <p>正在根据实时字幕生成「智能会议纪要」…</p>
                </div>
            </section>
        </div>

        <div v-else-if="phase === 'fileTranscribing'" class="meeting-file-processing">
            <div class="meeting-file-processing-inner">
                <span class="meeting-mv-spinner" aria-hidden="true"></span>
                <p class="meeting-file-processing-title">正在上传并转写音频…</p>
                <p class="meeting-file-processing-hint">文件经磁盘缓冲后流式送 ASR，完成后将自动生成纪要与待办</p>
                <div class="meeting-upload-progress" role="progressbar" :aria-valuenow="uploadProgress" aria-valuemin="0" aria-valuemax="100">
                    <div class="meeting-upload-progress__fill" :style="{ width: `${uploadProgress}%` }"></div>
                </div>
                <p class="meeting-upload-progress-label">{{ uploadProgress }}%（上传阶段进度；转写完成后将跳转结果）</p>
            </div>
        </div>

        <div v-else class="meeting-upload-card">
            <h2 class="meeting-mv-title">智能会议纪要</h2>
            <p class="meeting-mv-lead">
                <strong>实时</strong>：麦克风 WebM 经全双工 WebSocket 由服务端 FFmpeg 转 PCM 后送 Qwen3 ASR，边录边出字。<strong>离线</strong>：可上传
                .mp3 / .wav / .m4a（≤52MB），服务端落盘缓冲后同样走 ASR，并自动提取待办写入「我的待办」。请在「API 与模型配置」→ <strong>语音识别 (ASR)</strong> 中填写百炼 Key。
            </p>

            <input
                ref="fileInputRef"
                type="file"
                class="meeting-mv-sr-only"
                accept=".mp3,.wav,.m4a,audio/mpeg,audio/wav,audio/x-m4a,audio/mp4"
                aria-hidden="true"
                tabindex="-1"
                @change="onMeetingFileSelected"
            />

            <section class="meeting-live-transcript-box" aria-label="实时字幕">
                <div class="meeting-live-transcript-head">
                    <span class="meeting-live-transcript-label">实时字幕</span>
                    <span v-if="phase === 'recording'" class="meeting-live-transcript-badge">录音中</span>
                    <span v-else-if="phase === 'paused'" class="meeting-live-transcript-badge meeting-live-transcript-badge--paused">已暂停</span>
                    <span v-else-if="phase === 'connecting'" class="meeting-live-transcript-badge meeting-live-transcript-badge--wait">连接中</span>
                </div>
                <div ref="liveTranscriptEl" class="meeting-live-transcript-scroll">
                    <p class="meeting-live-transcript-text">{{ transcript || '（开始录音后，识别文字将显示于此）' }}</p>
                </div>
            </section>

            <div class="meeting-live-bar meeting-live-bar--controls">
                <button
                    type="button"
                    class="btn btn-outline meeting-live-btn"
                    :disabled="!userStore.user?.email || isMeetingBusy || phase === 'recording' || phase === 'paused' || phase === 'connecting'"
                    @click="openMeetingFilePicker"
                >
                    <span class="meeting-btn-inner"><el-icon :size="18"><Folder /></el-icon>上传录音文件</span>
                </button>
                <button
                    type="button"
                    class="btn btn-primary meeting-live-btn"
                    :disabled="!userStore.user?.email || isMeetingBusy || phase === 'recording' || phase === 'paused' || phase === 'connecting'"
                    @click="startRecording"
                >
                    <span class="meeting-btn-inner"><el-icon :size="18"><VideoCamera /></el-icon>开始录音</span>
                </button>
                <button
                    type="button"
                    class="btn btn-outline meeting-live-btn"
                    :disabled="phase !== 'recording'"
                    @click="pauseRecording"
                >
                    <span class="meeting-btn-inner"><el-icon :size="18"><VideoPause /></el-icon>暂停</span>
                </button>
                <button
                    type="button"
                    class="btn btn-outline meeting-live-btn"
                    :disabled="phase !== 'paused'"
                    @click="resumeRecording"
                >
                    <span class="meeting-btn-inner"><el-icon :size="18"><VideoPlay /></el-icon>继续</span>
                </button>
                <button
                    type="button"
                    class="btn btn-accent meeting-live-btn"
                    :disabled="phase !== 'recording' && phase !== 'paused'"
                    @click="stopAndSummarize"
                >
                    <span class="meeting-btn-inner"><el-icon :size="18"><SwitchButton /></el-icon>结束并生成纪要</span>
                </button>
                <span class="meeting-live-hint" v-if="phase === 'recording' || phase === 'paused'">
                    <span class="meeting-rec-dot" aria-hidden="true"></span>
                    <span class="meeting-rec-timer" aria-live="polite">{{ formatRecSeconds }}</span>
                </span>
            </div>
        </div>
    </div>
</template>

<style scoped>
.meeting-btn-inner {
    display: inline-flex;
    align-items: center;
    gap: 6px;
}

.meeting-minutes-view {
    max-width: 1280px;
    margin: 0 auto;
    padding: 16px 20px 32px;
    box-sizing: border-box;
}

.meeting-upload-card {
    background: #fff;
    border: 1px solid #e5e7eb;
    border-radius: 16px;
    padding: 24px 22px 28px;
    box-shadow: 0 4px 24px rgba(15, 23, 42, 0.06);
}

.meeting-mv-sr-only {
    position: absolute;
    width: 1px;
    height: 1px;
    padding: 0;
    margin: -1px;
    overflow: hidden;
    clip: rect(0, 0, 0, 0);
    white-space: nowrap;
    border: 0;
}

.meeting-file-processing {
    max-width: 560px;
    margin: 0 auto;
    padding: 48px 20px;
}

.meeting-file-processing-inner {
    text-align: center;
    padding: 28px 22px;
    background: #fff;
    border: 1px solid #e5e7eb;
    border-radius: 16px;
    box-shadow: 0 4px 24px rgba(15, 23, 42, 0.06);
}

.meeting-file-processing-title {
    margin: 12px 0 6px;
    font-size: 1.05rem;
    font-weight: 700;
    color: #0f172a;
}

.meeting-file-processing-hint {
    margin: 0 0 18px;
    font-size: 0.85rem;
    color: #64748b;
    line-height: 1.5;
}

.meeting-upload-progress {
    height: 10px;
    border-radius: 999px;
    background: #e2e8f0;
    overflow: hidden;
    max-width: 360px;
    margin: 0 auto;
}

.meeting-upload-progress__fill {
    height: 100%;
    border-radius: 999px;
    background: linear-gradient(90deg, #3b82f6, #6366f1);
    transition: width 0.2s ease;
}

.meeting-upload-progress-label {
    margin: 10px 0 0;
    font-size: 0.78rem;
    color: #94a3b8;
}

.meeting-mv-title {
    margin: 0 0 8px;
    font-size: 1.35rem;
    font-weight: 700;
    color: #0f172a;
    letter-spacing: -0.02em;
}

.meeting-mv-lead {
    margin: 0 0 16px;
    font-size: 0.9rem;
    color: #64748b;
    line-height: 1.55;
}

.meeting-live-transcript-box {
    margin-bottom: 18px;
    border-radius: 12px;
    border: 1px solid #e2e8f0;
    background: linear-gradient(180deg, #f8fafc 0%, #fff 48%);
    overflow: hidden;
}

.meeting-live-transcript-head {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 10px 14px;
    border-bottom: 1px solid #e2e8f0;
    background: #f1f5f9;
}

.meeting-live-transcript-label {
    font-size: 0.82rem;
    font-weight: 700;
    color: #0f172a;
}

.meeting-live-transcript-badge {
    font-size: 0.72rem;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 999px;
    background: #fee2e2;
    color: #b91c1c;
}

.meeting-live-transcript-badge--paused {
    background: #fef3c7;
    color: #b45309;
}

.meeting-live-transcript-badge--wait {
    background: #e0e7ff;
    color: #4338ca;
}

.meeting-live-transcript-scroll {
    min-height: 140px;
    max-height: min(36vh, 320px);
    overflow-y: auto;
    padding: 12px 14px 14px;
    scroll-behavior: smooth;
}

.meeting-live-transcript-text {
    margin: 0;
    font-size: 0.95rem;
    line-height: 1.65;
    color: #1e293b;
    white-space: pre-wrap;
    word-break: break-word;
}

.meeting-live-bar {
    padding: 14px 16px;
    border-radius: 12px;
    background: #f8fafc;
    border: 1px solid #e2e8f0;
}

.meeting-live-bar--controls {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 10px;
}

.meeting-live-btn {
    margin: 0;
}

.meeting-live-hint {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    font-size: 0.8rem;
    color: #64748b;
}

.meeting-rec-dot {
    display: inline-block;
    width: 10px;
    height: 10px;
    border-radius: 50%;
    background: #ef4444;
    animation: meeting-rec-pulse 1.1s ease-out infinite;
}

@keyframes meeting-rec-pulse {
    0% {
        box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.55);
    }
    70% {
        box-shadow: 0 0 0 10px rgba(239, 68, 68, 0);
    }
    100% {
        box-shadow: 0 0 0 0 rgba(239, 68, 68, 0);
    }
}

.meeting-rec-timer {
    font-variant-numeric: tabular-nums;
    font-weight: 700;
    font-size: 1.05rem;
    color: #b91c1c;
}

.meeting-stream-dual {
    display: grid;
    grid-template-columns: 1fr minmax(220px, 320px);
    gap: 16px;
    align-items: stretch;
    margin-bottom: 8px;
}

@media (max-width: 900px) {
    .meeting-stream-dual {
        grid-template-columns: 1fr;
    }
}

.meeting-stream-col {
    background: #fff;
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    padding: 16px 16px 14px;
    box-shadow: 0 2px 14px rgba(15, 23, 42, 0.05);
    min-height: 200px;
    display: flex;
    flex-direction: column;
}

.meeting-stream-title {
    margin: 0 0 10px;
    font-size: 0.95rem;
    font-weight: 700;
    color: #0f172a;
}

.meeting-stream-pre-wrap {
    flex: 1;
    min-height: 160px;
    max-height: min(52vh, 480px);
    overflow: auto;
    border-radius: 10px;
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    padding: 10px 12px;
}

.meeting-stream-pre {
    margin: 0;
    white-space: pre-wrap;
    word-break: break-word;
    font-size: 0.82rem;
    line-height: 1.55;
    color: #1e293b;
    font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
}

.meeting-stream-status {
    flex: 1;
    display: flex;
    flex-direction: column;
    align-items: flex-start;
    gap: 10px;
    color: #475569;
    font-size: 0.88rem;
    line-height: 1.5;
}

.meeting-mv-spinner {
    display: inline-block;
    width: 40px;
    height: 40px;
    margin-bottom: 12px;
    border-radius: 50%;
    border: 3px solid #e2e8f0;
    border-top-color: #3b82f6;
    animation: meeting-mv-spin 0.85s linear infinite;
}

.meeting-mv-spinner--inline {
    width: 22px;
    height: 22px;
    margin: 0 8px 0 0;
    vertical-align: middle;
    display: inline-block;
}

@keyframes meeting-mv-spin {
    to {
        transform: rotate(360deg);
    }
}

.meeting-result-split {
    display: flex;
    flex-direction: column;
    gap: 12px;
}

.meeting-result-toolbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
}

.meeting-result-done {
    font-size: 0.88rem;
    font-weight: 600;
    color: #0f172a;
}

.meeting-result-cols {
    display: grid;
    grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
    gap: 14px;
    min-height: min(72vh, 720px);
}

@media (max-width: 900px) {
    .meeting-result-cols {
        grid-template-columns: 1fr;
        min-height: auto;
    }
}

.meeting-result-col {
    display: flex;
    flex-direction: column;
    min-height: 0;
    background: #fff;
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    overflow: hidden;
}

.meeting-result-col-title {
    margin: 0;
    font-size: 0.92rem;
    font-weight: 700;
    color: #0f172a;
}

.meeting-result-col-head {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    padding: 12px 14px;
    border-bottom: 1px solid #f1f5f9;
    flex-shrink: 0;
}

.meeting-result-col--left .meeting-result-col-title {
    padding: 12px 14px;
    border-bottom: 1px solid #f1f5f9;
}

.meeting-result-scroll {
    flex: 1;
    min-height: 0;
    overflow: auto;
    padding: 12px 14px 16px;
}

.meeting-result-pre {
    margin: 0;
    white-space: pre-wrap;
    word-break: break-word;
    font-size: 0.82rem;
    line-height: 1.55;
    color: #334155;
    font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono', 'Courier New', monospace;
}

.meeting-mv-md {
    font-size: 0.88rem;
    line-height: 1.55;
}
</style>

<style>
.tab-meeting-minutes .meeting-mv-md.markdown-body {
    font-size: 0.88rem;
    line-height: 1.62;
    color: #1f2937;
    word-break: break-word;
    overflow-wrap: anywhere;
}
.tab-meeting-minutes .meeting-mv-md.markdown-body h2 {
    font-size: 1.05rem;
    margin: 1rem 0 0.5rem;
    color: #0f172a;
    border-bottom: 1px solid #e5e7eb;
    padding-bottom: 0.25rem;
}
.tab-meeting-minutes .meeting-mv-md.markdown-body p {
    margin: 0.45em 0;
}
.tab-meeting-minutes .meeting-mv-md.markdown-body ul {
    margin: 0.45em 0;
    padding-left: 1.25rem;
    list-style: disc;
}
.tab-meeting-minutes .meeting-mv-md.markdown-body li {
    margin: 0.25em 0;
}
.tab-meeting-minutes .meeting-mv-md.markdown-body strong {
    font-weight: 700;
    color: #0f172a;
}
</style>
