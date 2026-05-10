<script setup>
import { ref, computed, onBeforeUnmount } from 'vue'
import { ElMessage } from 'element-plus'
import { ElUpload } from 'element-plus'
import 'element-plus/es/components/upload/style/css'
import { marked } from 'marked'
import DOMPurify from 'dompurify'
import { useUserStore } from '../store/user.js'

marked.setOptions({
    gfm: true,
    breaks: true,
})

const emit = defineEmits(['processed'])

const userStore = useUserStore()

/** 与会议纪要后端一致：上传/现场录音均走 process（含 Whisper + 纪要） */
const MEETING_PROCESS_URL = '/api/meeting/process'
/** 与 DashboardView「API & 模型配置」中 ASR 区块共用，便于随 FormData 带给后端 */
const WHISPER_ASR_STORAGE_KEY = 'openclaw_whisper_asr_v1'
const DEFAULT_WHISPER_ASR_BASE = 'https://api.openai.com/v1'
const MAX_BYTES = 20 * 1024 * 1024
const FETCH_TIMEOUT_MS = 900000

function readWhisperAsrForMeeting() {
    if (typeof localStorage === 'undefined') {
        return { whisperApiKey: '', whisperBaseUrl: DEFAULT_WHISPER_ASR_BASE }
    }
    try {
        const raw = localStorage.getItem(WHISPER_ASR_STORAGE_KEY)
        if (!raw) {
            return { whisperApiKey: '', whisperBaseUrl: DEFAULT_WHISPER_ASR_BASE }
        }
        const o = JSON.parse(raw)
        return {
            whisperApiKey: String(o.whisperApiKey ?? '').trim(),
            whisperBaseUrl: String(o.whisperBaseUrl || DEFAULT_WHISPER_ASR_BASE).trim() || DEFAULT_WHISPER_ASR_BASE,
        }
    } catch {
        return { whisperApiKey: '', whisperBaseUrl: DEFAULT_WHISPER_ASR_BASE }
    }
}

const phase = ref('idle')
const transcript = ref('')
const minutesMd = ref('')
const todosInserted = ref(0)
const uploadKey = ref(0)
const recordSeconds = ref(0)

let mediaRecorder = null
let recordStream = null
let recordChunks = []
let recordIntervalId = null

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

function pickRecorderMimeType() {
    const cands = [
        'audio/webm;codecs=opus',
        'audio/webm',
        'audio/ogg;codecs=opus',
        'audio/ogg',
        'audio/mp4',
    ]
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
    uploadKey.value += 1
}

function beforeUpload(file) {
    const name = String(file?.name || '').toLowerCase()
    const ok =
        name.endsWith('.mp3') ||
        name.endsWith('.wav') ||
        name.endsWith('.m4a') ||
        name.endsWith('.webm') ||
        name.endsWith('.ogg')
    if (!ok) {
        ElMessage.error('仅支持 .mp3 / .wav / .m4a / .webm / .ogg')
        return false
    }
    if (file.size > MAX_BYTES) {
        ElMessage.error('文件需小于 20MB')
        return false
    }
    if (!userStore.user?.email) {
        ElMessage.warning('请先登录后再上传会议录音')
        return false
    }
    return true
}

async function submitMeetingAudioFile(file) {
    const email = userStore.user?.email
    if (!email) {
        ElMessage.warning('请先登录')
        return
    }
    phase.value = 'working'
    transcript.value = ''
    minutesMd.value = ''
    const fd = new FormData()
    fd.append('userEmail', email)
    fd.append('audio', file)
    const asr = readWhisperAsrForMeeting()
    fd.append('whisperBaseUrl', asr.whisperBaseUrl)
    if (asr.whisperApiKey) {
        fd.append('whisperApiKey', asr.whisperApiKey)
    }
    const ac = new AbortController()
    const timer = setTimeout(() => ac.abort(), FETCH_TIMEOUT_MS)
    try {
        const r = await fetch(MEETING_PROCESS_URL, {
            method: 'POST',
            body: fd,
            signal: ac.signal,
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        transcript.value = d.transcript || ''
        minutesMd.value = d.minutesMarkdown || ''
        todosInserted.value = Number(d.todosInserted || 0)
        phase.value = 'result'
        emit('processed', { todosInserted: todosInserted.value })
        if (todosInserted.value > 0) {
            ElMessage.success(`处理完成，已写入 ${todosInserted.value} 条待办提醒`)
        } else {
            ElMessage.success('转写与纪要已生成')
        }
    } catch (e) {
        const msg =
            e?.name === 'AbortError' ? '请求超时，请稍后重试或缩短录音后分段上传' : String(e?.message || e || '处理失败')
        ElMessage.error(msg)
        phase.value = 'idle'
        throw e
    } finally {
        clearTimeout(timer)
    }
}

async function customRequest(options) {
    try {
        await submitMeetingAudioFile(options.file)
        options.onSuccess?.({})
    } catch (e) {
        options.onError?.(e)
    }
}

async function startRecording() {
    if (!userStore.user?.email) {
        ElMessage.warning('请先登录后再使用现场录音')
        return
    }
    if (phase.value === 'working' || phase.value === 'recording') {
        return
    }
    if (typeof navigator === 'undefined' || !navigator.mediaDevices?.getUserMedia) {
        ElMessage.error('当前浏览器不支持现场录音')
        return
    }
    try {
        const stream = await navigator.mediaDevices.getUserMedia({ audio: true, video: false })
        recordStream = stream
        recordChunks = []
        const mime = pickRecorderMimeType()
        const mr = mime ? new MediaRecorder(stream, { mimeType: mime }) : new MediaRecorder(stream)
        mediaRecorder = mr
        mr.ondataavailable = (ev) => {
            if (ev.data && ev.data.size > 0) {
                recordChunks.push(ev.data)
            }
        }
        mr.onerror = () => {
            ElMessage.error('录音过程出错，请重试')
            stopLiveRecordingCleanup()
        }
        mr.start(1000)
        phase.value = 'recording'
        recordSeconds.value = 0
        clearRecordTimer()
        recordIntervalId = window.setInterval(() => {
            recordSeconds.value += 1
        }, 1000)
    } catch (e) {
        const name = e?.name || ''
        if (name === 'NotAllowedError' || name === 'PermissionDeniedError') {
            ElMessage.error('麦克风权限被拒绝：请在浏览器设置中允许本站使用麦克风后重试')
        } else if (name === 'NotFoundError' || name === 'DevicesNotFoundError') {
            ElMessage.error('未检测到可用麦克风设备')
        } else {
            ElMessage.error(String(e?.message || e || '无法开启麦克风'))
        }
    }
}

function stopLiveRecordingCleanup() {
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
    recordChunks = []
    if (phase.value === 'recording') {
        phase.value = 'idle'
    }
}

async function stopRecording() {
    if (!mediaRecorder || phase.value !== 'recording') {
        return
    }
    clearRecordTimer()
    const mr = mediaRecorder
    mediaRecorder = null
    await new Promise((resolve) => {
        mr.addEventListener('stop', () => resolve(), { once: true })
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
    const blob = new Blob(recordChunks, { type: mr.mimeType || 'audio/webm' })
    recordChunks = []
    const ext = String(blob.type || '').includes('mp4') ? 'm4a' : blob.type.includes('ogg') ? 'ogg' : 'webm'
    const fname = `live-recording-${Date.now()}.${ext}`
    const file = new File([blob], fname, { type: blob.type || `audio/${ext === 'webm' ? 'webm' : ext}` })
    if (file.size > MAX_BYTES) {
        phase.value = 'idle'
        ElMessage.error(`录音超过 ${MAX_BYTES / 1024 / 1024}MB，请缩短录制时长后重试`)
        return
    }
    if (file.size < 256) {
        phase.value = 'idle'
        ElMessage.warning('录音过短，未检测到有效音频')
        return
    }
    try {
        await submitMeetingAudioFile(file)
    } catch {
        /* submitMeetingAudioFile 已 toast */
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
    recordChunks = []
})
</script>

<template>
    <div class="meeting-minutes-view">
        <div v-if="phase !== 'result'" class="meeting-upload-card">
            <h2 class="meeting-mv-title">智能会议纪要</h2>
            <p class="meeting-mv-lead">
                现场录音或上传文件：浏览器将使用 MediaRecorder 采集麦克风；服务端通过 OpenAI 兼容 Whisper 接口做<strong>真实转写</strong>，再生成结构化纪要。请在「API
                与模型配置」页的<strong>语音识别 (ASR) 配置</strong>中填写 Whisper Base URL（默认 OpenAI）与可选专用 Key；也可在服务器环境变量中设置
                <code>WHISPER_BASE_URL</code> / <code>WHISPER_API_KEY</code>。
            </p>

            <div class="meeting-live-bar">
                <template v-if="phase !== 'recording'">
                    <button
                        type="button"
                        class="btn btn-primary meeting-live-btn"
                        :disabled="!userStore.user?.email || phase === 'working'"
                        @click="startRecording"
                    >
                        开始录音
                    </button>
                    <span class="meeting-live-hint">需授权麦克风；停止后自动上传并转写</span>
                </template>
                <template v-else>
                    <div class="meeting-recording-row">
                        <button type="button" class="btn btn-accent meeting-stop-btn" @click="stopRecording">停止录音</button>
                        <span class="meeting-rec-dot" aria-hidden="true" title="正在录音"></span>
                        <span class="meeting-rec-timer" aria-live="polite">{{ formatRecSeconds }}</span>
                    </div>
                </template>
            </div>

            <el-upload
                :key="uploadKey"
                class="meeting-mv-upload"
                drag
                :show-file-list="false"
                accept=".mp3,.wav,.m4a,.webm,.ogg,audio/*"
                :before-upload="beforeUpload"
                :http-request="customRequest"
                :disabled="phase === 'working' || phase === 'recording'"
            >
                <div class="meeting-mv-drop-inner">
                    <template v-if="phase !== 'working'">
                        <p class="meeting-mv-drop-icon" aria-hidden="true">&#128196;</p>
                        <p class="meeting-mv-drop-title">将录音拖到此处，或点击选择文件</p>
                        <p class="meeting-mv-drop-hint">支持 .mp3 / .wav / .m4a / .webm / .ogg，单文件不超过 20MB</p>
                    </template>
                    <template v-else>
                        <span class="meeting-mv-spinner" aria-hidden="true"></span>
                        <p class="meeting-mv-drop-title">正在真实转写中…</p>
                        <p class="meeting-mv-drop-hint">Whisper 转写与 AI 纪要生成中，请稍候（长音频或模型较慢时可能需数分钟）</p>
                    </template>
                </div>
            </el-upload>
        </div>

        <div v-else class="meeting-result-split">
            <div class="meeting-result-toolbar">
                <span class="meeting-result-done">处理完成（Whisper 转写 + 纪要）</span>
                <button type="button" class="btn btn-outline btn-sm" @click="resetFlow">重新上传 / 再录一段</button>
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
    </div>
</template>

<style scoped>
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

.meeting-mv-lead code {
    font-size: 0.78rem;
    padding: 1px 6px;
    border-radius: 6px;
    background: #f1f5f9;
    color: #0f172a;
}

.meeting-live-bar {
    margin-bottom: 18px;
    padding: 14px 16px;
    border-radius: 12px;
    background: #f8fafc;
    border: 1px solid #e2e8f0;
}

.meeting-live-btn {
    margin-right: 10px;
}

.meeting-live-hint {
    font-size: 0.8rem;
    color: #64748b;
}

.meeting-recording-row {
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 12px;
}

.meeting-stop-btn {
    font-weight: 600;
}

.meeting-rec-dot {
    display: inline-block;
    width: 10px;
    height: 10px;
    border-radius: 50%;
    background: #ef4444;
    box-shadow: 0 0 0 0 rgba(239, 68, 68, 0.55);
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

.meeting-mv-upload :deep(.el-upload) {
    width: 100%;
}

.meeting-mv-upload :deep(.el-upload-dragger) {
    width: 100%;
    padding: 48px 20px;
    border-radius: 14px;
    border: 2px dashed #cbd5e1;
    background: #f8fafc;
    transition: border-color 0.2s ease, background 0.2s ease;
}

.meeting-mv-upload :deep(.el-upload-dragger:hover) {
    border-color: #3b82f6;
    background: #eff6ff;
}

.meeting-mv-drop-inner {
    text-align: center;
}

.meeting-mv-drop-icon {
    margin: 0 0 10px;
    font-size: 2rem;
    line-height: 1;
}

.meeting-mv-drop-title {
    margin: 0 0 8px;
    font-size: 1rem;
    font-weight: 600;
    color: #1e293b;
}

.meeting-mv-drop-hint {
    margin: 0;
    font-size: 0.82rem;
    color: #94a3b8;
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
