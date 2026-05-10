<script setup>
import { ref, computed, nextTick, watch, onBeforeUnmount, onMounted } from 'vue'
import { useRouter, useRoute } from 'vue-router'
import { ElMessage, ElCard, ElButton, ElMessageBox, ElDialog, ElDrawer, ElEmpty, ElButtonGroup } from 'element-plus'
import 'element-plus/es/components/message/style/css'
import 'element-plus/es/components/card/style/css'
import 'element-plus/es/components/button/style/css'
import 'element-plus/es/components/message-box/style/css'
import 'element-plus/es/components/dialog/style/css'
import 'element-plus/es/components/drawer/style/css'
import 'element-plus/es/components/empty/style/css'
import { useUserStore } from '../store/user.js'
import { useOpenclawContextStore } from '../store/openclawContext.js'
import { renderAsync } from 'docx-preview'
import Docxtemplater from 'docxtemplater'
import PizZip from 'pizzip'
import { saveAs } from 'file-saver'
import { marked } from 'marked'
import DOMPurify from 'dompurify'
import MeetingMinutesView from '../components/MeetingMinutesView.vue'

marked.setOptions({
    gfm: true,
    breaks: true,
})

/** 助手正文 Markdown → 安全 HTML（流式过程中多次调用，失败时回退转义） */
function renderMarkdown(text) {
    if (text == null || text === '') {
        return ''
    }
    try {
        const rawHtml = marked.parse(String(text))
        return DOMPurify.sanitize(rawHtml)
    } catch {
        return DOMPurify.sanitize(String(text))
    }
}

const router = useRouter()
const route = useRoute()
const userStore = useUserStore()
const openclawCtx = useOpenclawContextStore()
const currentTab = ref('overview')
const quickActionOpen = ref(false)
const wordContainer = ref(null)
const isPreviewLoading = ref(false)
const previewError = ref('')
const extractedVariables = ref([])
const currentTemplateId = ref(null)
const isFilling = ref(false)
const fillError = ref('')
const currentFile = ref(null)
const templateArrayBuffer = ref(null)
const lastWordBlob = ref(null)
let fillAbortController = null
let syncTimer = null

const tabTitles = {
  overview: '工作台概览',
  templates: '智能模板引擎',
  chat: '对话助手',
  meeting: '智能会议纪要',
  paper: '文献论文摘要',
  ppt: 'PPT大纲生成',
  companion: '情感陪伴式对话',
  identity: '多身份切换',
  reminders: '主动提醒推送',
  apiConfig: 'API & 模型配置',
}

const pageTitle = computed(() => tabTitles[currentTab.value] || '工作台')

/** 对话助手 / 情感陪伴：全屏沉浸（与路由 meta.hideHeader 对齐） */
const isImmersiveWorkbenchTab = computed(
    () => route.meta.hideHeader === true || currentTab.value === 'chat' || currentTab.value === 'companion',
)

const dashboardStats = ref({
    conversation_turns: 0,
    papers_processed: 0,
    skills_count: 0,
})

const companionSettings = ref({
    persona: 'default',
    address: '',
    customTraits: '',
})

/** 主对话工作台身份（与「情感陪伴」人格分离）：personaId 随 /api/chat 上报，服务端写入 system */
const CHAT_PERSONA_STORAGE_KEY = 'openclaw_chat_persona_v1'
const BUILTIN_CHAT_PERSONAS = [
    { id: 'meeting_assistant', title: '会议助理', emoji: '📋', desc: '干练结构化，议程与行动项优先' },
    { id: 'strict_mentor', title: '严厉导师', emoji: '🎓', desc: '高标准反馈，先风险后步骤' },
    { id: 'geek_blaze', title: '爆裂极客', emoji: '⚡', desc: '快节奏技术梗，结论必须落地' },
    { id: 'workplace_helper', title: '职场助手', emoji: '💼', desc: '商务礼貌，结论先行' },
    { id: 'caring_partner', title: '知心大姐', emoji: '💬', desc: '先共情再建议，短句柔和' },
]

const chatActivePersonaId = ref('workplace_helper')
const chatCustomPersonas = ref([])
const customPersonaDialogOpen = ref(false)
const customPersonaForm = ref({
    name: '',
    emoji: '🙂',
    systemPrompt: '',
})

function sanitizeClientCustomPersonas(arr) {
    if (!Array.isArray(arr)) {
        return []
    }
    const out = []
    for (const item of arr.slice(0, 24)) {
        if (!item || typeof item !== 'object') {
            continue
        }
        const id = String(item.id || '').trim().slice(0, 80)
        if (!id.startsWith('custom_')) {
            continue
        }
        const name = String(item.name || '自定义伙伴').trim().slice(0, 80)
        const emoji = String(item.emoji || '').trim().slice(0, 8)
        const systemPrompt = String(item.systemPrompt || '').trim().slice(0, 2500)
        if (!systemPrompt) {
            continue
        }
        out.push({ id, name, emoji, systemPrompt })
    }
    return out
}

function loadChatPersonaLocalStorage() {
    if (typeof localStorage === 'undefined') {
        return
    }
    try {
        const raw = localStorage.getItem(CHAT_PERSONA_STORAGE_KEY)
        if (!raw) {
            return
        }
        const o = JSON.parse(raw)
        if (o.chat_active_persona_id && String(o.chat_active_persona_id).trim()) {
            chatActivePersonaId.value = String(o.chat_active_persona_id).trim().slice(0, 64)
        }
        if (Array.isArray(o.chat_custom_personas)) {
            chatCustomPersonas.value = sanitizeClientCustomPersonas(o.chat_custom_personas)
        }
    } catch {
        /* ignore */
    }
}

function saveChatPersonaLocalStorage() {
    if (typeof localStorage === 'undefined') {
        return
    }
    try {
        localStorage.setItem(
            CHAT_PERSONA_STORAGE_KEY,
            JSON.stringify({
                chat_active_persona_id: chatActivePersonaId.value,
                chat_custom_personas: chatCustomPersonas.value,
            }),
        )
    } catch {
        /* ignore */
    }
}

async function persistChatPersonaToServer() {
    const email = userStore.user?.email
    if (!email) {
        return
    }
    try {
        await fetch('/api/user/profile', {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                email,
                profile: {
                    chat_active_persona_id: chatActivePersonaId.value,
                    chat_custom_personas: chatCustomPersonas.value,
                },
            }),
        })
    } catch {
        /* ignore */
    }
}

function selectWorkbenchPersona(id) {
    const sid = String(id || '').trim().slice(0, 64)
    if (!sid) {
        return
    }
    chatActivePersonaId.value = sid
    saveChatPersonaLocalStorage()
    void persistChatPersonaToServer()
    ElMessage.success('对话身份已切换，打开「对话助手」发送消息即可按新身份回复')
}

function openCustomPersonaDialog() {
    customPersonaForm.value = { name: '', emoji: '🙂', systemPrompt: '' }
    customPersonaDialogOpen.value = true
}

async function submitCustomPersona() {
    const name = String(customPersonaForm.value.name || '').trim().slice(0, 80)
    const emoji = String(customPersonaForm.value.emoji || '').trim().slice(0, 8) || '🙂'
    const systemPrompt = String(customPersonaForm.value.systemPrompt || '').trim().slice(0, 2500)
    if (!name) {
        ElMessage.warning('请填写伙伴名称')
        return
    }
    if (!systemPrompt) {
        ElMessage.warning('请填写性格设定（将作为 System 指令合并进对话）')
        return
    }
    const id = `custom_${Date.now()}`
    chatCustomPersonas.value = [...chatCustomPersonas.value, { id, name, emoji, systemPrompt }]
    chatActivePersonaId.value = id
    customPersonaDialogOpen.value = false
    saveChatPersonaLocalStorage()
    await persistChatPersonaToServer()
    ElMessage.success('已保存自定义伙伴并设为当前对话身份')
}

async function deleteCustomPersona(p) {
    if (!p?.id) {
        return
    }
    try {
        await ElMessageBox.confirm(`删除自定义伙伴「${p.name}」？`, '删除', { type: 'warning' })
    } catch {
        return
    }
    chatCustomPersonas.value = chatCustomPersonas.value.filter((x) => x.id !== p.id)
    if (chatActivePersonaId.value === p.id) {
        chatActivePersonaId.value = 'workplace_helper'
    }
    saveChatPersonaLocalStorage()
    await persistChatPersonaToServer()
    ElMessage.success('已删除')
}

function resolvePersonaPayloadForChat() {
    const id = String(chatActivePersonaId.value || '').trim()
    const builtinIds = new Set(BUILTIN_CHAT_PERSONAS.map((b) => b.id))
    if (builtinIds.has(id)) {
        return { personaId: id }
    }
    if (id.startsWith('custom_')) {
        const custom = chatCustomPersonas.value.find((c) => c.id === id)
        if (custom) {
            return {
                personaId: id,
                personaPrompt: custom.systemPrompt,
                personaLabel: custom.name,
                personaEmoji: custom.emoji,
            }
        }
        return { personaId: 'workplace_helper' }
    }
    return { personaId: 'workplace_helper' }
}

async function syncDashboardData() {
    await loadDashboardStats()
    await loadMyTemplates()
    ElMessage.success('已刷新统计与模板列表')
}

async function loadDashboardStats() {
    const email = userStore.user?.email
    if (!email) {
        dashboardStats.value = { conversation_turns: 0, papers_processed: 0, skills_count: 0 }
        return
    }
    try {
        const r = await fetch(`/api/stats?email=${encodeURIComponent(email)}`)
        const d = await r.json().catch(() => ({}))
        if (r.ok && d.success) {
            dashboardStats.value = {
                conversation_turns: Number(d.conversation_turns || 0),
                papers_processed: Number(d.papers_processed || 0),
                skills_count: Number(d.skills_count || 0),
            }
        }
    } catch {
        /* ignore */
    }
}

async function saveCompanionPersonaToProfile() {
    const email = userStore.user?.email
    if (!email) {
        ElMessage.warning('请先登录后再保存陪伴设定')
        return
    }
    try {
        const r = await fetch('/api/user/profile', {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                email,
                profile: {
                    role: userImpression.value.role,
                    personality: userImpression.value.personality,
                    preferences: userImpression.value.preferences,
                    current_status: userImpression.value.current_status,
                    emotion_history: userImpression.value.emotion_history,
                    companion_persona: companionSettings.value.persona,
                    companion_address: companionSettings.value.address,
                    companion_custom_traits: companionSettings.value.customTraits,
                },
            }),
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        ElMessage.success('陪伴人格已保存')
    } catch (e) {
        ElMessage.error(String(e?.message || e || '保存失败'))
    }
}

const chatInput = ref('')

/** 随对话请求上报，供后端注入「环境与时刻」（用户本地钟面） */
function buildClientLocalTimeLabel() {
    try {
        const d = new Date()
        const y = d.getFullYear()
        const mo = d.getMonth() + 1
        const day = d.getDate()
        const wd = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'][d.getDay()]
        const h = d.getHours()
        const mi = d.getMinutes()
        const hh = String(h).padStart(2, '0')
        const mm = String(mi).padStart(2, '0')
        let period = '深夜'
        if (h >= 5 && h < 12) {
            period = '上午'
        } else if (h >= 12 && h < 14) {
            period = '中午'
        } else if (h >= 14 && h < 18) {
            period = '下午'
        } else if (h >= 18 && h < 23) {
            period = '晚上'
        } else if (h >= 23 || h < 5) {
            period = h >= 23 ? '深夜' : '凌晨'
        }
        return `当前用户本地时间：${y}年${mo}月${day}日 ${wd} ${period} ${hh}:${mm}`
    } catch {
        return ''
    }
}

function getSpeechRecognitionCtor() {
    if (typeof window === 'undefined') {
        return null
    }
    return window.SpeechRecognition || window.webkitSpeechRecognition || null
}

const speechSupported = ref(false)
const speechListening = ref(false)
/** @type {import('vue').Ref<'chat' | 'companion' | null>} */
const speechTarget = ref(null)
let speechRecognitionInstance = null
let speechPrefixSnapshot = ''
let speechFinalAccum = ''
let speechUnexpectedRetries = 0
let speechEndRetryTimer = null
let speechUserStopping = false

function stopSpeechInput() {
    speechUserStopping = true
    speechListening.value = false
    speechTarget.value = null
    if (speechEndRetryTimer) {
        clearTimeout(speechEndRetryTimer)
        speechEndRetryTimer = null
    }
    try {
        speechRecognitionInstance?.stop?.()
    } catch {
        /* ignore */
    }
    speechRecognitionInstance = null
    speechUnexpectedRetries = 0
}

function toggleSpeechInput(mode) {
    const Ctor = getSpeechRecognitionCtor()
    if (!Ctor) {
        return
    }
    if (speechListening.value && speechTarget.value === mode) {
        stopSpeechInput()
        return
    }
    if (speechListening.value) {
        stopSpeechInput()
    }

    speechUserStopping = false
    speechPrefixSnapshot = mode === 'companion' ? companionInput.value : chatInput.value
    speechFinalAccum = ''
    speechUnexpectedRetries = 0

    const rec = new Ctor()
    speechRecognitionInstance = rec
    rec.continuous = true
    rec.interimResults = true
    rec.lang = 'zh-CN'

    speechTarget.value = mode
    speechListening.value = true

    rec.onresult = (event) => {
        let interim = ''
        for (let i = event.resultIndex; i < event.results.length; i += 1) {
            const r = event.results[i]
            const piece = r[0]?.transcript ?? ''
            if (r.isFinal) {
                speechFinalAccum += piece
            } else {
                interim += piece
            }
        }
        const merged = speechPrefixSnapshot + speechFinalAccum + interim
        if (mode === 'companion') {
            companionInput.value = merged
        } else {
            chatInput.value = merged
        }
    }

    rec.onerror = (event) => {
        const err = String(event?.error || '').trim()
        if (err === 'aborted' || err === 'no-speech') {
            return
        }
        /** 网络或服务端识别不可用（不向用户展示底层英文 error 文案） */
        const networkOrService = err === 'network' || err === 'service-not-allowed'
        if (networkOrService) {
            ElMessage.warning('当前浏览器语音服务网络波动，建议使用文字输入哦～')
            stopSpeechInput()
            return
        }
        if (err === 'not-allowed') {
            ElMessage.warning('无法使用麦克风，请检查浏览器权限（建议在 HTTPS 或 localhost 下使用）')
        } else if (err === 'audio-capture') {
            ElMessage.warning('未检测到可用麦克风，请检查设备或权限')
        } else if (err === 'language-not-supported') {
            ElMessage.warning('当前环境暂不支持该语言的语音识别，建议使用文字输入哦～')
        } else {
            console.warn('[speech]', err)
            ElMessage.warning('语音输入暂时不可用，建议使用文字输入哦～')
        }
        stopSpeechInput()
    }

    rec.onend = () => {
        if (speechUserStopping || !speechListening.value || speechRecognitionInstance !== rec) {
            return
        }
        if (speechUnexpectedRetries >= 2) {
            ElMessage.info('语音识别已暂停；环境嘈杂时可改用键盘输入')
            stopSpeechInput()
            return
        }
        speechUnexpectedRetries += 1
        speechEndRetryTimer = setTimeout(() => {
            try {
                if (
                    speechListening.value &&
                    speechTarget.value === mode &&
                    speechRecognitionInstance === rec &&
                    !speechUserStopping
                ) {
                    rec.start()
                }
            } catch {
                stopSpeechInput()
            }
        }, 420)
    }

    try {
        rec.start()
    } catch (e) {
        console.warn('[speech] start', e)
        ElMessage.error('无法启动语音识别，请检查浏览器是否允许麦克风')
        stopSpeechInput()
    }
}

/** PPT */
const pptTopic = ref('')
const pptPages = ref(10)
const pptStyle = ref('商务专业风')
const pptOutlineMd = ref('')
const pptGenLoading = ref(false)
const pptExportLoading = ref(false)

/** 待办：仅顶栏铃铛 + 抽屉展示，不在对话流插入卡片 */
const upcomingSoon = ref([])
const upcomingOverdue = ref([])
const todoPendingCount = ref(0)
const todoDrawerVisible = ref(false)
const todosDrawerList = ref([])
const todosDrawerLoading = ref(false)
const TODO_NOTIFY_STORAGE_KEY = 'openclaw_todo_notified_v1'
let todoPollTimer = null

/** 沉浸页悬浮待办铃铛：拖拽、边界、localStorage 记忆 */
const OPENCLAW_BELL_POSITION_KEY = 'openclaw_bell_position'
const BELL_HIT_SIZE = 44
const BELL_DRAG_THRESHOLD_PX = 6

function clampBellScreenPos(x, y) {
    if (typeof window === 'undefined') {
        return { x: 16, y: 16 }
    }
    const pad = 4
    const maxX = window.innerWidth - BELL_HIT_SIZE - pad
    const maxY = window.innerHeight - BELL_HIT_SIZE - pad
    return {
        x: Math.min(Math.max(pad, x), maxX),
        y: Math.min(Math.max(pad, y), maxY),
    }
}

function defaultBellScreenPos() {
    if (typeof window === 'undefined') {
        return { x: 16, y: 16 }
    }
    return clampBellScreenPos(window.innerWidth - BELL_HIT_SIZE - 24, 16)
}

function loadBellScreenPosFromStorage() {
    if (typeof window === 'undefined' || typeof localStorage === 'undefined') {
        return { x: 16, y: 16 }
    }
    try {
        const raw = localStorage.getItem(OPENCLAW_BELL_POSITION_KEY)
        if (!raw) {
            return defaultBellScreenPos()
        }
        const o = JSON.parse(raw)
        const x = Number(o.x)
        const y = Number(o.y)
        if (!Number.isFinite(x) || !Number.isFinite(y)) {
            return defaultBellScreenPos()
        }
        return clampBellScreenPos(x, y)
    } catch {
        return defaultBellScreenPos()
    }
}

function persistBellScreenPos() {
    if (typeof localStorage === 'undefined') {
        return
    }
    try {
        localStorage.setItem(
            OPENCLAW_BELL_POSITION_KEY,
            JSON.stringify({ x: bellScreenPos.value.x, y: bellScreenPos.value.y }),
        )
    } catch {
        /* ignore */
    }
}

const bellScreenPos = ref(loadBellScreenPosFromStorage())
const bellDragging = ref(false)

const bellPointerMoveOpts = { passive: false, capture: false }

let bellDragSession = {
    startClientX: 0,
    startClientY: 0,
    originX: 0,
    originY: 0,
    movedPastThreshold: false,
    listenersActive: false,
}

function teardownBellDragListeners() {
    if (!bellDragSession.listenersActive) {
        return
    }
    document.removeEventListener('pointermove', onBellPointerMove, bellPointerMoveOpts)
    document.removeEventListener('pointerup', onBellPointerUp, true)
    document.removeEventListener('pointercancel', onBellPointerUp, true)
    bellDragSession.listenersActive = false
}

function onBellPointerMove(ev) {
    const dx = ev.clientX - bellDragSession.startClientX
    const dy = ev.clientY - bellDragSession.startClientY
    if (Math.hypot(dx, dy) > BELL_DRAG_THRESHOLD_PX) {
        bellDragSession.movedPastThreshold = true
        bellDragging.value = true
    }
    if (!bellDragSession.movedPastThreshold) {
        return
    }
    bellScreenPos.value = clampBellScreenPos(bellDragSession.originX + dx, bellDragSession.originY + dy)
    ev.preventDefault()
}

function onBellPointerUp() {
    teardownBellDragListeners()
    const wasDrag = bellDragSession.movedPastThreshold
    bellDragging.value = false
    bellDragSession.movedPastThreshold = false
    if (wasDrag) {
        persistBellScreenPos()
    } else {
        openTodoDrawer()
    }
}

function onBellPointerDown(ev) {
    if (typeof ev.button === 'number' && ev.button !== 0) {
        return
    }
    teardownBellDragListeners()
    bellDragSession.startClientX = ev.clientX
    bellDragSession.startClientY = ev.clientY
    bellDragSession.originX = bellScreenPos.value.x
    bellDragSession.originY = bellScreenPos.value.y
    bellDragSession.movedPastThreshold = false
    bellDragging.value = false
    document.addEventListener('pointermove', onBellPointerMove, bellPointerMoveOpts)
    document.addEventListener('pointerup', onBellPointerUp, true)
    document.addEventListener('pointercancel', onBellPointerUp, true)
    bellDragSession.listenersActive = true
    ev.preventDefault()
}

function onBellWindowResize() {
    bellScreenPos.value = clampBellScreenPos(bellScreenPos.value.x, bellScreenPos.value.y)
}

/** 对话会话（左侧历史） */
const chatSessions = ref([])
const chatSessionsLoading = ref(false)
const currentChatSessionId = ref(null)
/** 窄屏时历史栏以抽屉覆盖显示，由汉堡按钮切换 */
const chatHistoryOverlayOpen = ref(false)
const chatInputTextareaRef = ref(null)

/** 我的 Word 模板库 */
const myTemplates = ref([])
const myTemplatesLoading = ref(false)

const modelConfigs = ref([])
const modelConfigsLoading = ref(false)
const modelConfigsError = ref('')
const modelModalOpen = ref(false)
const modelFormSaving = ref(false)
const modelTestLoading = ref(false)
const modelTestMessage = ref('')
/** null=未测，true=已通过，false=未通过 */
const modelTestPassed = ref(null)
const editingModelId = ref(null)
const modelForm = ref({
    name: '',
    type: 'aliyun',
    base_url: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
    model_name: 'qwen-plus',
    api_key: '',
    priority: 10,
    is_enabled: true,
})

/** 语音识别 Whisper：与对话模型分离，仅存浏览器 localStorage，随会议上传带给后端 */
const WHISPER_ASR_STORAGE_KEY = 'openclaw_whisper_asr_v1'
const DEFAULT_WHISPER_ASR_BASE = 'https://api.openai.com/v1'
const whisperAsrApiKey = ref('')
const whisperAsrBaseUrl = ref(DEFAULT_WHISPER_ASR_BASE)

function loadWhisperAsrFromStorage() {
    if (typeof localStorage === 'undefined') {
        return
    }
    try {
        const raw = localStorage.getItem(WHISPER_ASR_STORAGE_KEY)
        if (!raw) {
            whisperAsrApiKey.value = ''
            whisperAsrBaseUrl.value = DEFAULT_WHISPER_ASR_BASE
            return
        }
        const o = JSON.parse(raw)
        whisperAsrApiKey.value = String(o.whisperApiKey ?? '')
        whisperAsrBaseUrl.value = String(o.whisperBaseUrl || DEFAULT_WHISPER_ASR_BASE).trim() || DEFAULT_WHISPER_ASR_BASE
    } catch {
        whisperAsrApiKey.value = ''
        whisperAsrBaseUrl.value = DEFAULT_WHISPER_ASR_BASE
    }
}

function saveWhisperAsrToStorage() {
    if (typeof localStorage === 'undefined') {
        ElMessage.warning('当前环境无法写入本地存储')
        return
    }
    const payload = {
        whisperApiKey: String(whisperAsrApiKey.value || '').trim(),
        whisperBaseUrl: String(whisperAsrBaseUrl.value || '').trim() || DEFAULT_WHISPER_ASR_BASE,
    }
    localStorage.setItem(WHISPER_ASR_STORAGE_KEY, JSON.stringify(payload))
    ElMessage.success('语音识别 (ASR) 配置已保存到本机，会议转写时会自动随请求发送到服务器。')
}

function getInitialChatMessages() {
    return [
        {
            role: 'assistant',
            isHtml: true,
            content:
                '你好！我是 OpenClaw 对话助手。我可以帮你：<br><br>' +
                '&#8226; <strong>直接操作工作台</strong> — 比如"帮我创建会议纪要"、"打开模板引擎"<br>' +
                '&#8226; <strong>智能问答</strong> — 回答关于功能使用的问题<br>' +
                '&#8226; <strong>快捷跳转</strong> — 说"去API配置"即可跳转到配置页<br><br>' +
                '试试输入：<em>"帮我生成一个PPT大纲"</em> 或 <em>"配置OpenAI的API"</em>',
        },
    ]
}

const chatMessages = ref(getInitialChatMessages())

/** 对话助手：本轮/最近注入的文献文件名提示（非服务端多轮 Session，仅 UI 提示） */
const chatLoadedPaperLabel = ref('')
/** 与后端 MAX_PAPER_CONTEXT_BODY_CHARS 一致，控制单次请求 system 体积 */
const CHAT_INJECT_FULLTEXT_MAX = 5000
const CHAT_INJECT_SUMMARY_JSON_MAX = 8000

/** Agent 印象面板（持久化至 user_profile）；情绪曲线数据仅来自服务端 emotion_history */
const DEFAULT_USER_IMPRESSION = {
    role: '普通用户',
    personality: '待观察',
    preferences: '',
    current_status: '',
    emotion_history: [],
}
const userImpression = ref({ ...DEFAULT_USER_IMPRESSION })
const impressionLoading = ref(false)
const impressionSaving = ref(false)
/** 「眼中的你」侧栏改为抽屉，不占对话主区域 */
const impressionDrawerVisible = ref(false)
const editingImpressionField = ref(null)
const IMPRESSION_FIELD_META = [
    { key: 'role', label: '角色', placeholder: '如：科研工作者、产品经理…' },
    { key: 'personality', label: '性格倾向', placeholder: '如：理性直率、偏细节控…' },
    { key: 'preferences', label: '工作偏好 / 沟通风格', placeholder: '如：喜欢要点列表、少寒暄…' },
    { key: 'current_status', label: '近期状态 / 焦虑点', placeholder: '如：赶 deadline、担心答辩…' },
]

watch(
    () => [route.name, route.query.tab, route.query.skillId, route.query.wake],
    async () => {
        if (!['dashboard', 'dashboard-chat', 'dashboard-companion'].includes(String(route.name))) {
            return
        }
        if (route.name === 'dashboard-chat') {
            currentTab.value = 'chat'
        } else if (route.name === 'dashboard-companion') {
            currentTab.value = 'companion'
        } else {
            const t = route.query.tab
            if (typeof t === 'string' && tabTitles[t]) {
                currentTab.value = t
            } else {
                currentTab.value = 'overview'
            }
        }
        const sid = route.query.skillId
        if (sid == null || String(sid).trim() === '') {
            return
        }
        const id = Number.parseInt(String(sid), 10)
        if (!Number.isFinite(id)) {
            return
        }
        const email = userStore.user?.email
        if (email && savedSkills.value.length === 0) {
            await fetchSavedSkills()
        }
        const sk = savedSkills.value.find((s) => s.id === id)
        if (sk) {
            activeChatSkill.value = { id: sk.id, name: sk.name }
            const w = String(route.query.wake || '').trim()
            if (w) {
                chatInput.value = w
            }
        }
        router.replace({ query: { ...route.query, skillId: undefined, wake: undefined } })
    },
    { immediate: true },
)

watch(
    () => userStore.user?.email,
    () => {
        void loadDashboardStats()
        void loadTodosUpcoming()
        void loadMyTemplates()
        void loadChatSessions()
        void loadTodoPendingCount()
    },
)

/** 自定义 Skill：打包提炼、@ 唤醒、随请求注入 system */
const savedSkills = ref([])
const activeChatSkill = ref(null)
const skillExtractLoading = ref(false)
const skillPickerOpen = ref(false)
const skillPickerQuery = ref('')
const filteredSkills = computed(() => {
    const list = savedSkills.value
    const q = skillPickerQuery.value.trim().toLowerCase()
    if (!q) {
        return list
    }
    return list.filter(
        (s) =>
            String(s.name || '')
                .toLowerCase()
                .includes(q) ||
            String(s.description || '')
                .toLowerCase()
                .includes(q),
    )
})
const canPackSkill = computed(() => {
    const n = chatMessages.value.filter(
        (m) =>
            !m.hidden &&
            (m.role === 'user' || m.role === 'assistant') &&
            String(m.content || '').trim().length > 0,
    ).length
    return n >= 2
})

/** 自动画像建模：侧栏微弱动效 */
const impressionUpdating = ref(false)
/** 输入框上方智能推荐 */
const suggestedSkill = ref(null)
let skillSuggestTimer = null
let skillSuggestDismissId = null

const CHAT_MODE_COMPANION = 'companion'
const COMPANION_OPENING = '嘿！今天心情怎么样？不管好坏，我都在这儿陪你扣字呢～ (๑•̀ㅂ•́)و✧'

const companionInput = ref('')
const companionMessages = ref([{ role: 'assistant', content: COMPANION_OPENING }])

const companionMoodChips = [
    { label: 'QAQ 想被安慰', text: '…今天真的有点绷不住，就想找你说说，能安慰我吗 QAQ' },
    { label: '太燃了求分享', text: '我刚遇到一件超燃的事！！很想马上跟你分享一下～' },
    { label: '单纯想吐槽', text: '纯吐槽模式开启：有件事憋在心里好久了，听我叨叨会儿行嘛' },
]

/** —— 文献论文摘要 —— */
const paperUrl = ref('')
const paperTitle = ref('')
const paperRawText = ref('')
const paperPasteBuffer = ref('')
const paperCharCount = ref(0)
const paperSourceTruncated = ref(false)
const paperFileKind = ref('none')
const paperPdfObjectUrl = ref('')
const paperLocalFile = ref(null)
const paperFileInputRef = ref(null)
const paperDragging = ref(false)
const paperParseLoading = ref(false)
/** 生成摘要前：FormData 上传并解析（先于「研读」提示） */
const paperUploadParsing = ref(false)
const paperFetchLoading = ref(false)
const paperSummarizeLoading = ref(false)
/** 解析成功后调用摘要 API 阶段 */
const paperAiReading = ref(false)
const paperError = ref('')
const paperSummary = ref(null)
const PAPER_CHAT_ATTACH_MAX = 14000
const PAPER_PREVIEW_CHARS = 12000

const paperPreviewLeft = computed(() => {
    const t = (paperRawText.value || '').trim()
    if (!t) {
        return '暂无可预览正文。请上传 PDF/Docx 并点击「解析文档」、粘贴链接后「拉取网页」，或在下方摘录区编辑后「合并到正文」。'
    }
    if (t.length > PAPER_PREVIEW_CHARS) {
        return `${t.slice(0, PAPER_PREVIEW_CHARS)}\n\n…（预览已截断，完整文本已保留并用于摘要）`
    }
    return t
})

/** 文献主工作台：已有本地文件或任意正文（链接拉取 / 解析 / 粘贴）后再展示双栏预览与笔记 */
const paperWorkbenchHasSource = computed(
    () => !!paperLocalFile.value || !!(paperRawText.value || '').trim(),
)

/** 动态维度卡片数据（兼容旧版五字段结构） */
const paperSummaryDimensions = computed(() => {
    const s = paperSummary.value
    if (!s || typeof s !== 'object') {
        return {}
    }
    const dim = s.dimensions
    if (dim && typeof dim === 'object' && !Array.isArray(dim)) {
        return dim
    }
    const legacy = {}
    if (s.research_question) {
        legacy['研究问题'] = String(s.research_question)
    }
    if (s.innovation) {
        legacy['方法创新'] = String(s.innovation)
    }
    if (s.results) {
        legacy['实验结果'] = String(s.results)
    }
    if (s.limitations) {
        legacy['局限性'] = String(s.limitations)
    }
    if (Array.isArray(s.citations) && s.citations.length) {
        legacy['可引观点'] = s.citations.map((x) => String(x || '').trim()).filter(Boolean).join('\n')
    }
    return legacy
})

const paperSummaryContentType = computed(() => {
    const s = paperSummary.value
    if (!s || typeof s !== 'object') {
        return ''
    }
    return String(s.content_type || '').trim() || '未分类'
})

const paperSummaryToneClass = computed(() => {
    const t = paperSummaryContentType.value
    if (/论文|学术|期刊|预印本|arxiv/i.test(t)) {
        return 'paper-tone--academic'
    }
    if (/新闻|报道|快讯|消息|媒体/i.test(t)) {
        return 'paper-tone--news'
    }
    if (/会议|纪要|座谈|复盘|研讨/i.test(t)) {
        return 'paper-tone--meeting'
    }
    if (/产品|手册|说明书|规格|白皮书|需求/i.test(t)) {
        return 'paper-tone--product'
    }
    if (/法律|法规|合同|条款/i.test(t)) {
        return 'paper-tone--legal'
    }
    return 'paper-tone--default'
})

function paperDimensionTitleClass(key) {
    const len = String(key || '').length
    return {
        'nb-title--long': len > 14,
        'nb-title--mid': len > 9 && len <= 14,
    }
}

/** 兼容空 Key、过长 Key 与异常字符，避免 v-for 崩溃或布局撑破 */
function sanitizePaperDimensionKey(k, index = 0) {
    const s = String(k == null ? '' : k).replace(/\s+/g, ' ').trim()
    if (!s) {
        return `要点 ${index + 1}`
    }
    return s.length > 80 ? `${s.slice(0, 80)}…` : s
}

const paperSummaryCardEntries = computed(() => {
    const dim = paperSummaryDimensions.value
    return Object.entries(dim)
        .filter(([, v]) => String(v ?? '').trim() !== '')
        .map(([k, v], idx) => {
            const displayKey = sanitizePaperDimensionKey(k, idx)
            const raw = String(k ?? '')
            const vueKey = raw ? `${idx}-${encodeURIComponent(raw).slice(0, 160)}` : `idx-${idx}`
            return {
                displayKey,
                value: String(v ?? '').trim(),
                vueKey,
            }
        })
})

/** 供后端合并进单条 system 的附加块（JSON，便于模型解析） */
function buildPaperSystemContextBlock() {
    const ct = paperSummaryContentType.value || '未分类'
    const dim = paperSummaryDimensions.value
    const payload = { content_type: ct, ...dim }
    return `【OpenClaw·材料结构化摘要】以下 JSON 为工作台提炼的键值要点，供本轮对话使用。\n${JSON.stringify(payload)}`
}

function getPaperChatFileName() {
    const f = paperLocalFile.value
    if (f && f.name) {
        return String(f.name)
    }
    const t = paperTitle.value.trim()
    return t || '未命名文献'
}

function buildPaperSummaryDataJson() {
    const ct = paperSummaryContentType.value || '未分类'
    const dim = paperSummaryDimensions.value
    const payload = { content_type: ct, dimensions: dim || {} }
    let s = JSON.stringify(payload)
    if (s.length > CHAT_INJECT_SUMMARY_JSON_MAX) {
        s = `${s.slice(0, CHAT_INJECT_SUMMARY_JSON_MAX)}…`
    }
    return s
}

/** 注入到对话列表的隐藏用户消息（评委可见线程内已载入上下文） */
function buildPaperContextInjectionBlock() {
    const fileName = getPaperChatFileName()
    const summaryData = buildPaperSummaryDataJson()
    const fullText = String(paperRawText.value || '').trim().slice(0, CHAT_INJECT_FULLTEXT_MAX)
    return (
        `【上下文注入】：用户刚刚解析了一份文献。\n` +
        `文件名：${fileName}\n` +
        `结构化摘要：${summaryData}\n` +
        `全文内容（前 ${CHAT_INJECT_FULLTEXT_MAX} 字）：${fullText}\n` +
        `请基于以上内容准备好回答用户的问题。`
    )
}

function pushHiddenPaperContextMessage() {
    chatMessages.value.push({
        role: 'user',
        hidden: true,
        content: buildPaperContextInjectionBlock(),
    })
}

function formatPaperSummaryForChat() {
    const ct = paperSummaryContentType.value
    const dim = paperSummaryDimensions.value
    const lines = Object.entries(dim).map(([k, v]) => `【${k}】${String(v || '').trim()}`)
    const head = ct ? `【内容类型】${ct}` : '【内容类型】未分类'
    return [head, ...lines].filter(Boolean).join('\n\n')
}

/** 将当前文献卡片同步到 OpenClaw 全局上下文（Pinia），供其他 Tab / 对话读取 */
function syncLiteratureToOpenclawStore() {
    const s = paperSummary.value
    if (!s || typeof s !== 'object') {
        openclawCtx.clearLiteratureCard()
        return
    }
    const dim = paperSummaryDimensions.value
    const keys = Object.keys(dim || {})
    if (keys.length === 0) {
        openclawCtx.clearLiteratureCard()
        return
    }
    openclawCtx.setLiteratureCard({
        content_type: paperSummaryContentType.value || '未分类',
        dimensions: { ...dim },
        documentTitle: paperTitle.value.trim(),
    })
}

watch([paperSummary, paperTitle], syncLiteratureToOpenclawStore, { deep: true })

/** 指挥 OpenClaw：快捷意图 → 对话助手 + 文献 system 上下文 + 低温严谨参数 */
const PAPER_OPENCLAW_QUICK_COMMANDS = [
    {
        id: 'email',
        label: '基于此写邮件',
        userMessage:
            '请根据系统附加上下文中的材料结构化摘要，帮我写一封专业、简洁的中文邮件：说明背景、核心诉求与下一步行动项；若摘要信息不足，请列出需我补充的要点，勿编造摘要中未出现的事实。',
    },
    {
        id: 'study_plan',
        label: '基于此制定复习计划',
        userMessage:
            '请根据系统附加上下文中的材料结构化摘要，为我制定一份可执行的复习/学习计划：分阶段目标、每日时间建议、自测方式；严格依据摘要内容，不要臆造原文细节。',
    },
    {
        id: 'logic_audit',
        label: '基于此寻找逻辑漏洞',
        userMessage:
            '请根据系统附加上下文中的材料结构化摘要，以学术评阅视角审视：论证链条是否完整、有无偷换概念或过度外推、证据与结论是否匹配；逐条列出「可疑点—理由—可如何核实或补强」。不得捏造摘要未给出的数据或结论。',
    },
]

function openPaperClawQuickCommand(commandId) {
    if (!paperSummary.value || paperSummaryCardEntries.value.length === 0) {
        ElMessage.warning('请先生成结构化卡片')
        return
    }
    const cmd = PAPER_OPENCLAW_QUICK_COMMANDS.find((c) => c.id === commandId)
    if (!cmd) {
        return
    }
    pushHiddenPaperContextMessage()
    chatLoadedPaperLabel.value = getPaperChatFileName()
    const paperFullSlice = String(paperRawText.value || '').trim().slice(0, CHAT_INJECT_FULLTEXT_MAX)
    const systemContext = buildPaperSystemContextBlock().slice(0, CHAT_INJECT_SUMMARY_JSON_MAX)
    goTab('chat')
    nextTick(() => {
        sendChatMessage(cmd.userMessage, {
            systemContext,
            paperContext: paperFullSlice,
            chatContextLabel: getPaperChatFileName(),
            temperatureProfile: 'literature',
        })
    })
}

function goTab(tab) {
    currentTab.value = tab
    if (tab === 'chat') {
        router.push({ name: 'dashboard-chat' })
        return
    }
    if (tab === 'companion') {
        router.push({ name: 'dashboard-companion' })
        return
    }
    router.push({ name: 'dashboard', query: { tab } })
}

watch(currentTab, (tab) => {
    if (tab === 'apiConfig') {
        loadWhisperAsrFromStorage()
        fetchModelConfigs()
    }
    if (tab === 'chat') {
        loadUserImpressionProfile()
        fetchSavedSkills()
        void loadChatSessions()
        nextTick(() => syncChatComposerHeight())
    }
    if (tab === 'templates') {
        void loadMyTemplates()
    }
})

/** 对话区滚动容器（主对话 / 陪伴） */
const chatMessagesScrollRef = ref(null)
const companionMessagesScrollRef = ref(null)
let streamScrollRafId = 0

/**
 * 流式输出时合并到同一帧滚动到底部，避免每 token 强制同步 reflow；
 * 使用 scrollTo + behavior:auto 紧跟内容高度，兼容 SSE 高频 chunk。
 */
function scheduleScrollChatToBottom(isCompanion) {
    if (typeof window === 'undefined') {
        return
    }
    if (streamScrollRafId) {
        cancelAnimationFrame(streamScrollRafId)
    }
    streamScrollRafId = window.requestAnimationFrame(() => {
        streamScrollRafId = 0
        const el = isCompanion ? companionMessagesScrollRef.value : chatMessagesScrollRef.value
        if (!el) {
            return
        }
        const target = el.scrollHeight
        el.scrollTo({ top: target, behavior: 'auto' })
    })
}

function scheduleScrollForStreamMessagesRef(messagesRef) {
    scheduleScrollChatToBottom(messagesRef === companionMessages)
}

function getTodoNotifiedIds() {
    try {
        const raw = localStorage.getItem(TODO_NOTIFY_STORAGE_KEY)
        const a = raw ? JSON.parse(raw) : []
        return new Set(Array.isArray(a) ? a.map((x) => Number(x)).filter((n) => Number.isFinite(n)) : [])
    } catch {
        return new Set()
    }
}

function markTodoNotified(id) {
    const s = getTodoNotifiedIds()
    s.add(Number(id))
    localStorage.setItem(TODO_NOTIFY_STORAGE_KEY, JSON.stringify([...s].slice(-400)))
}

async function loadTodosUpcoming() {
    const email = userStore.user?.email
    if (!email) {
        upcomingSoon.value = []
        upcomingOverdue.value = []
        return
    }
    try {
        const r = await fetch(`/api/todos/upcoming?email=${encodeURIComponent(email)}&hours=72`)
        const d = await r.json().catch(() => ({}))
        if (r.ok && d.success) {
            upcomingSoon.value = d.items || []
            upcomingOverdue.value = d.overdue || []
        }
    } catch {
        /* ignore */
    }
    void loadTodoPendingCount()
}

async function loadTodoPendingCount() {
    const email = userStore.user?.email
    if (!email) {
        todoPendingCount.value = 0
        return
    }
    try {
        const r = await fetch(`/api/todos?email=${encodeURIComponent(email)}`)
        const d = await r.json().catch(() => ({}))
        if (r.ok && d.success) {
            const items = d.items || []
            todoPendingCount.value = items.filter((x) => Number(x.done) === 0).length
        }
    } catch {
        todoPendingCount.value = 0
    }
}

async function loadTodosForDrawer() {
    const email = userStore.user?.email
    if (!email) {
        todosDrawerList.value = []
        return
    }
    todosDrawerLoading.value = true
    try {
        const r = await fetch(`/api/todos?email=${encodeURIComponent(email)}`)
        const d = await r.json().catch(() => ({}))
        if (r.ok && d.success) {
            todosDrawerList.value = d.items || []
        }
    } catch {
        todosDrawerList.value = []
    } finally {
        todosDrawerLoading.value = false
    }
}

async function markTodoDoneRow(row) {
    const email = userStore.user?.email
    if (!email || !row?.id) {
        return
    }
    try {
        const r = await fetch(`/api/todos/${row.id}/done`, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email }),
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        void loadTodosForDrawer()
        void loadTodoPendingCount()
        void loadTodosUpcoming()
    } catch (e) {
        ElMessage.error(String(e?.message || e || '操作失败'))
    }
}

async function deleteTodoRow(row) {
    const email = userStore.user?.email
    if (!email || !row?.id) {
        return
    }
    try {
        const r = await fetch(`/api/todos/${row.id}?email=${encodeURIComponent(email)}`, { method: 'DELETE' })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        void loadTodosForDrawer()
        void loadTodoPendingCount()
        void loadTodosUpcoming()
    } catch (e) {
        ElMessage.error(String(e?.message || e || '删除失败'))
    }
}

function openTodoDrawer() {
    todoDrawerVisible.value = true
    void loadTodosForDrawer()
}

function checkDueTodoNotifications() {
    if (typeof window === 'undefined' || !('Notification' in window)) {
        return
    }
    if (Notification.permission !== 'granted') {
        return
    }
    const now = Date.now()
    const notified = getTodoNotifiedIds()
    for (const t of [...upcomingSoon.value, ...upcomingOverdue.value]) {
        const due = Date.parse(t.remindAt)
        if (!Number.isFinite(due) || due > now) {
            continue
        }
        if (notified.has(t.id)) {
            continue
        }
        try {
            new Notification('OpenClaw 待办提醒', { body: t.title })
            markTodoNotified(t.id)
        } catch {
            /* ignore */
        }
    }
}

async function loadMyTemplates() {
    const email = userStore.user?.email
    if (!email) {
        myTemplates.value = []
        return
    }
    myTemplatesLoading.value = true
    try {
        const r = await fetch(`/api/templates/my?email=${encodeURIComponent(email)}`)
        const d = await r.json().catch(() => ({}))
        if (r.ok && d.success) {
            myTemplates.value = d.items || []
        }
    } catch {
        myTemplates.value = []
    } finally {
        myTemplatesLoading.value = false
    }
}

async function deleteMyTemplate(row) {
    const email = userStore.user?.email
    if (!email || !row?.id) {
        return
    }
    try {
        await ElMessageBox.confirm(`确定删除模板「${row.originalName || row.original_name || ''}」？`, '删除模板', {
            type: 'warning',
        })
    } catch {
        return
    }
    try {
        const r = await fetch(`/api/templates/${row.id}?email=${encodeURIComponent(email)}`, { method: 'DELETE' })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        ElMessage.success('已删除')
        if (currentTemplateId.value === row.id) {
            resetTemplateFile()
        }
        void loadMyTemplates()
    } catch (e) {
        ElMessage.error(String(e?.message || e || '删除失败'))
    }
}

async function renameMyTemplate(row) {
    const email = userStore.user?.email
    if (!email || !row?.id) {
        return
    }
    try {
        const { value } = await ElMessageBox.prompt('新的显示名称', '重命名模板', {
            inputValue: row.originalName || row.original_name || '',
            inputPattern: /\S+/,
            inputErrorMessage: '名称不能为空',
        })
        const name = String(value || '').trim().slice(0, 240)
        const r = await fetch(`/api/templates/${row.id}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email, original_name: name }),
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        ElMessage.success('已更新名称')
        void loadMyTemplates()
    } catch (e) {
        if (e !== 'cancel') {
            ElMessage.error(String(e?.message || e || '重命名失败'))
        }
    }
}

async function openMyTemplate(row) {
    const email = userStore.user?.email
    if (!email || !row?.id) {
        ElMessage.warning('请先登录')
        return
    }
    previewError.value = ''
    isPreviewLoading.value = true
    try {
        const r = await fetch(`/api/templates/${row.id}/file?email=${encodeURIComponent(email)}`)
        if (!r.ok) {
            const err = await r.json().catch(() => ({}))
            throw new Error(err.message || `HTTP ${r.status}`)
        }
        const blob = await r.blob()
        const fname = row.originalName || row.original_name || 'template.docx'
        const file = new File([blob], fname, {
            type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        })
        currentFile.value = file
        await nextTick()
        if (!wordContainer.value) {
            throw new Error('预览容器未就绪')
        }
        wordContainer.value.innerHTML = ''
        extractedVariables.value = []
        templateArrayBuffer.value = null
        currentTemplateId.value = row.id
        const ab = await blob.arrayBuffer()
        templateArrayBuffer.value = ab
        await renderDocxPreview(ab)
        const plainText = wordContainer.value?.innerText || ''
        syncVariableList(extractVariablesFromText(plainText))
        ElMessage.success('模板已载入预览')
    } catch (e) {
        console.error(e)
        previewError.value = String(e?.message || e || '载入失败')
        ElMessage.error(previewError.value)
    } finally {
        isPreviewLoading.value = false
    }
}

function onMeetingProcessed() {
    void loadTodosUpcoming()
    void loadTodoPendingCount()
}

async function generatePptOutlineAction() {
    const topic = pptTopic.value.trim()
    if (!topic) {
        ElMessage.warning('请填写 PPT 主题')
        return
    }
    pptGenLoading.value = true
    try {
        const r = await fetch('/api/ppt/generate-outline', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                topic,
                pages: pptPages.value,
                style: pptStyle.value,
            }),
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        pptOutlineMd.value = d.markdown || ''
        ElMessage.success('提纲已生成，可继续编辑后导出 PPTX')
    } catch (e) {
        ElMessage.error(String(e?.message || e || '生成失败'))
    } finally {
        pptGenLoading.value = false
    }
}

async function exportPptFile() {
    if (!pptOutlineMd.value.trim()) {
        ElMessage.warning('请先生成或粘贴 Markdown 提纲')
        return
    }
    pptExportLoading.value = true
    try {
        const r = await fetch('/api/ppt/export', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                markdown: pptOutlineMd.value,
                title: pptTopic.value.trim() || '演示文稿',
            }),
        })
        if (!r.ok) {
            const err = await r.json().catch(() => ({}))
            throw new Error(err.message || `HTTP ${r.status}`)
        }
        const blob = await r.blob()
        const url = URL.createObjectURL(blob)
        const a = document.createElement('a')
        a.href = url
        const base = (pptTopic.value.trim() || 'outline').replace(/[^\w\u4e00-\u9fff\-]+/g, '_') || 'outline'
        a.download = `${base}.pptx`
        a.rel = 'noopener'
        document.body.appendChild(a)
        a.click()
        a.remove()
        URL.revokeObjectURL(url)
        ElMessage.success('已开始下载 PPTX')
    } catch (e) {
        ElMessage.error(String(e?.message || e || '导出失败'))
    } finally {
        pptExportLoading.value = false
    }
}

async function requestTodoNotificationPermission() {
    if (typeof window === 'undefined' || !('Notification' in window)) {
        return
    }
    if (Notification.permission === 'default') {
        try {
            await Notification.requestPermission()
        } catch {
            /* ignore */
        }
    }
}

onMounted(() => {
    if (typeof window !== 'undefined') {
        speechSupported.value = !!getSpeechRecognitionCtor()
        bellScreenPos.value = loadBellScreenPosFromStorage()
        window.addEventListener('resize', onBellWindowResize, { passive: true })
        loadWhisperAsrFromStorage()
        loadChatPersonaLocalStorage()
    }
    fetchSavedSkills()
    void loadDashboardStats()
    void loadUserImpressionProfile()
    void loadTodosUpcoming()
    void loadTodoPendingCount()
    void loadChatSessions()
    void requestTodoNotificationPermission()
    todoPollTimer = window.setInterval(() => {
        void loadTodosUpcoming()
            .then(() => checkDueTodoNotifications())
            .then(() => loadTodoPendingCount())
    }, 45000)
    nextTick(() => {
        void maybeOpenWelcomeWizard()
    })
})

function showToast(type, msg) {
  console.log('[toast]', type, msg)
}

function handleLogout() {
  userStore.clearUser()
  router.push('/login')
}

function clearChat() {
    currentChatSessionId.value = null
    chatMessages.value = getInitialChatMessages()
    chatInput.value = ''
    chatLoadedPaperLabel.value = ''
    activeChatSkill.value = null
    nextTick(() => {
        syncChatComposerHeight()
        scheduleScrollChatToBottom(false)
    })
    showToast('info', '已清空当前草稿（未删除侧栏历史）')
}

async function fetchSavedSkills() {
    const email = userStore.user?.email
    if (!email) {
        savedSkills.value = []
        return
    }
    try {
        const r = await fetch(`/api/skills?email=${encodeURIComponent(email)}`)
        const data = await r.json()
        if (data.success && Array.isArray(data.items)) {
            savedSkills.value = data.items
        }
    } catch (e) {
        console.warn('[skills] list failed', e)
    }
}

function handleChatInputSkill(evt) {
    if (currentTab.value !== 'chat') {
        return
    }
    const el = evt.target
    const v = typeof el.value === 'string' ? el.value : String(chatInput.value || '')
    const pos = typeof el.selectionStart === 'number' ? el.selectionStart : v.length
    const before = v.slice(0, pos)
    const at = before.lastIndexOf('@')
    if (at === -1) {
        skillPickerOpen.value = false
        return
    }
    const seg = before.slice(at + 1)
    if (/[\s\n]/.test(seg)) {
        skillPickerOpen.value = false
        return
    }
    skillPickerQuery.value = seg
    skillPickerOpen.value = savedSkills.value.length > 0
}

function closeSkillPicker() {
    skillPickerOpen.value = false
}

function pickSkillFromMenu(skill) {
    activeChatSkill.value = { id: skill.id, name: skill.name }
    skillPickerOpen.value = false
    const el = document.getElementById('chatInput')
    if (el && typeof el.value === 'string') {
        let v = el.value
        const pos = typeof el.selectionStart === 'number' ? el.selectionStart : v.length
        const before = v.slice(0, pos)
        const at = before.lastIndexOf('@')
        if (at >= 0) {
            v = v.slice(0, at) + v.slice(pos)
            el.value = v
            chatInput.value = v
            nextTick(() => {
                try {
                    el.focus()
                    el.selectionStart = el.selectionEnd = at
                } catch {
                    /* ignore */
                }
            })
        }
    }
    ElMessage.success(`已启用 Skill：${skill.name}`)
}

function clearActiveChatSkill() {
    activeChatSkill.value = null
}

async function extractSkillFromConversation() {
    if (!userStore.user?.email) {
        ElMessage.warning('请先登录后再打包 Skill')
        return
    }
    if (!canPackSkill.value) {
        ElMessage.warning('请先进行至少一轮有效对话后再打包')
        return
    }
    await runSkillExtractFromMessages()
}

/** 从当前对话提炼 Skill（供顶栏与 👍 复用） */
async function runSkillExtractFromMessages() {
    if (!userStore.user?.email) {
        ElMessage.warning('请先登录后再打包 Skill')
        return
    }
    if (!canPackSkill.value) {
        ElMessage.warning('请先进行至少一轮有效对话后再打包')
        return
    }
    const msgs = chatMessages.value
        .filter((m) => !m.hidden && (m.role === 'user' || m.role === 'assistant'))
        .map((m) => ({
            role: m.role,
            content: typeof m.content === 'string' ? m.content : '',
        }))
        .filter((m) => m.content.trim().length > 0)
    if (msgs.length < 2) {
        ElMessage.warning('对话内容不足，无法提炼')
        return
    }
    skillExtractLoading.value = true
    const ac = new AbortController()
    const tid = setTimeout(() => ac.abort(), 120000)
    try {
        const r = await fetch('/api/skills/extract', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                messages: msgs,
                userEmail: userStore.user.email,
            }),
            signal: ac.signal,
        })
        let data = {}
        try {
            data = await r.json()
        } catch {
            data = {}
        }
        if (!r.ok || !data.success) {
            throw new Error(data.message || `请求失败 HTTP ${r.status}`)
        }
        ElMessage.success(data.message || 'Skill 已保存')
        await fetchSavedSkills()
    } catch (e) {
        const msg =
            e?.name === 'AbortError'
                ? '提炼超时（120s），请缩短对话记录或稍后再试'
                : String(e?.message || e || '提炼失败')
        console.error('[skills/extract]', e)
        ElMessage.error(msg)
    } finally {
        try {
            clearTimeout(tid)
        } catch {
            /* ignore */
        }
        skillExtractLoading.value = false
    }
}

/** 对话助手 AI 气泡：仅正常完成流式输出后展示 👍 */
function canShowAssistantLikeBtn(message) {
    if (!message || message.role !== 'assistant' || message.isHtml) {
        return false
    }
    if (message.clarificationCard && !message.clarificationCard.resolved) {
        return false
    }
    if (message.awaitingContent) {
        return false
    }
    if (message.streamFailed) {
        return false
    }
    return String(message.content || '').trim().length > 0
}

function onAssistantThumbsUp() {
    ElMessage.success('感谢认可！')
    void runSkillExtractFromMessages()
}

const welcomeWizardOpen = ref(false)
const welcomeForm = ref({
    identity: '',
    worry: '',
    tone: '',
})

function onboardingStorageKey() {
    const em = userStore.user?.email
    return em ? `openclaw_onboarding_v1_${em}` : ''
}

async function maybeOpenWelcomeWizard() {
    const email = userStore.user?.email
    if (!email) {
        return
    }
    if (typeof localStorage !== 'undefined') {
        const k = onboardingStorageKey()
        if (k && localStorage.getItem(k) === '1') {
            return
        }
    }
    await loadUserImpressionProfile()
    const r = String(userImpression.value.role || '').trim()
    if (r && r !== '待观察') {
        return
    }
    welcomeForm.value = { identity: '', worry: '', tone: '' }
    welcomeWizardOpen.value = true
}

async function submitWelcomeWizard() {
    const email = userStore.user?.email
    if (!email) {
        welcomeWizardOpen.value = false
        return
    }
    const identity = String(welcomeForm.value.identity || '').trim() || '同事'
    const worry = String(welcomeForm.value.worry || '').trim() || '（暂未填写）'
    const tone = String(welcomeForm.value.tone || '').trim() || '自然、简洁'
    try {
        const r = await fetch('/api/user/profile', {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                email,
                profile: {
                    role: identity.slice(0, 500),
                    personality: String(userImpression.value.personality || '待观察').slice(0, 500),
                    preferences: tone.slice(0, 1000),
                    current_status: worry.slice(0, 1000),
                },
            }),
        })
        let data = {}
        try {
            data = await r.json()
        } catch {
            data = {}
        }
        if (!r.ok || !data.success) {
            throw new Error(data.message || `HTTP ${r.status}`)
        }
        if (data.profile) {
            userImpression.value = { ...DEFAULT_USER_IMPRESSION, ...data.profile }
            if (!Array.isArray(userImpression.value.emotion_history)) {
                userImpression.value.emotion_history = []
            }
        }
        if (typeof localStorage !== 'undefined' && onboardingStorageKey()) {
            localStorage.setItem(onboardingStorageKey(), '1')
        }
        welcomeWizardOpen.value = false
        ElMessage.success('已保存，OpenClaw 会更懂你的语境～')
    } catch (e) {
        ElMessage.error(String(e?.message || e || '保存失败'))
    }
}

function tokenizeLoose(s) {
    return String(s || '')
        .toLowerCase()
        .replace(/[^\w\u4e00-\u9fff]+/g, ' ')
        .split(/\s+/)
        .filter((w) => w.length > 1)
}

function scoreSkillMatch(input, skill) {
    const ti = new Set(tokenizeLoose(input))
    const blob = `${skill.name || ''} ${skill.description || ''}`
    const sj = new Set(tokenizeLoose(blob))
    if (ti.size === 0 || sj.size === 0) {
        return 0
    }
    let inter = 0
    for (const x of ti) {
        if (sj.has(x)) {
            inter += 1
        }
    }
    const union = new Set([...ti, ...sj]).size
    let jac = union ? inter / union : 0
    const simpleInput = String(input).toLowerCase()
    for (const w of sj) {
        if (w.length >= 2 && simpleInput.includes(w)) {
            jac += 0.055
        }
    }
    return Math.min(1, jac)
}

function refreshSkillSuggestion() {
    suggestedSkill.value = null
    const q = String(chatInput.value || '').trim()
    if (q.length < 12 || activeChatSkill.value?.id) {
        return
    }
    let best = null
    let bestScore = 0
    for (const s of savedSkills.value) {
        const sc = scoreSkillMatch(q, s)
        if (sc > bestScore) {
            bestScore = sc
            best = s
        }
    }
    if (skillSuggestDismissId != null && best?.id === skillSuggestDismissId) {
        return
    }
    if (best && bestScore >= 0.24 && bestScore < 0.92) {
        suggestedSkill.value = { skill: best, score: bestScore }
    }
}

function dismissSkillSuggestion() {
    if (suggestedSkill.value?.skill?.id != null) {
        skillSuggestDismissId = suggestedSkill.value.skill.id
    }
    suggestedSkill.value = null
}

function applySuggestedSkill() {
    const sk = suggestedSkill.value?.skill
    if (!sk) {
        return
    }
    activeChatSkill.value = { id: sk.id, name: sk.name }
    suggestedSkill.value = null
    ElMessage.success(`已启用技能：${sk.name}`)
}

watch(chatInput, () => {
    clearTimeout(skillSuggestTimer)
    skillSuggestTimer = setTimeout(refreshSkillSuggestion, 420)
    nextTick(() => syncChatComposerHeight())
})

async function loadUserImpressionProfile() {
    const email = userStore.user?.email
    if (!email) {
        userImpression.value = { ...DEFAULT_USER_IMPRESSION }
        return
    }
    impressionLoading.value = true
    try {
        const r = await fetch(`/api/user/profile?email=${encodeURIComponent(email)}`)
        let data = {}
        try {
            data = await r.json()
        } catch {
            data = {}
        }
        if (r.ok && data.success && data.profile) {
            userImpression.value = { ...DEFAULT_USER_IMPRESSION, ...data.profile }
            if (!Array.isArray(userImpression.value.emotion_history)) {
                userImpression.value.emotion_history = []
            }
            companionSettings.value = {
                persona: data.profile.companion_persona || 'default',
                address: data.profile.companion_address || '',
                customTraits: data.profile.companion_custom_traits || '',
            }
            const srvActive = data.profile.chat_active_persona_id
            if (srvActive && String(srvActive).trim()) {
                chatActivePersonaId.value = String(srvActive).trim().slice(0, 64)
            }
            const srvCustom = data.profile.chat_custom_personas
            if (Array.isArray(srvCustom) && srvCustom.length > 0) {
                chatCustomPersonas.value = sanitizeClientCustomPersonas(srvCustom)
            }
            saveChatPersonaLocalStorage()
            return
        }
        if (!r.ok || r.status === 404 || r.status >= 500) {
            console.warn(
                '[WARN] 个人画像接口尚未就绪，使用本地默认数据',
                'HTTP',
                r.status,
            )
        } else {
            console.warn('[WARN] 个人画像接口尚未就绪，使用本地默认数据', data.message || '')
        }
        userImpression.value = { ...DEFAULT_USER_IMPRESSION }
    } catch (e) {
        console.warn('[WARN] 个人画像接口尚未就绪，使用本地默认数据', e?.message || e)
        userImpression.value = { ...DEFAULT_USER_IMPRESSION }
    } finally {
        impressionLoading.value = false
    }
}

function submitClarificationReply(messageIndex) {
    try {
        const list = chatMessages.value
        const msg = list[messageIndex]
        if (!msg || msg.role !== 'assistant' || !msg.clarificationCard || msg.clarificationCard.resolved) {
            return
        }
        const { missing_info: missingInfo, answers } = msg.clarificationCard
        const lines = []
        for (let i = 0; i < missingInfo.length; i += 1) {
            const q = missingInfo[i]
            const a = String(answers[i] ?? '').trim()
            lines.push(a ? `${q}\n${a}` : `${q}\n（暂未填写）`)
        }
        const payload = `关于你的疑问，我的补充是：\n\n${lines.join('\n\n')}`
        msg.clarificationCard.resolved = true
        list.splice(messageIndex, 1, { ...msg })
        sendChatMessage(payload)
    } catch (e) {
        console.error('[clarification] submit failed', e)
        ElMessage.error('提交失败，请重试')
    }
}

async function saveUserImpressionProfile() {
    const email = userStore.user?.email
    if (!email) {
        ElMessage.warning('请先登录后保存「眼中的你」画像')
        return
    }
    impressionSaving.value = true
    try {
        const r = await fetch('/api/user/profile', {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                email,
                profile: {
                    role: userImpression.value.role,
                    personality: userImpression.value.personality,
                    preferences: userImpression.value.preferences,
                    current_status: userImpression.value.current_status,
                },
            }),
        })
        let data = {}
        try {
            data = await r.json()
        } catch {
            data = {}
        }
        if (!r.ok || !data.success) {
            if (!r.ok && (r.status === 404 || r.status >= 500)) {
                console.warn(
                    '[WARN] 个人画像接口尚未就绪，使用本地默认数据',
                    'HTTP',
                    r.status,
                )
            } else if (!data.success) {
                console.warn('[WARN] 个人画像接口尚未就绪，使用本地默认数据', data.message || '')
            }
            throw new Error(data.message || `HTTP ${r.status}`)
        }
        if (data.profile) {
            userImpression.value = { ...DEFAULT_USER_IMPRESSION, ...data.profile }
            if (!Array.isArray(userImpression.value.emotion_history)) {
                userImpression.value.emotion_history = []
            }
        }
        editingImpressionField.value = null
        ElMessage.success('印象已保存，后续对话将自动注入')
    } catch (e) {
        ElMessage.error(String(e?.message || e || '保存失败'))
    } finally {
        impressionSaving.value = false
    }
}

/** 答辩演示：清空对话 + 服务端画像与缓冲；Skill 库不动 */
async function confirmDemoPresentationReset() {
    const email = userStore.user?.email
    if (!email) {
        ElMessage.warning('请先登录后再使用演示重置')
        return
    }
    try {
        await ElMessageBox.confirm(
            '将清空当前对话与服务器上的「眼中的你」画像缓存（Skill 库保留）。适合答辩反复演示。',
            '演示重置',
            { confirmButtonText: '重置', cancelButtonText: '取消', type: 'warning' },
        )
    } catch {
        return
    }
    try {
        const r = await fetch('/api/user/demo-reset', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email }),
        })
        let data = {}
        try {
            data = await r.json()
        } catch {
            data = {}
        }
        if (!r.ok || !data.success) {
            throw new Error(data.message || `HTTP ${r.status}`)
        }
        chatMessages.value = getInitialChatMessages()
        chatInput.value = ''
        chatLoadedPaperLabel.value = ''
        activeChatSkill.value = null
        currentChatSessionId.value = null
        clearCompanionChat()
        userImpression.value = { ...DEFAULT_USER_IMPRESSION, ...data.profile }
        if (!Array.isArray(userImpression.value.emotion_history)) {
            userImpression.value.emotion_history = []
        }
        ElMessage.success('演示状态已重置')
    } catch (e) {
        ElMessage.error(String(e?.message || e || '重置失败'))
    }
}

function startEditImpression(key) {
    editingImpressionField.value = key
}

function blurImpressionField() {
    editingImpressionField.value = null
}

function sendQuickChat(text) {
    sendChatMessage(text)
}

function onChatComposerKeydown(e) {
    if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault()
        sendChatMessage()
    }
}

function tryAutoNavigate(reply) {
    if (reply.includes('PPT')) {
        goTab('ppt')
        return
    }

    if (reply.includes('会议')) {
        goTab('meeting')
        return
    }

    if (reply.includes('模板')) {
        goTab('templates')
    }
}

/** 从单个 SSE 帧（事件块以空行分隔）中拼接全部 data: 行内容 */
function collectSseDataFromPart(part) {
    const lines = part.split('\n')
    const pieces = []
    for (const line of lines) {
        const m = line.match(/^\s*data:\s*(.*)$/i)
        if (m) {
            pieces.push(m[1])
        }
    }
    return pieces.join('\n').trim()
}

/**
 * 解析 SSE data 负载：区分 streamDone / chunkType(thought|delta) / text。
 * 所有 JSON.parse 均包在 try-catch 内，避免半截 chunk 抛错。
 */
function parseChatStreamPayload(raw) {
    let s = String(raw ?? '').trim()
    s = s.replace(/^\s*data:\s*/i, '').trim()
    if (!s || s === '[DONE]') {
        return { chunkType: null, text: '', streamDone: false }
    }

    const tryParseObject = (jsonStr) => {
        try {
            return JSON.parse(jsonStr)
        } catch {
            return null
        }
    }

    if (s.startsWith('{') || s.startsWith('[')) {
        const obj = tryParseObject(s)
        if (obj && typeof obj === 'object' && !Array.isArray(obj)) {
            if (obj.type === 'done') {
                return { chunkType: null, text: '', streamDone: true }
            }
            if (obj.type === 'thought' && typeof obj.text === 'string') {
                return { chunkType: 'thought', text: obj.text, streamDone: false }
            }
            if (obj.type === 'delta' && typeof obj.text === 'string') {
                return { chunkType: 'delta', text: obj.text, streamDone: false }
            }
            if (obj.type === 'agent_status' && typeof obj.text === 'string') {
                return { chunkType: 'agent_status', text: obj.text, streamDone: false }
            }
            if (obj.type === 'agent_nav' && typeof obj.tab === 'string') {
                return { chunkType: 'agent_nav', text: obj.tab, streamDone: false }
            }
            if (obj.type === 'clarification' && Array.isArray(obj.missing_info)) {
                return {
                    chunkType: 'clarification',
                    clarification: {
                        missing_info: obj.missing_info.map((x) => String(x ?? '')),
                    },
                    text: '',
                    streamDone: false,
                }
            }
            if (typeof obj.content === 'string') {
                return { chunkType: 'delta', text: obj.content, streamDone: false }
            }
            try {
                const deltaContent = obj?.choices?.[0]?.delta?.content
                if (typeof deltaContent === 'string') {
                    return { chunkType: 'delta', text: deltaContent, streamDone: false }
                }
                const deltaThought = obj?.choices?.[0]?.delta?.reasoning_content
                if (typeof deltaThought === 'string') {
                    return { chunkType: 'thought', text: deltaThought, streamDone: false }
                }
            } catch {
                return { chunkType: null, text: '', streamDone: false }
            }
        }
        return { chunkType: null, text: '', streamDone: false }
    }

    return { chunkType: 'delta', text: s, streamDone: false }
}

function processChatSseEventPart(part, onChunk) {
    const dataRaw = collectSseDataFromPart(part)
    if (!dataRaw) {
        return false
    }
    let parsed
    try {
        parsed = parseChatStreamPayload(dataRaw)
    } catch {
        return false
    }
    if (parsed.streamDone) {
        return true
    }
    if (parsed.chunkType === 'clarification' && parsed.clarification) {
        onChunk('clarification', '', parsed.clarification)
        return false
    }
    if (parsed.text && parsed.chunkType) {
        onChunk(parsed.chunkType, parsed.text)
    }
    return false
}

/** 对话流：thought 写入 thoughtChain，delta 写入 content；splice 触发视图更新 */
function appendToStreamMessages(assistantMessage, assistantIndex, chunkType, text, messagesRef, meta) {
    if (chunkType === 'clarification' && meta && Array.isArray(meta.missing_info)) {
        const raw = meta.missing_info.map((x) => String(x ?? '').trim()).filter(Boolean)
        const questions = raw.length > 0 ? raw : ['请补充更具体的需求（场景、受众、格式等）。']
        assistantMessage.clarificationCard = {
            missing_info: questions,
            answers: questions.map(() => ''),
            resolved: false,
        }
        assistantMessage.awaitingContent = false
        assistantMessage.statusBanner = ''
        messagesRef.value.splice(assistantIndex, 1, { ...assistantMessage })
        scheduleScrollForStreamMessagesRef(messagesRef)
        return
    }
    if (chunkType === 'agent_status') {
        assistantMessage.statusBanner = text
        messagesRef.value.splice(assistantIndex, 1, { ...assistantMessage })
        scheduleScrollForStreamMessagesRef(messagesRef)
        return
    }
    if (chunkType === 'agent_nav') {
        assistantMessage.pendingNavTab = text
        messagesRef.value.splice(assistantIndex, 1, { ...assistantMessage })
        scheduleScrollForStreamMessagesRef(messagesRef)
        return
    }
    if (chunkType === 'thought') {
        assistantMessage.thoughtChain = (assistantMessage.thoughtChain || '') + text
    } else {
        assistantMessage.awaitingContent = false
        assistantMessage.statusBanner = ''
        assistantMessage.content += text
    }
    messagesRef.value.splice(assistantIndex, 1, { ...assistantMessage })
    scheduleScrollForStreamMessagesRef(messagesRef)
}

function mapDbRowToChatMessage(row) {
    if (!row || row.role === 'user') {
        return { role: 'user', content: String(row?.content || '') }
    }
    let extra = {}
    try {
        const raw = row.extraJson != null ? row.extraJson : row.extra_json
        extra = raw ? JSON.parse(raw) : {}
    } catch {
        extra = {}
    }
    const m = {
        role: 'assistant',
        content: String(row.content || ''),
        thoughtChain: String(extra.thoughtChain || ''),
        awaitingContent: false,
        statusBanner: '',
        pendingNavTab: '',
        streamFailed: !!extra.streamFailed,
    }
    if (extra.isHtml) {
        m.isHtml = true
    }
    if (extra.clarificationCard && typeof extra.clarificationCard === 'object') {
        m.clarificationCard = extra.clarificationCard
    }
    return m
}

async function loadChatSessions() {
    const email = userStore.user?.email
    if (!email) {
        chatSessions.value = []
        return
    }
    chatSessionsLoading.value = true
    try {
        const r = await fetch(`/api/chat/sessions?email=${encodeURIComponent(email)}`)
        const d = await r.json().catch(() => ({}))
        if (r.ok && d.success) {
            chatSessions.value = d.items || []
        }
    } catch {
        chatSessions.value = []
    } finally {
        chatSessionsLoading.value = false
    }
}

async function ensureChatSessionBeforeSend() {
    const email = userStore.user?.email
    if (!email) {
        return null
    }
    if (currentChatSessionId.value) {
        return currentChatSessionId.value
    }
    try {
        const r = await fetch('/api/chat/sessions', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email, title: '新对话' }),
        })
        const d = await r.json().catch(() => ({}))
        if (r.ok && d.success && d.session?.id) {
            currentChatSessionId.value = d.session.id
            await loadChatSessions()
            return d.session.id
        }
    } catch (e) {
        console.warn('[chat session create]', e)
    }
    return null
}

async function persistChatTurnToServer(userText, assistantMsg) {
    const email = userStore.user?.email
    const sid = currentChatSessionId.value
    if (!email || !sid || !userText) {
        return
    }
    try {
        await fetch(`/api/chat/sessions/${sid}/append-turn`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                email,
                userText,
                assistant: {
                    content: assistantMsg.content,
                    thoughtChain: assistantMsg.thoughtChain || '',
                    clarificationCard: assistantMsg.clarificationCard || null,
                    isHtml: !!assistantMsg.isHtml,
                    streamFailed: !!assistantMsg.streamFailed,
                },
            }),
        })
        void loadChatSessions()
    } catch (e) {
        console.warn('[chat persist]', e)
    }
}

async function selectChatSession(sessionId) {
    const email = userStore.user?.email
    if (!email || !sessionId) {
        return
    }
    chatHistoryOverlayOpen.value = false
    try {
        const r = await fetch(`/api/chat/sessions/${sessionId}/messages?email=${encodeURIComponent(email)}`)
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        currentChatSessionId.value = sessionId
        const items = d.items || []
        if (items.length === 0) {
            chatMessages.value = getInitialChatMessages()
        } else {
            chatMessages.value = items.map(mapDbRowToChatMessage)
        }
        chatLoadedPaperLabel.value = ''
        activeChatSkill.value = null
        await nextTick()
        scheduleScrollChatToBottom(false)
    } catch (e) {
        ElMessage.error(String(e?.message || e || '加载会话失败'))
    }
}

async function createNewChatSession() {
    const email = userStore.user?.email
    if (!email) {
        ElMessage.warning('请先登录后再使用多会话')
        return
    }
    try {
        const r = await fetch('/api/chat/sessions', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email, title: '新对话' }),
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success || !d.session?.id) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        currentChatSessionId.value = d.session.id
        chatMessages.value = getInitialChatMessages()
        chatInput.value = ''
        chatLoadedPaperLabel.value = ''
        activeChatSkill.value = null
        await loadChatSessions()
        await nextTick()
        syncChatComposerHeight()
        scheduleScrollChatToBottom(false)
        ElMessage.success('已开始新对话')
    } catch (e) {
        ElMessage.error(String(e?.message || e || '创建失败'))
    }
}

async function deleteChatSessionRow(row, ev) {
    if (ev) {
        ev.stopPropagation()
    }
    const email = userStore.user?.email
    if (!email || !row?.id) {
        return
    }
    try {
        await ElMessageBox.confirm('确定删除该会话及其消息？', '删除会话', { type: 'warning' })
    } catch {
        return
    }
    try {
        const r = await fetch(`/api/chat/sessions/${row.id}?email=${encodeURIComponent(email)}`, { method: 'DELETE' })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        if (currentChatSessionId.value === row.id) {
            currentChatSessionId.value = null
            chatMessages.value = getInitialChatMessages()
            chatInput.value = ''
            chatLoadedPaperLabel.value = ''
            activeChatSkill.value = null
            await nextTick()
            syncChatComposerHeight()
            scheduleScrollChatToBottom(false)
        }
        await loadChatSessions()
        ElMessage.success('已删除')
    } catch (e) {
        ElMessage.error(String(e?.message || e || '删除失败'))
    }
}

async function clearAllChatSessions() {
    const email = userStore.user?.email
    if (!email) {
        return
    }
    try {
        await ElMessageBox.confirm('将删除所有历史会话与消息，不可恢复。确定？', '清空全部', { type: 'warning' })
    } catch {
        return
    }
    try {
        const r = await fetch(`/api/chat/sessions/all?email=${encodeURIComponent(email)}`, { method: 'DELETE' })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        currentChatSessionId.value = null
        chatMessages.value = getInitialChatMessages()
        chatInput.value = ''
        chatLoadedPaperLabel.value = ''
        activeChatSkill.value = null
        await loadChatSessions()
        await nextTick()
        syncChatComposerHeight()
        scheduleScrollChatToBottom(false)
        ElMessage.success('已清空全部会话')
    } catch (e) {
        ElMessage.error(String(e?.message || e || '清空失败'))
    }
}

function syncChatComposerHeight() {
    const el = chatInputTextareaRef.value
    if (!el) {
        return
    }
    el.style.height = 'auto'
    const maxH = 200
    el.style.height = `${Math.min(maxH, Math.max(44, el.scrollHeight))}px`
    el.style.overflowY = el.scrollHeight > maxH ? 'auto' : 'hidden'
}

async function sendChatMessage(overrideText, options = {}) {
    const chatMode = options.chatMode === CHAT_MODE_COMPANION ? CHAT_MODE_COMPANION : 'default'
    const isCompanion = chatMode === CHAT_MODE_COMPANION
    const messagesRef = isCompanion ? companionMessages : chatMessages
    const inputRef = isCompanion ? companionInput : chatInput

    const rawText = typeof overrideText === 'string' ? overrideText : inputRef.value
    let content = rawText.trim()
    const systemContext = typeof options.systemContext === 'string' ? options.systemContext.trim() : ''
    const paperContext =
        !isCompanion && typeof options.paperContext === 'string'
            ? options.paperContext.trim().slice(0, CHAT_INJECT_FULLTEXT_MAX)
            : ''
    if (!isCompanion && typeof options.chatContextLabel === 'string' && options.chatContextLabel.trim()) {
        chatLoadedPaperLabel.value = options.chatContextLabel.trim()
    }
    const temperatureProfile =
        !isCompanion && options.temperatureProfile === 'literature' ? 'literature' : ''
    if (!content && systemContext) {
        content =
            '（已附上材料摘要作为系统附加上下文）请基于摘要与我开始对话：先用两三句话复述你理解的核心结论，再列出你认为最值得深入追问的 3 个方向。'
    }
    if (!content) {
        return
    }

    if (speechListening.value) {
        stopSpeechInput()
    }

    if (!isCompanion && userStore.user?.email) {
        await ensureChatSessionBeforeSend()
    }

    messagesRef.value.push({ role: 'user', content })
    inputRef.value = ''
    if (!isCompanion) {
        nextTick(() => syncChatComposerHeight())
    }
    const assistantMessage = {
        role: 'assistant',
        content: '',
        thoughtChain: '',
        awaitingContent: true,
        statusBanner: '',
        pendingNavTab: '',
        streamFailed: false,
    }
    const assistantIndex = messagesRef.value.length
    messagesRef.value.push(assistantMessage)

    await nextTick()
    scheduleScrollChatToBottom(isCompanion)

    try {
        const agentPayload = isCompanion
            ? {}
            : { enableTools: true, agentContext: await buildAgentContextForChat() }
        const companionPayload =
            isCompanion && userStore.user?.email
                ? {
                      companionPersona: companionSettings.value.persona,
                      companionAddress: String(companionSettings.value.address || '').trim(),
                      companionCustomTraits: String(companionSettings.value.customTraits || '').trim(),
                  }
                : {}
        const personaPayload = !isCompanion ? resolvePersonaPayloadForChat() : {}
        const response = await fetch('/api/chat', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                message: content,
                stream: true,
                ...agentPayload,
                ...(isCompanion ? { chatMode: CHAT_MODE_COMPANION } : {}),
                ...companionPayload,
                ...(!isCompanion ? personaPayload : {}),
                ...(systemContext && !isCompanion ? { systemContext: systemContext.slice(0, CHAT_INJECT_SUMMARY_JSON_MAX) } : {}),
                ...(paperContext && !isCompanion ? { paperContext } : {}),
                ...(temperatureProfile ? { temperatureProfile } : {}),
                ...(userStore.user?.email ? { userEmail: userStore.user.email } : {}),
                ...(!isCompanion && activeChatSkill.value?.id ? { skillId: activeChatSkill.value.id } : {}),
                ...(buildClientLocalTimeLabel() ? { clientLocalTimeLabel: buildClientLocalTimeLabel() } : {}),
            }),
        })

        if (!response.ok || !response.body) {
            throw new Error(`HTTP ${response.status}`)
        }

        const reader = response.body.getReader()
        const decoder = new TextDecoder('utf-8')
        let buffer = ''

        let streamEnded = false
        let sawStreamDone = false
        try {
            while (!streamEnded) {
                let chunk
                try {
                    chunk = await reader.read()
                } catch (readErr) {
                    console.error('[chat] SSE read', readErr)
                    throw readErr
                }
                const { value, done } = chunk
                if (value?.byteLength) {
                    buffer += decoder.decode(value, { stream: !done })
                }
                const parts = buffer.split('\n\n')
                buffer = parts.pop() || ''

                for (const part of parts) {
                    if (!part.trim()) {
                        continue
                    }
                    if (
                        processChatSseEventPart(part, (chunkType, t, meta) => {
                            appendToStreamMessages(assistantMessage, assistantIndex, chunkType, t, messagesRef, meta)
                        })
                    ) {
                        sawStreamDone = true
                        streamEnded = true
                        break
                    }
                }

                await Promise.resolve()

                if (done) {
                    break
                }
            }
        } catch (streamErr) {
            assistantMessage.awaitingContent = false
            assistantMessage.streamFailed = true
            const tip =
                streamErr?.name === 'AbortError'
                    ? '请求已中断。'
                    : '流式连接异常断开，请检查网络后重试；若回复不完整可缩小单次提问范围。'
            if (!String(assistantMessage.content || '').trim()) {
                assistantMessage.content = tip
            }
            messagesRef.value.splice(assistantIndex, 1, { ...assistantMessage })
            await nextTick()
            scheduleScrollChatToBottom(isCompanion)
            ElMessage.warning(tip)
            return
        }

        buffer += decoder.decode(new Uint8Array(), { stream: false })
        if (buffer.trim()) {
            const tailParts = buffer.split('\n\n')
            for (const part of tailParts) {
                if (!part.trim()) {
                    continue
                }
                if (
                    processChatSseEventPart(part, (chunkType, t, meta) => {
                        appendToStreamMessages(assistantMessage, assistantIndex, chunkType, t, messagesRef, meta)
                    })
                ) {
                    sawStreamDone = true
                    break
                }
            }
        }

        if (!sawStreamDone) {
            assistantMessage.streamFailed = true
            const tip = '未收到完整结束标记，连接可能已中断；若内容不完整请重试。'
            ElMessage.warning(tip)
        } else {
            assistantMessage.streamFailed = false
        }

        assistantMessage.awaitingContent = false
        messagesRef.value.splice(assistantIndex, 1, { ...assistantMessage })
        await nextTick()
        scheduleScrollChatToBottom(isCompanion)
        if (!isCompanion && assistantMessage.pendingNavTab && tabTitles[assistantMessage.pendingNavTab]) {
            const tab = assistantMessage.pendingNavTab
            goTab(tab)
            if (tab === 'apiConfig') {
                fetchModelConfigs()
            }
        }
        if (!isCompanion) {
            tryAutoNavigate(assistantMessage.content)
        }

        if (!isCompanion && userStore.user?.email && currentChatSessionId.value) {
            const list = messagesRef.value
            const last = list[list.length - 1]
            const prev = list[list.length - 2]
            if (last?.role === 'assistant' && prev?.role === 'user' && typeof prev.content === 'string') {
                await persistChatTurnToServer(prev.content, last)
            }
        }

        if (userStore.user?.email) {
            const list = messagesRef.value
            const last = list[list.length - 1]
            const prev = list[list.length - 2]
            if (
                last?.role === 'assistant' &&
                prev?.role === 'user' &&
                typeof prev.content === 'string' &&
                typeof last.content === 'string'
            ) {
                try {
                    const rec = await fetch('/api/user/profile/record-turn', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            userEmail: userStore.user.email,
                            userMessage: prev.content,
                            assistantMessage: last.content,
                            chatMode: isCompanion ? 'companion' : 'default',
                        }),
                    })
                    const rd = await rec.json().catch(() => ({}))
                    void loadDashboardStats()
                    if (!isCompanion && userStore.user?.email && prev.content && last.content) {
                        void fetch('/api/todos/extract-from-chat', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({
                                userEmail: userStore.user.email,
                                userMessage: prev.content,
                                assistantMessage: last.content,
                            }),
                        })
                            .then(() => {
                                void loadTodosUpcoming()
                                void loadTodoPendingCount()
                            })
                            .catch(() => {})
                    }
                    if (!isCompanion && rec.ok && rd.hintPulse) {
                        impressionUpdating.value = true
                        setTimeout(() => {
                            impressionUpdating.value = false
                            loadUserImpressionProfile()
                        }, 12000)
                    }
                } catch {
                    /* 画像异步建模失败不影响对话 */
                }
            }
        }
    } catch (error) {
        console.error('[chat] failed', error)
        assistantMessage.awaitingContent = false
        assistantMessage.streamFailed = true
        const aborted = error?.name === 'AbortError'
        const friendly = aborted ? '请求已取消。' : '请求失败，请检查网络或服务状态后重试。'
        if (!String(assistantMessage.content || '').trim()) {
            assistantMessage.content = friendly
        }
        messagesRef.value.splice(assistantIndex, 1, { ...assistantMessage })
        await nextTick()
        scheduleScrollChatToBottom(isCompanion)
        ElMessage.warning(friendly)
    }
}

function sendCompanionQuick(text) {
    sendChatMessage(text, { chatMode: CHAT_MODE_COMPANION })
}

function clearCompanionChat() {
    companionMessages.value = [{ role: 'assistant', content: COMPANION_OPENING }]
    companionInput.value = ''
}

function revokePaperPdfUrl() {
    if (paperPdfObjectUrl.value) {
        try {
            URL.revokeObjectURL(paperPdfObjectUrl.value)
        } catch {
            /* ignore */
        }
        paperPdfObjectUrl.value = ''
    }
}

function bindPaperLocalFile(file) {
    if (!file) {
        return
    }
    revokePaperPdfUrl()
    paperLocalFile.value = file
    const name = (file.name || '').toLowerCase()
    if (file.type === 'application/pdf' || name.endsWith('.pdf')) {
        paperFileKind.value = 'pdf'
        paperPdfObjectUrl.value = URL.createObjectURL(file)
    } else {
        paperFileKind.value = 'docx'
    }
}

function onPaperPickClick() {
    paperFileInputRef.value?.click()
}

function onPaperFileSelected(ev) {
    const f = ev.target?.files?.[0]
    if (!f) {
        return
    }
    bindPaperLocalFile(f)
    ev.target.value = ''
}

function onPaperDragOver(ev) {
    ev.preventDefault()
    paperDragging.value = true
}

function onPaperDragLeave() {
    paperDragging.value = false
}

function onPaperDrop(ev) {
    ev.preventDefault()
    paperDragging.value = false
    const f = ev.dataTransfer?.files?.[0]
    if (!f) {
        return
    }
    bindPaperLocalFile(f)
}

/**
 * 通过 FormData 上传并由后端解析为纯文本（动态路径，不写死磁盘路径）
 * @param {File} file
 * @param {{ showSuccessToast?: boolean }} [opts]
 */
async function parsePaperViaFormData(file, opts = {}) {
    const { showSuccessToast = true } = opts
    if (!file) {
        throw new Error('未选择文件')
    }
    const fd = new FormData()
    fd.append('file', file)
    const em = userStore.user?.email
    if (em) {
        fd.append('userEmail', em)
    }
    const r = await fetch('/api/papers/parse-file', { method: 'POST', body: fd })
    let data = {}
    try {
        data = await r.json()
    } catch {
        data = {}
    }
    if (!r.ok || !data.success) {
        throw new Error(data.message || `HTTP ${r.status}`)
    }
    paperRawText.value = data.text || ''
    paperCharCount.value = data.charCount || 0
    paperSourceTruncated.value = !!data.truncated
    if (!paperTitle.value.trim() && file.name) {
        paperTitle.value = file.name.replace(/\.(pdf|docx)$/i, '')
    }
    if (showSuccessToast) {
        ElMessage.success('文档已解析为纯文本')
    }
    return data
}

async function parsePaperUpload() {
    if (!paperLocalFile.value) {
        ElMessage.warning('请先选择 PDF 或 Word 文件')
        return
    }
    paperParseLoading.value = true
    paperError.value = ''
    try {
        await parsePaperViaFormData(paperLocalFile.value, { showSuccessToast: true })
    } catch (e) {
        paperError.value = String(e?.message || e)
        ElMessage.error(paperError.value)
    } finally {
        paperParseLoading.value = false
    }
}

async function fetchPaperFromUrl() {
    const u = paperUrl.value.trim()
    if (!u) {
        ElMessage.warning('请粘贴 http(s) 链接')
        return
    }
    paperFetchLoading.value = true
    paperError.value = ''
    try {
        const r = await fetch('/api/papers/fetch-url', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                url: u,
                ...(userStore.user?.email ? { userEmail: userStore.user.email } : {}),
            }),
        })
        let data = {}
        try {
            data = await r.json()
        } catch {
            data = {}
        }
        if (!r.ok || !data.success) {
            throw new Error(data.message || `HTTP ${r.status}`)
        }
        paperRawText.value = data.text || ''
        paperCharCount.value = data.charCount || 0
        paperSourceTruncated.value = !!data.truncated
        paperFileKind.value = 'none'
        revokePaperPdfUrl()
        paperLocalFile.value = null
        ElMessage.success('链接正文已获取')
    } catch (e) {
        paperError.value = String(e?.message || e)
        ElMessage.error(paperError.value)
    } finally {
        paperFetchLoading.value = false
    }
}

function mergePaperPasteIntoRaw() {
    const p = paperPasteBuffer.value.trim()
    if (!p) {
        ElMessage.info('摘录区为空')
        return
    }
    paperRawText.value = paperRawText.value.trim() ? `${paperRawText.value.trim()}\n\n${p}` : p
    paperPasteBuffer.value = ''
    ElMessage.success('已合并到正文')
}

async function runPaperSummarize() {
    paperError.value = ''
    paperSummary.value = null
    paperUploadParsing.value = false
    paperAiReading.value = false
    paperSummarizeLoading.value = false

    try {
        if (paperLocalFile.value) {
            paperUploadParsing.value = true
            try {
                await parsePaperViaFormData(paperLocalFile.value, { showSuccessToast: false })
            } finally {
                paperUploadParsing.value = false
            }
        }

        const text = (paperRawText.value || '').trim()
        if (!text) {
            ElMessage.warning('请先选择文件并等待解析，或拉取链接 / 合并摘录正文')
            return
        }

        paperAiReading.value = true
        paperSummarizeLoading.value = true
        const r = await fetch('/api/papers/summarize', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                text,
                title: paperTitle.value.trim(),
            }),
        })
        let data = {}
        try {
            data = await r.json()
        } catch {
            data = {}
        }
        if (!r.ok || !data.success) {
            throw new Error(data.message || `HTTP ${r.status}`)
        }
        paperSummary.value = data.summary || null
        ElMessage.success('结构化卡片已生成')
    } catch (e) {
        paperError.value = String(e?.message || e)
        ElMessage.error(paperError.value)
    } finally {
        paperUploadParsing.value = false
        paperAiReading.value = false
        paperSummarizeLoading.value = false
    }
}

function openPaperContextChat() {
    if (!paperSummary.value || paperSummaryCardEntries.value.length === 0) {
        ElMessage.warning('请先生成结构化卡片')
        return
    }
    pushHiddenPaperContextMessage()
    chatLoadedPaperLabel.value = getPaperChatFileName()
    const systemContext = buildPaperSystemContextBlock().slice(0, CHAT_INJECT_SUMMARY_JSON_MAX)
    const paperContext = String(paperRawText.value || '').trim().slice(0, CHAT_INJECT_FULLTEXT_MAX)
    const starter =
        '请基于系统附加上下文中的材料摘要与文献正文摘录与我对话：先用两三句话复述你理解的核心结论，再列出你认为最值得深入追问的 3 个方向；若摘要与常识冲突，请指出并说明依据。'
    goTab('chat')
    nextTick(() => {
        sendChatMessage(starter, {
            systemContext,
            paperContext,
            chatContextLabel: getPaperChatFileName(),
            temperatureProfile: 'literature',
        })
    })
}

function openPaperDeepDialog() {
    const excerpt = (paperRawText.value || '').trim().slice(0, PAPER_CHAT_ATTACH_MAX)
    if (!excerpt) {
        ElMessage.warning('请先有文献正文')
        return
    }
    const truncatedNote =
        paperSourceTruncated.value || (paperRawText.value || '').trim().length > PAPER_CHAT_ATTACH_MAX
            ? '（正文较长，仅附部分节录作为对话依据；请勿编造节录外的数据或结论。）'
            : ''
    const hasSummaryCards = paperSummary.value && paperSummaryCardEntries.value.length > 0
    const summaryBlock = hasSummaryCards ? formatPaperSummaryForChat() : ''
    const summaryBudget = 3600
    const summaryClipped =
        summaryBlock.length > summaryBudget ? `${summaryBlock.slice(0, summaryBudget)}\n…（要点已截断）` : summaryBlock
    const room = Math.max(1200, PAPER_CHAT_ATTACH_MAX - (summaryClipped.length || 0) - 400)
    const excerptClipped = excerpt.length > room ? `${excerpt.slice(0, room)}\n…（正文摘录已截断）` : excerpt
    const summarySection =
        hasSummaryCards && summaryClipped.trim().length > 0
            ? `以下为根据原文自动提炼的**动态要点卡片**（维度随内容类型变化），请先承认这些要点，再结合正文与我深入讨论：\n\n---\n${summaryClipped}\n---\n\n`
            : ''
    const msg =
        `【基于材料与要点开启深度对话】${truncatedNote}\n\n` +
        summarySection +
        `以下为原文摘录，可作为事实依据：\n\n---\n${excerptClipped}\n---`
    pushHiddenPaperContextMessage()
    chatLoadedPaperLabel.value = getPaperChatFileName()
    const paperContext = excerpt.slice(0, CHAT_INJECT_FULLTEXT_MAX)
    goTab('chat')
    nextTick(() => {
        sendChatMessage(msg, {
            temperatureProfile: 'literature',
            paperContext,
            chatContextLabel: getPaperChatFileName(),
        })
    })
}

onBeforeUnmount(() => {
    revokePaperPdfUrl()
    stopSpeechInput()
    teardownBellDragListeners()
    if (typeof window !== 'undefined') {
        window.removeEventListener('resize', onBellWindowResize)
        if (streamScrollRafId) {
            cancelAnimationFrame(streamScrollRafId)
            streamScrollRafId = 0
        }
        if (todoPollTimer) {
            clearInterval(todoPollTimer)
            todoPollTimer = null
        }
    }
})

async function fetchModelConfigs() {
    modelConfigsLoading.value = true
    modelConfigsError.value = ''
    try {
        const response = await fetch('/api/model-configs')
        const data = await response.json()
        if (!data.success) {
            throw new Error(data.message || '加载失败')
        }
        modelConfigs.value = data.items || []
    } catch (error) {
        modelConfigsError.value = String(error?.message || error)
    } finally {
        modelConfigsLoading.value = false
    }
}

function openModelModal(row) {
    modelTestMessage.value = ''
    modelTestPassed.value = null
    if (row) {
        editingModelId.value = row.id
        modelForm.value = {
            name: row.name,
            type: row.type,
            base_url: row.base_url,
            model_name: String(row.model_name || '').trim().toLowerCase(),
            api_key: '',
            priority: row.priority,
            is_enabled: !!row.is_enabled,
        }
    } else {
        editingModelId.value = null
        modelForm.value = {
            name: '',
            type: 'aliyun',
            base_url: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
            model_name: 'qwen-plus',
            api_key: '',
            priority: 10,
            is_enabled: true,
        }
    }
    modelModalOpen.value = true
}

function closeModelModal() {
    modelModalOpen.value = false
}

async function submitModelForm() {
    const f = modelForm.value
    if (!f.name?.trim() || !f.base_url?.trim() || !f.model_name?.trim()) {
        showToast('error', '请填写名称、Base URL 与模型 ID')
        return
    }
    if (!editingModelId.value && !f.api_key?.trim()) {
        showToast('error', '新增配置必须填写 API Key')
        return
    }
    if (modelTestPassed.value !== true) {
        const ok = window.confirm('建议先点击「连接测试」并确认通过后再保存。是否仍要继续保存？')
        if (!ok) {
            return
        }
    }
    modelFormSaving.value = true
    try {
        const isEdit = Boolean(editingModelId.value)
        const url = isEdit ? `/api/model-configs/${editingModelId.value}` : '/api/model-configs'
        const body = {
            name: f.name.trim(),
            type: f.type,
            base_url: f.base_url.trim(),
            model_name: f.model_name.trim().toLowerCase(),
            priority: Number(f.priority) || 0,
            is_enabled: f.is_enabled,
            api_key: isEdit && !f.api_key?.trim() ? '__UNCHANGED__' : f.api_key.trim(),
        }
        const response = await fetch(url, {
            method: isEdit ? 'PUT' : 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
        })
        const data = await response.json()
        if (!data.success) {
            throw new Error(data.message || '保存失败')
        }
        showToast('success', isEdit ? '已更新配置' : '已新增配置')
        closeModelModal()
        await fetchModelConfigs()
    } catch (error) {
        console.error('[model-config]', error)
        showToast('error', String(error?.message || error))
    } finally {
        modelFormSaving.value = false
    }
}

async function deleteModelConfig(row) {
    if (!window.confirm(`确定删除「${row.name}」？`)) {
        return
    }
    try {
        const response = await fetch(`/api/model-configs/${row.id}`, { method: 'DELETE' })
        const data = await response.json()
        if (!data.success) {
            throw new Error(data.message || '删除失败')
        }
        showToast('success', '已删除')
        await fetchModelConfigs()
    } catch (error) {
        showToast('error', String(error?.message || error))
    }
}

async function persistModelRowPartial(row, patch) {
    const response = await fetch(`/api/model-configs/${row.id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            name: row.name,
            type: row.type,
            base_url: row.base_url,
            model_name: String(row.model_name || '').trim().toLowerCase(),
            priority: patch.priority != null ? patch.priority : row.priority,
            is_enabled: patch.is_enabled != null ? patch.is_enabled : row.is_enabled,
            api_key: '__UNCHANGED__',
        }),
    })
    const data = await response.json()
    if (!data.success) {
        throw new Error(data.message || '更新失败')
    }
    await fetchModelConfigs()
}

async function onToggleModelEnabled(row) {
    try {
        await persistModelRowPartial(row, { is_enabled: row.is_enabled })
    } catch (error) {
        row.is_enabled = !row.is_enabled
        showToast('error', String(error?.message || error))
    }
}

async function onModelPriorityChange(row) {
    try {
        await persistModelRowPartial(row, { priority: row.priority })
    } catch (error) {
        showToast('error', String(error?.message || error))
        await fetchModelConfigs()
    }
}

async function testModelConnection() {
    const f = modelForm.value
    if (!f.api_key?.trim()) {
        showToast('error', '请先填写 API Key 再测试连接')
        return
    }
    if (!f.base_url?.trim() || !f.model_name?.trim()) {
        showToast('error', '请先填写 Base URL 与模型 ID')
        return
    }
    modelTestLoading.value = true
    modelTestMessage.value = ''
    modelTestPassed.value = null
    try {
        const response = await fetch('/api/models/test', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                type: f.type,
                base_url: f.base_url.trim(),
                model_name: f.model_name.trim().toLowerCase(),
                api_key: f.api_key.trim(),
            }),
        })
        let data
        try {
            data = await response.json()
        } catch {
            modelTestPassed.value = false
            modelTestMessage.value = '无法解析服务器响应'
            ElMessage.error(modelTestMessage.value)
            return
        }
        if (data.success) {
            modelTestPassed.value = true
            modelTestMessage.value = '连接测试通过'
            ElMessage.success('连接测试成功')
        } else {
            modelTestPassed.value = false
            modelTestMessage.value = data.message || '测试失败'
            ElMessage.error(modelTestMessage.value)
        }
    } catch (error) {
        modelTestPassed.value = false
        modelTestMessage.value = String(error?.message || error || '连接超时')
        ElMessage.error(modelTestMessage.value)
    } finally {
        modelTestLoading.value = false
    }
}

function modelTypeLabel(type) {
    const map = { aliyun: '阿里云百炼', openai: 'OpenAI 兼容', openclaw: 'OpenClaw 网关' }
    return map[type] || type
}

function openQuick() {
  quickActionOpen.value = true
}

function closeQuick() {
  quickActionOpen.value = false
}

async function renderDocxPreview(arrayBuffer) {
    if (!wordContainer.value) {
        return
    }

    wordContainer.value.innerHTML = ''
    await renderAsync(arrayBuffer, wordContainer.value, null, {
        ignoreWidth: false,
        ignoreHeight: false,
    })
    highlightTemplateVariables(wordContainer.value)
}

function highlightTemplateVariables(container) {
    if (!container) {
        return
    }

    const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT)
    const targets = []
    let node = walker.nextNode()
    while (node) {
        if (node.nodeValue && node.nodeValue.includes('{') && node.nodeValue.includes('}')) {
            targets.push(node)
        }
        node = walker.nextNode()
    }

    const regex = /\{[^}]+\}/g
    targets.forEach((textNode) => {
        const text = textNode.nodeValue || ''
        let match = regex.exec(text)
        if (!match) {
            return
        }

        regex.lastIndex = 0
        const fragment = document.createDocumentFragment()
        let lastIndex = 0
        while ((match = regex.exec(text)) !== null) {
            const before = text.slice(lastIndex, match.index)
            if (before) {
                fragment.appendChild(document.createTextNode(before))
            }
            const span = document.createElement('span')
            span.className = 'word-highlight'
            span.textContent = match[0]
            fragment.appendChild(span)
            lastIndex = match.index + match[0].length
        }
        const after = text.slice(lastIndex)
        if (after) {
            fragment.appendChild(document.createTextNode(after))
        }
        textNode.parentNode.replaceChild(fragment, textNode)
    })
}

function buildTemplateValues() {
    return extractedVariables.value.reduce((acc, item) => {
        acc[item.name] = item.value || ''
        return acc
    }, {})
}

const MAX_AGENT_FILE_BYTES = 8 * 1024 * 1024

function fileToBase64(file) {
    return new Promise((resolve, reject) => {
        const r = new FileReader()
        r.onload = () => {
            const s = String(r.result || '')
            const i = s.indexOf(',')
            resolve(i >= 0 ? s.slice(i + 1) : s)
        }
        r.onerror = () => reject(new Error('read failed'))
        r.readAsDataURL(file)
    })
}

/** 供对话 Agent 工具链：文献当前文件 + 模板变量（体积受限） */
async function buildAgentContextForChat() {
    const ctx = {}
    const parsed = String(paperRawText.value || '').trim()
    if (parsed.length > 0) {
        const cap = 100000
        ctx.parsedPaperText = parsed.length > cap ? parsed.slice(0, cap) : parsed
        ctx.paperFileName = getPaperChatFileName()
    }
    const f = paperLocalFile.value
    if (f && typeof f.size === 'number' && f.size > 0 && f.size <= MAX_AGENT_FILE_BYTES) {
        try {
            ctx.documentBase64 = await fileToBase64(f)
            ctx.documentName = f.name || 'upload.bin'
        } catch {
            /* ignore */
        }
    }
    if (currentTemplateId.value) {
        ctx.templateId = currentTemplateId.value
        ctx.variables = buildTemplateValues()
    }
    if (currentChatSessionId.value) {
        ctx.chatSessionId = currentChatSessionId.value
    }
    return ctx
}

async function syncToWord({ render = true } = {}) {
    if (!templateArrayBuffer.value) {
        return
    }

    const zip = new PizZip(templateArrayBuffer.value)
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        delimiters: { start: '{', end: '}' },
        nullGetter: (part) => `{${part.value}}`,
    })

    doc.render(buildTemplateValues())
    const output = doc.getZip().generate({ type: 'uint8array' })
    lastWordBlob.value = new Blob([output], {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    })

    if (render) {
        await renderDocxPreview(output.buffer)
    }
}

function scheduleSyncToWord() {
    if (syncTimer) {
        return
    }

    syncTimer = setTimeout(async () => {
        syncTimer = null
        try {
            await syncToWord({ render: true })
        } catch (error) {
            console.error('[docx preview] failed', error)
        }
    }, 800)
}

function extractVariablesFromText(text) {
    if (!text) {
        return []
    }

    const regex = /\{([^}]+)\}/g
    const found = new Set()
    let match

    while ((match = regex.exec(text)) !== null) {
        const value = match[1]?.trim()
        if (value) {
            found.add(value)
        }
    }

    return Array.from(found)
}

function syncVariableList(names) {
    extractedVariables.value = names.map((name) => ({ name, value: '', isTyping: false }))
}

function updateVariableValue(name, value) {
    const entry = extractedVariables.value.find((item) => item.name === name)
    if (entry) {
        entry.value = value
        entry.isTyping = false
        extractedVariables.value = [...extractedVariables.value]
        console.log('[fill-stream] 已更新变量:', name, value)
        scheduleSyncToWord()
        return
    }

    extractedVariables.value = [
        ...extractedVariables.value,
        { name, value, isTyping: false },
    ]
    console.log('[fill-stream] 已新增变量:', name, value)
    scheduleSyncToWord()
}

function parseSseChunk(chunk) {
    const lines = chunk.split('\n')
    let event = 'message'
    let data = ''

    lines.forEach((line) => {
        if (line.startsWith('event:')) {
            event = line.replace('event:', '').trim()
            return
        }
        if (line.startsWith('data:')) {
            data += line.replace('data:', '').trim()
        }
    })

    return { event, data }
}

async function startAiFill() {
    if (!currentTemplateId.value) {
        showToast('error', '请先上传模板')
        return
    }

    if (isFilling.value) {
        return
    }

    fillError.value = ''
    isFilling.value = true
    extractedVariables.value.forEach((item) => {
        item.value = ''
        item.isTyping = true
    })

    if (fillAbortController) {
        fillAbortController.abort()
    }
    fillAbortController = new AbortController()

    try {
        const response = await fetch('/api/templates/fill-stream', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                templateId: currentTemplateId.value,
                userEmail: userStore.user?.email || '',
                context: {
                    user: userStore.user || null,
                },
            }),
            signal: fillAbortController.signal,
        })

        if (!response.ok || !response.body) {
            throw new Error(`HTTP ${response.status}`)
        }

        const reader = response.body.getReader()
        const decoder = new TextDecoder('utf-8')
        let buffer = ''
        let fullTextBuffer = ''
        const matchedKeys = new Set()

        while (true) {
            let readOut
            try {
                readOut = await reader.read()
            } catch (streamErr) {
                console.error('[fill-stream] read', streamErr)
                fillError.value = '填充流意外断开，请检查网络后重试。'
                ElMessage.warning(fillError.value)
                break
            }
            const { value, done } = readOut
            if (done) {
                break
            }

            const chunk = decoder.decode(value, { stream: true })
            console.log('收到原始流数据:', chunk)
            buffer += chunk
            const parts = buffer.split('\n\n')
            buffer = parts.pop() || ''

            for (const part of parts) {
                if (!part.trim()) {
                    continue
                }

                const { event, data } = parseSseChunk(part)
                if (!part.includes('data:')) {
                    console.log('未匹配 data 行:', part)
                }
                if (event === 'done') {
                    isFilling.value = false
                    continue
                }

                if (event === 'error') {
                    fillError.value = data || 'AI 填充失败'
                    isFilling.value = false
                    continue
                }

                if (event === 'variable') {
                    if (!data) {
                        continue
                    }

                    let payload
                    try {
                        payload = JSON.parse(data)
                    } catch (error) {
                        console.error('[fill-stream] parse failed', error)
                        continue
                    }

                    if (payload?.name) {
                        updateVariableValue(payload.name, payload.value || '')
                    }
                    continue
                }

                if (!data) {
                    continue
                }

                let payload
                try {
                    payload = JSON.parse(data)
                } catch (error) {
                    console.warn('[fill-stream] non-json chunk', data)
                    continue
                }

                if (payload?.type === 'delta' && typeof payload.text === 'string') {
                    fullTextBuffer += payload.text
                    const regex = /"([^"]+)"\s*:\s*"([^"]*)"/g
                    let match
                    while ((match = regex.exec(fullTextBuffer)) !== null) {
                        const name = match[1]
                        const valueText = match[2]
                        if (!matchedKeys.has(name)) {
                            matchedKeys.add(name)
                            updateVariableValue(name, valueText)
                        }
                    }
                }
            }
        }
        scheduleSyncToWord()
    } catch (error) {
        console.error('[fill-stream] failed', error)
        fillError.value = 'AI 填充请求失败，请稍后重试。'
    } finally {
        isFilling.value = false
        extractedVariables.value.forEach((item) => {
            item.isTyping = false
        })
    }
}

async function uploadTemplateFile(file) {
    const formData = new FormData()
    formData.append('file', file)
    const em = userStore.user?.email
    if (em) {
        formData.append('userEmail', em)
    }

    const response = await fetch('/api/templates/upload', {
        method: 'POST',
        body: formData,
    })

    if (!response.ok) {
        throw new Error(`HTTP ${response.status}`)
    }

    return response.json()
}

async function handleFileUpload(event) {
    const file = event.target.files?.[0]
    if (!file) {
        return
    }

    currentFile.value = file

    previewError.value = ''
    isPreviewLoading.value = true

    try {
        await nextTick()
        if (!wordContainer.value) {
            throw new Error('容器未找到')
        }

        wordContainer.value.innerHTML = ''
        extractedVariables.value = []
        currentTemplateId.value = null
        templateArrayBuffer.value = null

        try {
            const uploadResponse = await uploadTemplateFile(file)
            currentTemplateId.value = uploadResponse?.template?.id || null
            showToast('success', '模板已上传')
            void loadMyTemplates()
        } catch (error) {
            console.error('[template upload] failed', error)
            showToast('error', '模板上传失败，已跳过上传')
        }

        const arrayBuffer = await file.arrayBuffer()
        templateArrayBuffer.value = arrayBuffer
        await renderDocxPreview(arrayBuffer)
        const plainText = wordContainer.value?.innerText || ''
        syncVariableList(extractVariablesFromText(plainText))
    } catch (error) {
        console.error('[docx preview] failed', error)
        previewError.value = '预览失败，请检查文件格式。'
        showToast('error', '预览失败')
    } finally {
        isPreviewLoading.value = false
    }
}

function resetTemplateFile() {
    currentFile.value = null
    currentTemplateId.value = null
    extractedVariables.value = []
    previewError.value = ''
    fillError.value = ''
    if (wordContainer.value) {
        wordContainer.value.innerHTML = ''
    }
}

async function handleExportWord() {
    if (!templateArrayBuffer.value) {
        showToast('error', '请先上传模板')
        return
    }

    await syncToWord({ render: false })
    if (!lastWordBlob.value) {
        showToast('error', '导出失败，请稍后重试')
        return
    }

    saveAs(lastWordBlob.value, 'openclaw-template.docx')
}
</script>

<template>
    <div
        class="main-content workspace-dashboard-body"
        id="mainContent"
        :class="{
            'dashboard--companion': currentTab === 'companion',
            'dashboard--immersive': isImmersiveWorkbenchTab,
        }"
    >
        <button
            v-if="isImmersiveWorkbenchTab && userStore.user?.email"
            type="button"
            class="immersive-todo-floating immersive-todo-floating--root"
            :class="{ 'immersive-todo-floating--dragging': bellDragging }"
            :style="{ left: `${bellScreenPos.x}px`, top: `${bellScreenPos.y}px`, right: 'auto' }"
            title="待办事项（可拖动位置）"
            aria-label="待办事项"
            @pointerdown="onBellPointerDown"
        >
            <span class="topbar-todo-bell-icon" aria-hidden="true">&#128276;</span>
            <span v-if="todoPendingCount > 0" class="topbar-todo-badge">{{ todoPendingCount > 99 ? '99+' : todoPendingCount }}</span>
        </button>
            <!-- Topbar（沉浸 Tab 隐藏，把高度还给聊天区） -->
            <div v-show="!isImmersiveWorkbenchTab" class="topbar">
                <div>
                    <h1 class="page-title">{{ pageTitle }}</h1>
                    <span class="breadcrumb">首页 / 工作台</span>
                </div>
                <div class="topbar-actions">
                    <div class="search-box">
                        <label class="sr-only" for="dashboard-doc-search">搜索文档与模板</label>
                        <span class="search-icon">&#8981;</span>
                        <input
                            id="dashboard-doc-search"
                            name="dashboard_doc_search"
                            type="search"
                            placeholder="搜索文档、模板..."
                            autocomplete="off"
                        />
                    </div>
                    <button
                        v-if="userStore.user?.email"
                        type="button"
                        class="topbar-todo-bell"
                        title="待办事项"
                        aria-label="待办事项"
                        @click="openTodoDrawer"
                    >
                        <span class="topbar-todo-bell-icon" aria-hidden="true">&#128276;</span>
                        <span v-if="todoPendingCount > 0" class="topbar-todo-badge">{{ todoPendingCount > 99 ? '99+' : todoPendingCount }}</span>
                    </button>
                    <button class="btn btn-outline btn-sm" type="button" @click="syncDashboardData">同步</button>
                    <button class="btn btn-primary btn-sm" @click="openQuick">快速操作</button>
                    <button
                        v-if="userStore.user?.email && typeof Notification !== 'undefined'"
                        type="button"
                        class="btn btn-ghost btn-sm"
                        title="浏览器通知用于待办到时提醒"
                        @click="requestTodoNotificationPermission"
                    >
                        通知权限
                    </button>
                </div>
            </div>

            <el-drawer v-model="todoDrawerVisible" title="待办事项" direction="rtl" size="420px" append-to-body class="todo-drawer">
                <p v-if="!userStore.user?.email" class="todo-drawer-guest">登录后查看从对话与会议纪要中静默提取的待办。</p>
                <template v-else>
                    <div v-if="todosDrawerLoading" class="todo-drawer-loading">加载中…</div>
                    <ul v-else-if="todosDrawerList.length === 0" class="todo-drawer-empty">暂无待办</ul>
                    <ul v-else class="todo-drawer-list">
                        <li v-for="t in todosDrawerList" :key="t.id" class="todo-drawer-item" :class="{ 'todo-drawer-item--done': Number(t.done) === 1 }">
                            <div class="todo-drawer-item-main">
                                <span class="todo-drawer-title">{{ t.title }}</span>
                                <span class="todo-drawer-meta">{{ t.remindAt }}{{ t.source ? ' · ' + t.source : '' }}</span>
                                <p v-if="t.notes" class="todo-drawer-notes">{{ t.notes }}</p>
                            </div>
                            <div class="todo-drawer-actions">
                                <button v-if="Number(t.done) === 0" type="button" class="btn btn-outline btn-xs" @click="markTodoDoneRow(t)">完成</button>
                                <button type="button" class="btn btn-ghost btn-xs" @click="deleteTodoRow(t)">删除</button>
                            </div>
                        </li>
                    </ul>
                </template>
            </el-drawer>

            <!-- ===== TAB: OVERVIEW ===== -->
            <div v-show="currentTab === 'overview'" class="tab-content" id="tabOverview">
                <div class="stats-row">
                    <div class="stat-card">
                        <div class="stat-icon" style="background:#eef2ff;color:#4a6cf7;">D</div>
                        <div>
                            <div class="stat-value">{{ dashboardStats.conversation_turns }}</div>
                            <div class="stat-label">累计对话轮次</div>
                        </div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-icon" style="background:#ecfdf5;color:#10b981;">&#10003;</div>
                        <div>
                            <div class="stat-value">{{ dashboardStats.skills_count }}</div>
                            <div class="stat-label">已生成技能数</div>
                        </div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-icon" style="background:#fffbeb;color:#f59e0b;">T</div>
                        <div>
                            <div class="stat-value">{{ dashboardStats.papers_processed }}</div>
                            <div class="stat-label">文献处理次数</div>
                        </div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-icon" style="background:#f5f3ff;color:#7c3aed;">&#128077;</div>
                        <div>
                            <div class="stat-value">—</div>
                            <div class="stat-label">认可即沉淀</div>
                            <div class="stat-subhint">对话里点 👍 即可触发打包 Skill</div>
                        </div>
                    </div>
                </div>

                <div class="template-showcase">
                    <h2>智能模板引擎 — 核心创新</h2>
                    <p class="subtitle">使用 <strong>{变量}</strong>、<strong>[选项]</strong>、<strong>&#123;&#123;自动填充&#125;&#125;</strong> 语法创建模板，AI自动识别并精准填写。告别重复的格式指令！</p>
                    <div class="template-code-block">
                        <span style="color:#6c7086;"># 会议纪要模板 v2.0</span><br>
                        <span style="color:#a6e3a1;">## 会议主题：</span><span class="token-var">{会议主题}</span><br>
                        <span style="color:#a6e3a1;">## 时间地点：</span><span class="token-fill">&#123;&#123;自动提取时间地点&#125;&#125;</span><br>
                        <span style="color:#a6e3a1;">## 参会人员：</span><span class="token-fill">&#123;&#123;从录音中识别参会人&#125;&#125;</span><br>
                        <span style="color:#a6e3a1;">## 核心决议</span><br>
                        <span style="color:#f38ba8;">决议类型：</span><span class="token-bracket">[通过|否决|延期讨论]</span><br>
                        <span style="color:#a6e3a1;">## 待办事项</span><br>
                        <span class="token-var">{待办事项列表}</span> <span style="color:#6c7086;">AI自动提取并推送日历</span><br>
                        <span style="color:#a6e3a1;">## 分歧点</span><br>
                        <span class="token-var">{主要分歧}</span> <span style="color:#6c7086;">自动归纳不同意见</span>
                    </div>
                    <div class="template-demo-row">
                        <div class="template-demo-card" @click="goTab('templates')">
                            <div class="demo-icon">&#9679;</div>
                            <div class="label">会议纪要模板</div>
                            <div class="hint">12个变量 · 自动填充</div>
                        </div>
                        <div class="template-demo-card" @click="goTab('templates')">
                            <div class="demo-icon">&#9633;</div>
                            <div class="label">论文摘要模板</div>
                            <div class="hint">8个变量 · 结构化卡片</div>
                        </div>
                        <div class="template-demo-card" @click="goTab('templates')">
                            <div class="demo-icon">&#9650;</div>
                            <div class="label">PPT大纲模板</div>
                            <div class="hint">逐页自动生成</div>
                        </div>
                        <div class="template-demo-card" @click="goTab('templates')">
                            <div class="demo-icon">&#128203;</div>
                            <div class="label">周报汇报模板</div>
                            <div class="hint">结构化条目 · 一键填充</div>
                        </div>
                    </div>
                </div>

                <div class="section-header">
                    <h3>全部功能模块</h3>
                    <span class="see-all">拖拽排序</span>
                </div>
                <div class="cards-grid">
                    <div class="card" @click="goTab('meeting')">
                        <div class="card-icon blue">&#9679;</div>
                        <span class="card-badge badge-popular">热门</span>
                        <h3>智能会议纪要</h3>
                        <p>上传录音→Whisper转写→结构化纪要+日历待办自动推送</p>
                        <span class="card-arrow">&#8594;</span>
                    </div>
                    <div class="card" @click="goTab('paper')">
                        <div class="card-icon green">&#9633;</div>
                        <span class="card-badge badge-pro">PRO</span>
                        <h3>文献论文摘要</h3>
                        <p>上传PDF→读取+检索补充→结构化笔记卡片（方法/创新/局限）</p>
                        <span class="card-arrow">&#8594;</span>
                    </div>
                    <div class="card" @click="goTab('ppt')">
                        <div class="card-icon amber">&#9650;</div>
                        <span class="card-badge badge-new">NEW</span>
                        <h3>PPT大纲生成</h3>
                        <p>输入主题+页数→web_search搜集→逐页大纲+配图建议+讲稿</p>
                        <span class="card-arrow">&#8594;</span>
                    </div>
                    <div class="card" @click="goTab('companion')">
                        <div class="card-icon rose">&#9825;</div>
                        <h3>情感陪伴式对话</h3>
                        <p>情绪检测→身份适配→共情回复+适时建议+暖心提醒</p>
                        <span class="card-arrow">&#8594;</span>
                    </div>
                    <div class="card" @click="goTab('identity')">
                        <div class="card-icon teal">&#8644;</div>
                        <h3>多身份智能切换</h3>
                        <p>自动识别场景→会议助理/学术导师/职场助手/贴心伙伴</p>
                        <span class="card-arrow">&#8594;</span>
                    </div>
                </div>

                <div class="section-header mt-3">
                    <h3>最近活动</h3>
                </div>
                <p class="activity-placeholder">
                    暂无集中的操作流水；上方数字来自服务端统计。在各功能 Tab 中完成的文档、文献与对话会反映在数据中。
                </p>
            </div>

            <!-- ===== TAB: CHAT（Gemini 式双栏 + 悬浮输入） ===== -->
            <div v-show="currentTab === 'chat'" class="tab-content tab-chat-workbench tab-chat-gemini-root" id="tabChat">
                <div class="chat-gemini-layout" :class="{ 'chat-gemini-layout--history-open': chatHistoryOverlayOpen }">
                    <div class="chat-gemini-backdrop" @click="chatHistoryOverlayOpen = false"></div>
                    <aside class="chat-history-sidebar" aria-label="会话历史">
                        <button type="button" class="chat-new-session-btn" @click="createNewChatSession">+ 新对话</button>
                        <div v-if="!userStore.user?.email" class="chat-history-hint">登录后可同步多会话历史。</div>
                        <div v-else-if="chatSessionsLoading" class="chat-history-loading">加载中…</div>
                        <ul v-else class="chat-history-list">
                            <li
                                v-for="s in chatSessions"
                                :key="s.id"
                                class="chat-history-item"
                                :class="{ 'chat-history-item--active': currentChatSessionId === s.id }"
                                @click="selectChatSession(s.id)"
                            >
                                <span class="chat-history-title">{{ s.title || '未命名' }}</span>
                                <button
                                    type="button"
                                    class="chat-history-delete"
                                    title="删除会话"
                                    aria-label="删除会话"
                                    @click.stop="deleteChatSessionRow(s, $event)"
                                >
                                    🗑️
                                </button>
                            </li>
                        </ul>
                        <button
                            v-if="userStore.user?.email"
                            type="button"
                            class="chat-clear-all-btn"
                            @click="clearAllChatSessions"
                        >
                            清空所有记录
                        </button>
                    </aside>
                    <div class="chat-gemini-main">
                        <header class="chat-gemini-toolbar">
                            <button
                                type="button"
                                class="chat-history-menu-btn"
                                aria-label="展开或收起会话列表"
                                @click="chatHistoryOverlayOpen = !chatHistoryOverlayOpen"
                            >
                                ☰
                            </button>
                            <div class="chat-gemini-toolbar-center">
                                <span class="chat-gemini-brand">OpenClaw</span>
                                <span v-if="chatLoadedPaperLabel" class="chat-gemini-pill">{{ chatLoadedPaperLabel }}</span>
                                <span v-if="activeChatSkill" class="chat-gemini-pill chat-gemini-pill--skill">
                                    ⚡ {{ activeChatSkill.name }}
                                    <button type="button" class="chat-gemini-pill-x" title="取消 Skill" @click="clearActiveChatSkill">×</button>
                                </span>
                            </div>
                            <div class="chat-gemini-toolbar-actions">
                                <button type="button" class="btn btn-ghost btn-xs" @click="impressionDrawerVisible = true">眼中的你</button>
                                <button
                                    type="button"
                                    class="btn btn-ghost btn-xs"
                                    :disabled="skillExtractLoading || !userStore.user?.email || !canPackSkill"
                                    @click="extractSkillFromConversation"
                                >
                                    {{ skillExtractLoading ? '打包中…' : '打包 Skill' }}
                                </button>
                                <button type="button" class="btn btn-ghost btn-xs" @click="clearChat">清空</button>
                            </div>
                        </header>
                        <div ref="chatMessagesScrollRef" class="chat-messages chat-messages--gemini" id="chatMessages">
                                <div
                                    v-for="(message, index) in chatMessages"
                                    v-show="!message.hidden"
                                    :key="`msg-${index}`"
                                    class="chat-msg"
                                    :class="message.role === 'assistant' ? 'assistant' : 'user'"
                                >
                                    <div class="msg-avatar">{{ message.role === 'assistant' ? 'OC' : (userStore.user?.name || 'U').charAt(0).toUpperCase() }}</div>
                                    <div v-if="message.isHtml" class="msg-bubble" v-html="message.content"></div>
                                    <div v-else-if="message.role === 'assistant'" class="msg-bubble assistant-stream-bubble">
                                        <div v-if="message.statusBanner" class="chat-agent-status" role="status">{{ message.statusBanner }}</div>
                                        <div
                                            v-if="message.clarificationCard"
                                            class="clarification-card"
                                            :class="message.clarificationCard.resolved ? 'clarification-card--resolved' : 'clarification-card--active'"
                                        >
                                            <div class="clarification-card-title">💡 还需要一点细节</div>
                                            <template v-if="!message.clarificationCard.resolved">
                                                <p class="clarification-card-lead">请先补齐下列要点，便于我给出可用方案。</p>
                                                <div
                                                    v-for="(q, qi) in message.clarificationCard.missing_info"
                                                    :key="'clar-' + index + '-' + qi"
                                                    class="clarification-row"
                                                >
                                                    <label class="clarification-q">{{ q }}</label>
                                                    <textarea
                                                        v-model="message.clarificationCard.answers[qi]"
                                                        class="clarification-input"
                                                        rows="2"
                                                        :placeholder="'填写：与「' + (q.slice(0, 24) || '该问题') + '」相关的具体说明'"
                                                    ></textarea>
                                                </div>
                                                <div class="clarification-actions">
                                                    <button type="button" class="btn btn-primary clarification-submit" @click="submitClarificationReply(index)">
                                                        确认并继续
                                                    </button>
                                                </div>
                                            </template>
                                            <p v-else class="clarification-resolved-hint">已提交补充，正在生成后续回复…</p>
                                        </div>
                                        <details v-if="message.thoughtChain" class="chat-thought-fold">
                                            <summary class="chat-thought-summary">思考中…</summary>
                                            <pre class="chat-thought-pre">{{ message.thoughtChain }}</pre>
                                        </details>
                                        <div v-if="!message.clarificationCard || message.clarificationCard.resolved" class="chat-reply-line chat-reply-line--md">
                                            <span
                                                v-if="message.awaitingContent && !message.content && !message.clarificationCard"
                                                class="chat-thinking-placeholder"
                                            >AI 正在深度思考中…</span>
                                            <div
                                                v-else
                                                class="markdown-body"
                                                v-html="renderMarkdown(message.content)"
                                            ></div>
                                        </div>
                                        <div v-if="canShowAssistantLikeBtn(message)" class="chat-assistant-foot">
                                            <button
                                                type="button"
                                                class="chat-like-btn"
                                                title="认可本条回复并打包为专属 Skill"
                                                :disabled="skillExtractLoading || !userStore.user?.email"
                                                @click="onAssistantThumbsUp"
                                            >
                                                👍 <span class="chat-like-label">赞</span>
                                            </button>
                                        </div>
                                    </div>
                                    <div v-else class="msg-bubble">{{ message.content }}</div>
                                </div>
                            </div>
                        <div class="chat-gemini-bottom">
                            <transition name="skill-suggest-fade">
                                <div
                                    v-if="suggestedSkill && currentTab === 'chat'"
                                    class="skill-smart-suggest"
                                    role="status"
                                >
                                    <span class="skill-smart-suggest-text">
                                        检测到您可能需要「{{ suggestedSkill.skill.name }}」技能，点击一键启用？
                                    </span>
                                    <button type="button" class="btn btn-outline btn-xs" @click="applySuggestedSkill">启用</button>
                                    <button type="button" class="btn btn-ghost btn-xs" @click="dismissSkillSuggestion">忽略</button>
                                </div>
                            </transition>
                            <div class="chat-suggestions chat-suggestions--gemini" id="chatSuggestions">
                                <span class="chat-suggestion-chip" @click="sendQuickChat('帮我创建会议纪要')">创建会议纪要</span>
                                <span class="chat-suggestion-chip" @click="sendQuickChat('生成PPT大纲')">生成PPT大纲</span>
                                <span class="chat-suggestion-chip" @click="sendQuickChat('打开模板引擎')">打开模板引擎</span>
                                <span class="chat-suggestion-chip" @click="sendQuickChat('配置API')">配置API</span>
                                <span class="chat-suggestion-chip" @click="openTodoDrawer">待办事项</span>
                            </div>
                            <div class="chat-floating-composer-wrap">
                            <div
                                v-show="skillPickerOpen && savedSkills.length > 0"
                                class="skill-at-popup skill-at-popup--floating"
                                role="listbox"
                                aria-label="选择 Skill"
                            >
                                <button
                                    v-for="s in filteredSkills"
                                    :key="s.id"
                                    type="button"
                                    class="skill-at-item"
                                    @mousedown.prevent="pickSkillFromMenu(s)"
                                >
                                    <span class="skill-at-name">{{ s.name }}</span>
                                    <span v-if="s.description" class="skill-at-desc">{{ s.description }}</span>
                                </button>
                                <p v-if="filteredSkills.length === 0" class="skill-at-empty">无匹配 Skill，换个关键词试试</p>
                            </div>
                            <p v-if="speechListening && speechTarget === 'chat'" class="chat-speech-hint chat-speech-hint--composer" aria-live="polite">
                                <span class="chat-speech-dot"></span>
                                正在聆听…
                            </p>
                            <label class="sr-only" for="chatInput">对话输入</label>
                            <div class="chat-floating-composer">
                                <textarea
                                    id="chatInput"
                                    ref="chatInputTextareaRef"
                                    name="chat_input"
                                    v-model="chatInput"
                                    class="chat-composer-textarea"
                                    rows="1"
                                    placeholder="问问 OpenClaw…（Shift+Enter 换行，Enter 发送；输入 @ 唤醒技能）"
                                    @input="handleChatInputSkill"
                                    @keydown.esc.stop="closeSkillPicker"
                                    @keydown="onChatComposerKeydown"
                                ></textarea>
                                <div class="chat-composer-actions">
                                    <span class="chat-composer-hint" title="输入 @ 选择技能">@</span>
                                    <button
                                        type="button"
                                        class="chat-mic-btn chat-mic-btn--inset"
                                        :class="{ 'chat-mic-btn--listening': speechListening && speechTarget === 'chat', 'chat-mic-btn--disabled': !speechSupported }"
                                        :disabled="!speechSupported"
                                        :title="speechSupported ? '语音输入' : '不支持语音'"
                                        :aria-label="speechSupported ? '语音输入' : '语音不可用'"
                                        @click="toggleSpeechInput('chat')"
                                    >
                                        <span class="chat-mic-icon" aria-hidden="true">&#127908;</span>
                                    </button>
                                    <button type="button" class="send-btn send-btn--inset" title="发送" aria-label="发送" @click="sendChatMessage()">&#9654;</button>
                                </div>
                            </div>
                            </div>
                        </div>
                    </div>
                </div>
                <el-drawer
                    v-model="impressionDrawerVisible"
                    title="OpenClaw 眼中的你"
                    direction="rtl"
                    size="380px"
                    append-to-body
                    class="impression-drawer"
                >
                    <div class="impression-drawer-body">
                        <div class="impression-aside-head">
                            <p class="impression-aside-lead">编辑后保存，将写入长期画像并随对话注入模型。</p>
                            <p v-if="impressionUpdating" class="impression-live-pulse" aria-live="polite">
                                Agent 正在实时更新对你的了解…
                            </p>
                            <button type="button" class="btn btn-outline btn-xs impression-lib-inline" @click="router.push({ name: 'skills' })">
                                技能广场
                            </button>
                        </div>
                        <p v-if="!userStore.user?.email" class="impression-guest-hint">登录后可将画像同步到服务器。</p>
                        <p v-else-if="impressionLoading" class="impression-loading">加载中…</p>
                        <div v-else class="impression-cards impression-cards--drawer">
                            <div
                                v-for="field in IMPRESSION_FIELD_META"
                                :key="field.key"
                                class="impression-card"
                            >
                                <div class="impression-card-head">
                                    <span class="impression-card-label">{{ field.label }}</span>
                                    <button type="button" class="impression-card-edit" @click="startEditImpression(field.key)">编辑</button>
                                </div>
                                <textarea
                                    v-if="editingImpressionField === field.key"
                                    v-model="userImpression[field.key]"
                                    class="impression-card-input"
                                    rows="3"
                                    :placeholder="field.placeholder"
                                    @blur="blurImpressionField"
                                ></textarea>
                                <p
                                    v-else
                                    class="impression-card-body"
                                    :class="{ 'impression-card-body--empty': !userImpression[field.key]?.trim() }"
                                    :title="'双击编辑'"
                                    @dblclick="startEditImpression(field.key)"
                                >{{ userImpression[field.key]?.trim() || '双击或点「编辑」填写' }}</p>
                            </div>
                        </div>
                        <el-button
                            type="primary"
                            plain
                            class="impression-emotion-entry-btn"
                            :disabled="impressionLoading"
                            @click="router.push({ name: 'growth' })"
                        >
                            📊 查看我的情绪仪表盘
                        </el-button>
                        <el-button
                            type="primary"
                            class="impression-save-btn"
                            :loading="impressionSaving"
                            :disabled="!userStore.user?.email || impressionLoading"
                            @click="saveUserImpressionProfile"
                        >
                            保存印象
                        </el-button>
                    </div>
                </el-drawer>
            </div>

            <!-- ===== TAB: TEMPLATES ===== -->
            <div v-show="currentTab === 'templates'" class="tab-content" id="tabTemplates">
                <div class="templates-layout" :class="{ 'has-file': currentFile }">
                    <div v-if="!currentFile" class="card" style="grid-column:1/-1;max-width:840px;border:1.5px solid var(--border-light);">
                        <div class="card-icon blue">W</div>
                        <h3>上传 Word 模板</h3>
                        <p>支持 .docx 格式，上传后即可 1:1 原生感预览</p>
                        <div class="word-upload-area">
                            <label class="sr-only" for="templates-word-upload">上传 Word 模板文件</label>
                            <input
                                id="templates-word-upload"
                                name="templates_word_upload"
                                type="file"
                                accept=".docx"
                                class="word-upload-input"
                                aria-label="上传 Word 模板文件"
                                @change="handleFileUpload"
                            />
                            <div class="word-upload-hint">
                                <span class="word-upload-icon">&#8679;</span>
                                <div>
                                    <div style="font-weight:600;">点击或拖拽上传 .docx 文件</div>
                                    <div style="font-size:0.75rem;color:var(--text-muted);">最大 20MB，上传后自动渲染预览</div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div v-if="isPreviewLoading" class="word-upload-status">正在渲染预览，请稍候…</div>
                    <div v-if="previewError" class="word-upload-error">{{ previewError }}</div>
                    <div v-if="currentFile" class="templates-immersive">
                        <div class="word-preview immersive">
                            <div id="word-preview-container" ref="wordContainer" class="word-preview-canvas"></div>
                        </div>
                        <div class="word-variable-panel immersive">
                            <div class="panel-actions">
                                <button class="btn-fill" :disabled="isFilling || extractedVariables.length === 0" @click="startAiFill">
                                    {{ isFilling ? 'AI 填充中...' : 'AI 自动填充' }}
                                </button>
                                <button class="btn-export" :disabled="!currentTemplateId" @click="handleExportWord">导出 Word</button>
                                <button class="btn-reset" @click="resetTemplateFile">重选</button>
                            </div>
                            <div class="panel-title">变量配置面板</div>
                            <div v-if="extractedVariables.length === 0" class="panel-empty">暂无变量，上传模板后自动识别。</div>
                            <ul v-else class="panel-list">
                                <li v-for="item in extractedVariables" :key="item.name">
                                    <div class="panel-item-header">
                                        <span class="panel-name">{{ item.name }}</span>
                                        <span class="panel-status" :class="{
                                            pending: !item.isTyping && !item.value,
                                            filling: item.isTyping,
                                            completed: !item.isTyping && item.value
                                        }">
                                            <span v-if="item.isTyping">&#9203;</span>
                                            <span v-else-if="item.value">&#10003;</span>
                                            <span v-else>&#9679;</span>
                                            {{ item.isTyping ? '填充中' : (item.value ? '已完成' : '待处理') }}
                                        </span>
                                    </div>
                                    <div class="panel-value">
                                        <span v-if="item.isTyping">AI 正在思考...</span>
                                        <span v-else-if="item.value">{{ item.value }}</span>
                                        <span v-else class="panel-placeholder">待填充</span>
                                    </div>
                                </li>
                            </ul>
                            <div v-if="fillError" class="panel-error">{{ fillError }}</div>
                        </div>
                    </div>
                    <div v-if="!currentFile" class="template-syntax-guide">
                        <h4 class="template-syntax-title">占位符怎么用（真实生效规则）</h4>
                        <ul class="template-syntax-list">
                            <li v-pre>
                                在 Word 正文里用<strong>半角单花括号</strong>写变量名，例如 <code>{报告主题}</code>、<code>{汇报人}</code>。引擎（Docxtemplater）按花括号内的变量名识别槽位；说明文档里常见的双花括号写法如 <code>{{报告主题}}</code> 表示同一种占位——在 .docx 里请写成单花括号 <code>{报告主题}</code>。
                            </li>
                            <li>
                                上传模板后，右侧面板会列出识别到的变量。点击「AI 自动填充」时，模型按上下文填入；也可手动改每条变量再导出。
                            </li>
                            <li><code>[选项甲|选项乙]</code>：方括号 + 竖线，提示模型在语境中<strong>择一</strong>写入。</li>
                            <li>
                                需要 AI 扩写一段说明时，在花括号内写清楚指令，例如 <code>{请用三句话写本周风险与对策}</code>，描述越具体，填充越稳。
                            </li>
                        </ul>
                    </div>
                    <div v-if="!currentFile" class="my-templates-panel">
                        <div class="my-templates-head">
                            <h4>我的模板库</h4>
                            <span v-if="myTemplatesLoading" class="my-templates-hint">加载中…</span>
                            <span v-else class="my-templates-hint">共 {{ myTemplates.length }} 个 · 登录后上传即可入库</span>
                        </div>
                        <div v-if="!userStore.user?.email" class="my-templates-empty">登录后可在此管理已上传的 Word 模板。</div>
                        <ul v-else-if="myTemplates.length === 0" class="my-templates-empty">暂无模板：请点击上方上传 .docx，上传成功后会出现在此列表并可打开、重命名与删除。</ul>
                        <ul v-else class="my-templates-list">
                            <li v-for="row in myTemplates" :key="row.id" class="my-templates-row">
                                <span class="my-templates-name" :title="row.originalName">{{ row.originalName }}</span>
                                <span class="my-templates-meta">{{ row.uploadedAt || row.uploaded_at }}</span>
                                <div class="my-templates-actions">
                                    <button type="button" class="btn btn-outline btn-xs" @click="openMyTemplate(row)">打开</button>
                                    <button type="button" class="btn btn-outline btn-xs" @click="renameMyTemplate(row)">重命名</button>
                                    <button type="button" class="btn btn-ghost btn-xs" @click="deleteMyTemplate(row)">删除</button>
                                </div>
                            </li>
                        </ul>
                    </div>
                </div>
            </div>

            <!-- ===== TAB: MEETING ===== -->
            <div v-show="currentTab === 'meeting'" class="tab-content tab-meeting-minutes" id="tabMeeting">
                <MeetingMinutesView @processed="onMeetingProcessed" />
            </div>

    <!-- ===== TAB: PAPER（全屏学术工作站 · 双栏） ===== -->
    <div v-show="currentTab === 'paper'" class="tab-content tab-paper" id="tabPaper">
        <div v-if="!paperWorkbenchHasSource" class="paper-tab-empty">
            <el-empty description="请先上传文献，我将为您生成结构化摘要" :image-size="88">
                <template #default>
                    <p class="paper-empty-hint">支持 PDF / Word，或通过公开链接拉取正文后再生成结构化摘要。</p>
                    <div class="paper-empty-upload">
                        <label class="sr-only" for="tab-paper-upload-file-empty">上传 PDF 或 Word 文献</label>
                        <input
                            id="tab-paper-upload-file-empty"
                            ref="paperFileInputRef"
                            name="paper_upload_file_empty"
                            type="file"
                            accept=".pdf,.docx,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            class="paper-file-native"
                            @change="onPaperFileSelected"
                        />
                        <div class="paper-toolbar-actions paper-empty-toolbar">
                            <el-button size="small" type="primary" plain @click="onPaperPickClick">选择文件</el-button>
                            <div
                                class="paper-drop-zone paper-drop-zone--compact"
                                :class="{ 'paper-drop-zone--active': paperDragging }"
                                @dragover="onPaperDragOver"
                                @dragleave="onPaperDragLeave"
                                @drop="onPaperDrop"
                                @click="onPaperPickClick"
                            >
                                <span class="paper-drop-icon">&#8682;</span>
                                <span><strong>拖放</strong> PDF / Docx 到此处</span>
                            </div>
                        </div>
                        <div class="paper-url-row paper-url-row--toolbar paper-empty-url">
                            <label class="sr-only" for="tab-paper-fetch-url-empty">公开网页或 PDF 链接</label>
                            <input
                                id="tab-paper-fetch-url-empty"
                                v-model="paperUrl"
                                name="paper_fetch_url_empty"
                                type="url"
                                class="paper-input"
                                placeholder="https://… 公开网页 / arXiv 链接"
                                autocomplete="url"
                            />
                            <el-button size="small" type="primary" plain :loading="paperFetchLoading" @click="fetchPaperFromUrl">拉取网页</el-button>
                        </div>
                    </div>
                </template>
            </el-empty>
        </div>
        <template v-else>
        <p class="paper-lead">支持 PDF / Word / 公开网页链接。服务端截断超长文本，摘要模型低温输出 JSON 结构化卡片。</p>
        <div class="paper-workbench">
            <div class="paper-split">
                <div class="paper-col paper-col--source">
                    <el-card shadow="hover" class="paper-card paper-card--cue paper-card--stretch">
                        <template #header>
                            <span class="paper-card-head">原文 · PDF / 文本预览</span>
                        </template>
                        <div class="paper-source-body">
                            <div
                                class="paper-preview-shell"
                                :class="{ 'paper-preview-shell--drag': paperDragging }"
                                @dragover="onPaperDragOver"
                                @dragleave="onPaperDragLeave"
                                @drop="onPaperDrop"
                            >
                                <iframe
                                    v-if="paperFileKind === 'pdf' && paperPdfObjectUrl"
                                    :src="paperPdfObjectUrl"
                                    class="paper-pdf-frame"
                                    title="PDF 预览"
                                />
                                <pre v-else class="paper-cornell-pre">{{ paperPreviewLeft }}</pre>
                            </div>
                            <div class="paper-toolbar-below">
                                <label class="sr-only" for="tab-paper-upload-file">上传 PDF 或 Word 文献</label>
                                <input
                                    id="tab-paper-upload-file"
                                    ref="paperFileInputRef"
                                    name="paper_upload_file"
                                    type="file"
                                    accept=".pdf,.docx,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    class="paper-file-native"
                                    @change="onPaperFileSelected"
                                />
                                <div class="paper-toolbar-actions">
                                    <el-button size="small" type="primary" plain @click="onPaperPickClick">选择文件</el-button>
                                    <div
                                        class="paper-drop-zone paper-drop-zone--compact"
                                        :class="{ 'paper-drop-zone--active': paperDragging }"
                                        @dragover="onPaperDragOver"
                                        @dragleave="onPaperDragLeave"
                                        @drop="onPaperDrop"
                                        @click="onPaperPickClick"
                                    >
                                        <span class="paper-drop-icon">&#8682;</span>
                                        <span><strong>拖放</strong> PDF / Docx 到此处或预览区</span>
                                    </div>
                                    <el-button type="primary" size="small" :loading="paperParseLoading" :disabled="!paperLocalFile" @click="parsePaperUpload">
                                        解析文档
                                    </el-button>
                                    <span v-if="paperCharCount" class="paper-stat">约 {{ paperCharCount }} 字{{ paperSourceTruncated ? '（已截断）' : '' }}</span>
                                </div>
                                <div class="paper-url-row paper-url-row--toolbar">
                                    <label class="sr-only" for="tab-paper-fetch-url">公开网页或 PDF 链接</label>
                                    <input
                                        id="tab-paper-fetch-url"
                                        v-model="paperUrl"
                                        name="paper_fetch_url"
                                        type="url"
                                        class="paper-input"
                                        placeholder="https://… 公开网页 / arXiv 链接"
                                        autocomplete="url"
                                    />
                                    <el-button size="small" type="primary" plain :loading="paperFetchLoading" @click="fetchPaperFromUrl">拉取网页</el-button>
                                </div>
                                <details class="paper-details">
                                    <summary class="paper-details-summary">文献标题、摘录与合并（可选展开）</summary>
                                    <div class="paper-details-body">
                                        <label class="paper-mini-label" for="tab-paper-doc-title">文献标题（可选）</label>
                                        <input
                                            id="tab-paper-doc-title"
                                            v-model="paperTitle"
                                            name="paper_doc_title"
                                            type="text"
                                            class="paper-input"
                                            placeholder="如：Attention Is All You Need"
                                            autocomplete="off"
                                        />
                                        <label class="paper-mini-label" for="tab-paper-paste-buffer">摘录 / 补充段落（合并进正文）</label>
                                        <textarea
                                            id="tab-paper-paste-buffer"
                                            v-model="paperPasteBuffer"
                                            name="paper_paste_buffer"
                                            class="paper-paste-area paper-paste-area--adaptive"
                                            placeholder="可粘贴摘要、方法段落等…"
                                        ></textarea>
                                        <el-button size="small" class="paper-merge-btn" @click="mergePaperPasteIntoRaw">合并到正文</el-button>
                                    </div>
                                </details>
                                <p v-if="paperError" class="paper-err">{{ paperError }}</p>
                            </div>
                        </div>
                    </el-card>
                </div>
                <div class="paper-col paper-col--notes">
                    <el-card shadow="hover" class="paper-card paper-card--notes paper-card--stretch">
                        <template #header>
                            <span class="paper-card-head">笔记区 · 结构化卡片</span>
                        </template>
                        <div class="paper-notes-scroll">
                            <div v-if="!paperSummary" class="paper-notes-empty">
                                点击「生成结构化卡片」后，AI 会先判断材料类型（如论文、新闻、纪要等），再生成 4～6 张<strong>自适应要点卡片</strong>（标题与维度随内容而变）。
                            </div>
                            <div v-else class="paper-nb-wrap" :class="paperSummaryToneClass">
                                <div class="paper-type-chip" role="status">
                                    <span class="paper-type-chip__label">内容类型</span>
                                    <span class="paper-type-chip__value">{{ paperSummaryContentType }}</span>
                                </div>
                                <div class="paper-nb-grid paper-nb-grid--dynamic paper-nb-grid--fill">
                                    <section
                                        v-for="entry in paperSummaryCardEntries"
                                        :key="entry.vueKey"
                                        class="nb-block nb-block--dynamic"
                                    >
                                        <h4 class="nb-title" :class="paperDimensionTitleClass(entry.displayKey)">{{ entry.displayKey }}</h4>
                                        <p class="nb-body">{{ entry.value }}</p>
                                    </section>
                                </div>
                                <div class="paper-chat-actions">
                                    <el-button type="primary" class="paper-deep-btn paper-deep-btn--half" @click="openPaperContextChat">
                                        带入上下文对话
                                    </el-button>
                                    <el-button type="primary" plain class="paper-deep-btn paper-deep-btn--half paper-deep-btn--plain" @click="openPaperDeepDialog">
                                        要点 + 原文讨论
                                    </el-button>
                                </div>
                                <div class="paper-openclaw-command" role="navigation" aria-label="指挥 OpenClaw 处理此摘要">
                                    <div class="paper-openclaw-command__title">指挥 OpenClaw 处理此内容</div>
                                    <div class="paper-openclaw-command__grid">
                                        <el-button
                                            v-for="cmd in PAPER_OPENCLAW_QUICK_COMMANDS"
                                            :key="cmd.id"
                                            size="small"
                                            class="paper-openclaw-command__btn"
                                            @click="openPaperClawQuickCommand(cmd.id)"
                                        >
                                            {{ cmd.label }}
                                        </el-button>
                                    </div>
                                </div>
                            </div>
                            <el-button
                                type="success"
                                class="paper-gen-btn"
                                :loading="paperUploadParsing || paperSummarizeLoading"
                                :disabled="paperUploadParsing || paperSummarizeLoading"
                                @click="runPaperSummarize"
                            >
                                生成结构化卡片
                            </el-button>
                            <p v-if="paperUploadParsing" class="paper-phase-msg">正在通过服务器解析上传文件…</p>
                            <p v-else-if="paperAiReading" class="paper-phase-msg">AI 正在研读论文并提取要点…</p>
                        </div>
                    </el-card>
                </div>
            </div>
        </div>
        </template>
    </div>

<!-- ===== TAB: PPT ===== -->
<div v-show="currentTab === 'ppt'" class="tab-content" id="tabPpt">
    <div class="ppt-workbench">
        <div class="ppt-toolbar card-like">
            <div class="card-icon amber">&#9650;</div>
            <div>
                <h3>PPT 大纲与导出</h3>
                <p>由模型生成 Markdown 分页提纲，服务端用 pptxgenjs 合成真实 .pptx（下载不阻塞界面）。</p>
            </div>
        </div>
        <div class="ppt-fields">
            <label class="sr-only" for="ppt-topic-input">PPT 主题</label>
            <input
                id="ppt-topic-input"
                v-model="pptTopic"
                name="ppt_topic"
                type="text"
                class="ppt-input ppt-input--grow"
                placeholder="主题，例如：AI 在医疗影像中的应用"
            />
            <label class="sr-only" for="ppt-page-count">页数</label>
            <input
                id="ppt-page-count"
                v-model.number="pptPages"
                name="ppt_page_count"
                type="number"
                min="3"
                max="50"
                class="ppt-input ppt-input--num"
            />
            <label class="sr-only" for="ppt-style-select">风格</label>
            <select id="ppt-style-select" v-model="pptStyle" name="ppt_style" class="ppt-input ppt-input--style">
                <option>商务专业风</option>
                <option>学术严谨风</option>
                <option>活泼创意风</option>
                <option>极简主义风</option>
            </select>
            <button type="button" class="btn btn-accent btn-sm" :disabled="pptGenLoading" @click="generatePptOutlineAction">
                {{ pptGenLoading ? '生成中…' : '生成 Markdown 提纲' }}
            </button>
            <button type="button" class="btn btn-primary btn-sm" :disabled="pptExportLoading || !pptOutlineMd.trim()" @click="exportPptFile">
                {{ pptExportLoading ? '导出中…' : '导出为 PPT 文件' }}
            </button>
        </div>
        <label class="sr-only" for="ppt-outline-md">提纲 Markdown</label>
        <textarea
            id="ppt-outline-md"
            v-model="pptOutlineMd"
            name="ppt_outline_md"
            class="ppt-outline-editor"
            rows="14"
            placeholder='建议使用 ## 作为每一页标题，下一行用 "- " 列出要点，例如：&#10;# 季度复盘&#10;## 第 1 页：背景&#10;- 要点一&#10;- 要点二'
        ></textarea>
    </div>
</div>

<!-- ===== TAB: COMPANION ===== -->
<div v-show="currentTab === 'companion'" class="tab-content tab-companion" id="tabCompanion">
    <div class="chat-container companion-chat-shell chat-container--immersive">
        <div class="chat-header">
            <div class="chat-avatar companion-avatar">&#9825;</div>
            <div>
                <div class="chat-title">情感陪伴 · 暖心扣字</div>
                <div class="chat-status companion-status">在这儿陪你</div>
            </div>
            <button type="button" class="btn btn-ghost btn-xs" style="margin-left:auto;" @click="clearCompanionChat" title="清空并重置开场">清空</button>
        </div>
        <div class="companion-persona-panel">
            <div class="companion-persona-row">
                <label class="companion-persona-label" for="companion-persona-select">陪伴人格</label>
                <select id="companion-persona-select" v-model="companionSettings.persona" class="companion-persona-select" name="companion_persona">
                    <option value="default">默认 OpenClaw</option>
                    <option value="strict_mentor">严厉导师</option>
                    <option value="warm_sister">温柔知心大姐</option>
                    <option value="grumpy_bro">暴躁老哥</option>
                    <option value="custom">自定义</option>
                </select>
                <input
                    v-model="companionSettings.address"
                    type="text"
                    class="companion-persona-input"
                    name="companion_address"
                    placeholder="希望怎么称呼你（可选）"
                />
                <button type="button" class="btn btn-outline btn-xs" @click="saveCompanionPersonaToProfile">保存并用于对话</button>
            </div>
            <textarea
                v-show="companionSettings.persona === 'custom'"
                v-model="companionSettings.customTraits"
                class="companion-persona-textarea"
                name="companion_custom"
                rows="2"
                placeholder="描述自定义性格、语气与边界（保存后随请求发往服务端）"
            ></textarea>
        </div>
        <div ref="companionMessagesScrollRef" class="chat-messages companion-messages">
            <div
                v-for="(message, index) in companionMessages"
                :key="`cmp-${index}`"
                class="chat-msg"
                :class="message.role === 'assistant' ? 'assistant' : 'user'"
            >
                <div class="msg-avatar">{{ message.role === 'assistant' ? '&#9825;' : (userStore.user?.name || 'U').charAt(0).toUpperCase() }}</div>
                <div v-if="message.role === 'assistant'" class="msg-bubble assistant-stream-bubble companion-bubble">
                    <details v-if="message.thoughtChain" class="chat-thought-fold companion-thought-fold">
                        <summary class="chat-thought-summary">思考中…</summary>
                        <pre class="chat-thought-pre">{{ message.thoughtChain }}</pre>
                    </details>
                    <div class="chat-reply-line chat-reply-line--md">
                        <span
                            v-if="message.awaitingContent && !message.content"
                            class="chat-thinking-placeholder companion-thinking-hint"
                        >稍等哈，我在琢磨怎么说～</span>
                        <div
                            v-else
                            class="markdown-body markdown-body--companion"
                            v-html="renderMarkdown(message.content)"
                        ></div>
                    </div>
                </div>
                <div v-else class="msg-bubble companion-bubble-user">{{ message.content }}</div>
            </div>
        </div>
        <div class="companion-mood-chips" aria-label="情绪快捷键">
            <button
                v-for="(chip, idx) in companionMoodChips"
                :key="idx"
                type="button"
                class="companion-mood-chip"
                @click="sendCompanionQuick(chip.text)"
            >
                {{ chip.label }}
            </button>
        </div>
        <div class="chat-input-area companion-input-area">
            <p v-if="speechListening && speechTarget === 'companion'" class="chat-speech-hint chat-speech-hint--warm" aria-live="polite">
                <span class="chat-speech-dot"></span>
                正在聆听…
            </p>
            <label class="sr-only" for="companion-chat-input">陪伴对话输入</label>
            <div class="chat-input-toolbar companion-input-toolbar">
                <textarea
                    id="companion-chat-input"
                    v-model="companionInput"
                    name="companion_chat_input"
                    rows="1"
                    class="companion-textarea chat-input-textarea"
                    placeholder="随便聊聊，我都在～"
                    @keydown.enter.prevent="sendChatMessage(undefined, { chatMode: CHAT_MODE_COMPANION })"
                ></textarea>
                <div class="chat-input-trailing">
                    <button
                        type="button"
                        class="chat-mic-btn chat-mic-btn--warm"
                        :class="{ 'chat-mic-btn--listening': speechListening && speechTarget === 'companion', 'chat-mic-btn--disabled': !speechSupported }"
                        :disabled="!speechSupported"
                        :title="speechSupported ? '语音输入（点击开始/结束）' : '当前浏览器不支持语音输入'"
                        @click="toggleSpeechInput('companion')"
                    >
                        <span class="chat-mic-icon" aria-hidden="true">&#127908;</span>
                    </button>
                    <button type="button" class="send-btn companion-send" @click="sendChatMessage(undefined, { chatMode: CHAT_MODE_COMPANION })" title="发送">&#9654;</button>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- ===== TAB: IDENTITY ===== -->
<div v-show="currentTab === 'identity'" class="tab-content identity-switcher-tab" id="tabIdentity">
    <div class="identity-switcher-lead">
        <p>
            下列身份作用于<strong>对话助手</strong>主对话（与「情感陪伴」页的人格设置相互独立）。点击卡片即切换；下一次在对话里发送消息起，服务端会将对应指令合并进模型 system，语气会立即改变。
        </p>
        <button type="button" class="btn btn-primary" @click="openCustomPersonaDialog">+ 创建自定义伙伴</button>
    </div>
    <div class="identity-persona-grid">
        <div
            v-for="p in BUILTIN_CHAT_PERSONAS"
            :key="p.id"
            class="card identity-persona-card"
            :class="{ 'identity-persona-card--active': chatActivePersonaId === p.id }"
            role="button"
            tabindex="0"
            @click="selectWorkbenchPersona(p.id)"
            @keydown.enter.prevent="selectWorkbenchPersona(p.id)"
        >
            <div class="identity-persona-emoji" aria-hidden="true">{{ p.emoji }}</div>
            <h3>{{ p.title }}</h3>
            <p>{{ p.desc }}</p>
            <span class="identity-persona-pill">{{ chatActivePersonaId === p.id ? '当前使用' : '点击切换' }}</span>
        </div>
        <div
            v-for="p in chatCustomPersonas"
            :key="p.id"
            class="card identity-persona-card identity-persona-card--custom"
            :class="{ 'identity-persona-card--active': chatActivePersonaId === p.id }"
            role="button"
            tabindex="0"
            @click="selectWorkbenchPersona(p.id)"
            @keydown.enter.prevent="selectWorkbenchPersona(p.id)"
        >
            <div class="identity-persona-emoji" aria-hidden="true">{{ p.emoji }}</div>
            <h3>{{ p.name }}</h3>
            <p class="identity-custom-snippet">
                {{ p.systemPrompt.length > 96 ? `${p.systemPrompt.slice(0, 96)}…` : p.systemPrompt }}
            </p>
            <div class="identity-persona-actions">
                <span class="identity-persona-pill">{{ chatActivePersonaId === p.id ? '当前使用' : '点击切换' }}</span>
                <button type="button" class="btn btn-ghost btn-xs" @click.stop="deleteCustomPersona(p)">删除</button>
            </div>
        </div>
    </div>
</div>

<el-dialog
    v-model="customPersonaDialogOpen"
    title="创建自定义伙伴"
    width="480px"
    append-to-body
    destroy-on-close
    class="custom-persona-dialog"
>
    <label class="welcome-wizard-label" for="cp-name">名称</label>
    <input
        id="cp-name"
        v-model="customPersonaForm.name"
        class="welcome-wizard-input"
        maxlength="80"
        name="custom_persona_name"
        placeholder="如：答辩专用副驾"
        autocomplete="off"
    />
    <label class="welcome-wizard-label" for="cp-emoji">头像（Emoji）</label>
    <input
        id="cp-emoji"
        v-model="customPersonaForm.emoji"
        class="welcome-wizard-input"
        maxlength="8"
        name="custom_persona_emoji"
        placeholder="如 🤖"
    />
    <label class="welcome-wizard-label" for="cp-prompt">性格设定（System Prompt）</label>
    <textarea
        id="cp-prompt"
        v-model="customPersonaForm.systemPrompt"
        class="welcome-wizard-input welcome-wizard-textarea"
        rows="6"
        maxlength="2500"
        name="custom_persona_prompt"
        placeholder="描述语气、禁忌、输出结构等；保存后将合并进对话请求的系统指令。"
    />
    <template #footer>
        <el-button @click="customPersonaDialogOpen = false">取消</el-button>
        <el-button type="primary" @click="submitCustomPersona">保存并使用</el-button>
    </template>
</el-dialog>

<!-- ===== TAB: REMINDERS ===== -->
<div v-show="currentTab === 'reminders'" class="tab-content" id="tabReminders">
    <div class="cards-grid">
        <div class="card" style="grid-column:1/-1;max-width:600px;">
            <div class="card-icon amber">&#9200;</div>
            <h3>主动提醒与推送</h3>
            <p>从会议纪要自动提取待办，定时推送提醒</p>
            <div style="background:var(--warning-light);border:1px solid #fde68a;border-radius:var(--radius);padding:12px 16px;margin-top:10px;">
                <strong>下午3:00</strong> — 产品评审会（自动从纪要提取）<br>
                <strong>明天9:00</strong> — 提交文献综述初稿<br>
                <strong>周五</strong> — 周报截止提醒
            </div>
            <button class="btn btn-accent btn-sm mt-2" style="width:100%;">+ 手动添加提醒</button>
        </div>
    </div>
</div>

<!-- ===== TAB: API CONFIG (新增API配置) ===== -->
<div v-show="currentTab === 'apiConfig'" class="tab-content tab-api-config" id="tabApiConfig">
    <div class="api-config-toolbar">
        <el-button-group class="api-config-btn-group">
            <el-button :disabled="modelConfigsLoading" @click="fetchModelConfigs">刷新列表</el-button>
            <el-button type="primary" @click="openModelModal(null)">新增配置</el-button>
        </el-button-group>
    </div>
    <p class="api-config-intro">对话与模板填充将自动选用「已启用且优先级最高」的一条配置；数字越大优先级越高。列表中的 Key 为脱敏展示。【模型 ID】须与百炼 Code 一致（小写，如 <code>glm-5.1</code>、<code>qwen-plus</code>）；保存与请求时会自动转为小写。展示名可另填「GLM-5.1」等便于识别。</p>

    <div class="api-whisper-asr-card">
        <h3 class="api-whisper-asr-title">语音识别 (ASR) 配置</h3>
        <p class="api-whisper-asr-desc">
            与上方「文本对话模型」<strong>完全分离</strong>：Whisper 转写使用独立 Base URL（须支持 OpenAI 兼容的
            <code>/v1/audio/transcriptions</code>，例如官方 <code>https://api.openai.com/v1</code>）。仅对话用的阿里云/豆包地址<strong>不能</strong>用于转写，否则会 404。
        </p>
        <div class="api-whisper-asr-grid">
            <div class="api-whisper-asr-field">
                <label class="api-whisper-asr-label" for="whisper-asr-key">Whisper API Key</label>
                <input
                    id="whisper-asr-key"
                    v-model="whisperAsrApiKey"
                    name="whisper_asr_api_key"
                    type="password"
                    class="api-whisper-asr-input"
                    placeholder="留空则会议转写时与「已启用文本模型」共用 Key"
                    autocomplete="new-password"
                />
            </div>
            <div class="api-whisper-asr-field">
                <label class="api-whisper-asr-label" for="whisper-asr-base">Whisper Base URL</label>
                <input
                    id="whisper-asr-base"
                    v-model="whisperAsrBaseUrl"
                    name="whisper_asr_base_url"
                    type="url"
                    class="api-whisper-asr-input"
                    placeholder="https://api.openai.com/v1"
                    autocomplete="url"
                />
            </div>
        </div>
        <div class="api-whisper-asr-actions">
            <button type="button" class="btn btn-primary btn-sm" @click="saveWhisperAsrToStorage">保存 ASR 配置到本机</button>
            <span class="api-whisper-asr-note">配置写入浏览器 localStorage，刷新不丢失；发起会议纪要上传/录音时会一并提交给服务器。</span>
        </div>
    </div>

    <div v-if="modelConfigsLoading" class="model-config-status">加载配置中…</div>
    <div v-else-if="modelConfigsError" class="model-config-error">{{ modelConfigsError }}</div>
    <div v-else class="model-config-flow">
        <div
            v-for="row in modelConfigs"
            :key="row.id"
            class="model-config-card"
            :class="{ 'is-off': !row.is_enabled }"
        >
            <div class="mcc-head">
                <div class="mcc-head-text">
                    <div class="mcc-title">{{ row.name }}</div>
                    <div class="mcc-sub">{{ modelTypeLabel(row.type) }} · {{ row.model_name }}</div>
                </div>
                <label class="toggle-switch" :for="'mcc-enabled-' + row.id">
                    <input
                        :id="'mcc-enabled-' + row.id"
                        v-model="row.is_enabled"
                        :name="'mcc_enabled_' + row.id"
                        type="checkbox"
                        :aria-label="'启用「' + row.name + '」模型配置'"
                        @change="onToggleModelEnabled(row)"
                    />
                    <span class="slider"></span>
                </label>
            </div>
            <div class="mcc-body">
                <div class="mcc-line"><span class="mcc-k">Base</span><span class="mcc-v" :title="row.base_url">{{ row.base_url }}</span></div>
                <div class="mcc-line"><span class="mcc-k">Key</span><span class="mcc-v">{{ row.has_api_key ? row.api_key_masked : '未设置' }}</span></div>
                <div class="mcc-line mcc-pri-line">
                    <label class="mcc-k mcc-k--label" :for="'mcc-priority-' + row.id">优先级</label>
                    <input
                        :id="'mcc-priority-' + row.id"
                        v-model.number="row.priority"
                        class="mcc-priority-input"
                        type="number"
                        :name="'mcc_priority_' + row.id"
                        min="0"
                        max="9999"
                        @change="onModelPriorityChange(row)"
                    />
                </div>
            </div>
            <div class="mcc-foot">
                <el-button-group class="mcc-btn-group">
                    <el-button size="small" @click="openModelModal(row)">编辑</el-button>
                    <el-button size="small" type="danger" plain @click="deleteModelConfig(row)">删除</el-button>
                </el-button-group>
            </div>
        </div>
        <div v-if="modelConfigs.length === 0" class="model-config-empty">暂无模型配置。请点击「新增配置」，或在首次启动时编辑数据库中的默认占位项并填写密钥。</div>
    </div>
</div>

    </div>

<el-dialog
    v-model="welcomeWizardOpen"
    title="欢迎来到 OpenClaw"
    width="420px"
    align-center
    append-to-body
    class="welcome-wizard-dialog"
    :close-on-click-modal="false"
    :show-close="false"
    destroy-on-close
>
    <p class="welcome-wizard-lead">用 30 秒告诉我你是谁，我会更贴心地对齐语气与节奏。</p>
    <label class="welcome-wizard-label" for="wiz-identity">你的身份？</label>
    <input
        id="wiz-identity"
        v-model="welcomeForm.identity"
        type="text"
        class="welcome-wizard-input"
        maxlength="120"
        placeholder="如：产品经理、博士生、法务…"
    />
    <label class="welcome-wizard-label" for="wiz-worry">最近的烦恼？</label>
    <textarea
        id="wiz-worry"
        v-model="welcomeForm.worry"
        class="welcome-wizard-input welcome-wizard-textarea"
        rows="2"
        maxlength="500"
        placeholder="一两句话即可，可写赶 ddl、人际、健康焦虑等"
    ></textarea>
    <label class="welcome-wizard-label" for="wiz-tone">偏好的沟通语气？</label>
    <input
        id="wiz-tone"
        v-model="welcomeForm.tone"
        type="text"
        class="welcome-wizard-input"
        maxlength="200"
        placeholder="如：直爽少废话 / 温柔多鼓励 / 偏学术严谨…"
    />
    <template #footer>
        <el-button type="primary" style="width: 100%" @click="submitWelcomeWizard">保存并开始使用</el-button>
    </template>
</el-dialog>

<div v-show="modelModalOpen" class="modal-overlay" @click.self="closeModelModal">
    <div class="modal model-config-modal">
        <button type="button" class="modal-close" @click="closeModelModal">&#10005;</button>
        <h2 style="margin-bottom:12px;">{{ editingModelId ? '编辑模型配置' : '新增模型配置' }}</h2>
        <div
            v-if="modelTestPassed !== null"
            class="mcm-alert"
            :class="modelTestPassed ? 'mcm-alert--ok' : 'mcm-alert--fail'"
            role="status"
        >
            <template v-if="modelTestPassed">连接测试已通过，可以安全保存。</template>
            <template v-else>连接测试未通过：{{ modelTestMessage }}</template>
        </div>
        <div class="mcm-fields">
            <label class="mcm-label" for="mcm-display-name">显示名称</label>
            <input id="mcm-display-name" v-model="modelForm.name" name="mcm_display_name" type="text" class="mcm-input" placeholder="例如：生产百炼" />

            <label class="mcm-label" for="mcm-type">接入类型</label>
            <select id="mcm-type" v-model="modelForm.type" name="mcm_type" class="mcm-input">
                <option value="aliyun">阿里云百炼（OpenAI 兼容）</option>
                <option value="openai">OpenAI 兼容</option>
                <option value="openclaw">OpenClaw 本地网关</option>
            </select>

            <label class="mcm-label" for="mcm-base-url">Base URL</label>
            <input id="mcm-base-url" v-model="modelForm.base_url" name="mcm_base_url" type="text" class="mcm-input" placeholder="https://..." />

            <label class="mcm-label" for="mcm-model-name">模型 ID（官方接口名）</label>
            <input
                id="mcm-model-name"
                v-model="modelForm.model_name"
                name="mcm_model_name"
                type="text"
                class="mcm-input"
                placeholder="glm-5.1、qwen-plus 等（小写）"
                @blur="modelForm.model_name = String(modelForm.model_name || '').trim().toLowerCase()"
            />
            <p class="mcm-field-hint">
                百炼「模型 Code」须与文档一致、通常为小写（如 <code>glm-5.1</code>、<code>qwen-plus</code>）。保存与调用时会自动转为小写；界面展示名称请用上方「显示名称」。
            </p>

            <label class="mcm-label" for="mcm-api-key">API Key</label>
            <input
                id="mcm-api-key"
                v-model="modelForm.api_key"
                name="mcm_api_key"
                type="password"
                class="mcm-input"
                :placeholder="editingModelId ? '留空则不修改已存密钥' : '必填，完整密钥仅用于保存与测试'"
                autocomplete="new-password"
            />

            <label class="mcm-label" for="mcm-priority">优先级（越大越优先）</label>
            <input id="mcm-priority" v-model.number="modelForm.priority" name="mcm_priority" type="number" class="mcm-input" min="0" max="9999" />

            <label class="mcm-check" for="mcm-enabled"><input id="mcm-enabled" v-model="modelForm.is_enabled" name="mcm_enabled" type="checkbox" /> 启用此配置</label>

            <div class="mcm-test-row">
                <button type="button" class="btn btn-outline btn-sm" :disabled="modelTestLoading" @click="testModelConnection">
                    {{ modelTestLoading ? '测试中…' : '连接测试' }}
                </button>
            </div>
        </div>
        <div class="mcm-actions">
            <button type="button" class="btn btn-outline" @click="closeModelModal">取消</button>
            <button type="button" class="btn btn-primary" :disabled="modelFormSaving" @click="submitModelForm">{{ modelFormSaving ? '保存中…' : '保存' }}</button>
        </div>
    </div>
</div>

<!-- ==================== MODAL: QUICK ACTION ==================== -->
<div v-show="quickActionOpen" class="modal-overlay" id="quickActionModal" @click.self="closeQuick">
    <div class="modal">
        <button class="modal-close" @click="closeQuick">&#10005;</button>
        <h2 style="margin-bottom:12px;">快速操作</h2>
        <div style="display:flex;flex-direction:column;gap:8px;">
            <button class="btn btn-outline" style="width:100%;justify-content:flex-start;" @click="closeQuick(); goTab('meeting')">新建会议纪要</button>
            <button class="btn btn-outline" style="width:100%;justify-content:flex-start;" @click="closeQuick(); goTab('paper')">上传论文PDF</button>
            <button class="btn btn-outline" style="width:100%;justify-content:flex-start;" @click="closeQuick(); goTab('templates')">创建模板</button>
            <button class="btn btn-outline" style="width:100%;justify-content:flex-start;" @click="closeQuick(); goTab('chat')">对话助手</button>
            <button class="btn btn-outline" style="width:100%;justify-content:flex-start;" @click="closeQuick(); goTab('apiConfig')">配置API</button>
        </div>
    </div>
</div>
</template>

<style scoped>
.word-upload-area {
    position: relative;
    border: 2px dashed var(--border);
    border-radius: var(--radius-lg);
    padding: 24px;
    margin-top: 12px;
    transition: var(--transition);
    background: var(--surface-alt);
}

.word-upload-area:hover {
    border-color: var(--primary);
    background: var(--primary-light);
}

.word-upload-input {
    position: absolute;
    inset: 0;
    opacity: 0;
    cursor: pointer;
}

.word-upload-hint {
    display: flex;
    align-items: center;
    gap: 12px;
    color: var(--text-primary);
}

.word-upload-icon {
    font-size: 1.8rem;
    color: var(--primary);
}

.word-upload-status {
    margin-top: 10px;
    font-size: 0.85rem;
    color: var(--text-secondary);
}

.word-upload-error {
    margin-top: 10px;
    font-size: 0.85rem;
    color: var(--danger);
}

.word-preview {
    margin-top: 16px;
    border: 1px solid var(--border-light);
    border-radius: var(--radius-lg);
    background-color: transparent !important;
    color: #000000 !important;
    padding: 0;
    flex: 1 1 auto;
}

.word-preview-canvas {
    width: 100%;
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 24px;
    min-height: 320px;
}

.templates-immersive {
    display: flex;
    gap: 0;
    margin-top: 0;
    align-items: stretch;
    height: calc(100vh - 64px);
    overflow: hidden;
}

.word-preview.immersive {
    border: none;
    border-radius: 0;
    margin: 0;
    flex: 1 1 auto;
    background-color: rgb(245, 243, 240) !important;
    padding: 40px;
    display: flex;
    justify-content: center;
    overflow-y: auto;
    -webkit-overflow-scrolling: touch;
    touch-action: pan-x pan-y pinch-zoom;
}

#word-preview-container {
    background: transparent;
    width: 100%;
    min-height: auto;
    margin: 0;
    box-shadow: none;
    padding: 40px 0;
    display: flex;
    flex-direction: column;
    align-items: center;
}

#word-preview-container * {
    color: #333333 !important;
    background-color: #ffffff !important;
}

#word-preview-container .docx-wrapper {
    background-color: #ffffff !important;
    box-shadow: 0 8px 30px rgba(0, 0, 0, 0.15) !important;
    margin: 0 auto !important;
    width: 800px !important;
    min-height: 1100px !important;
    padding: 40px 48px !important;
    border: none !important;
}

.word-variable-panel {
    width: 220px;
    min-height: 600px;
    border: 1px solid var(--border-light);
    border-radius: var(--radius-lg);
    background: var(--surface-alt);
    padding: 12px;
}

.word-variable-panel.immersive {
    position: relative;
    width: 320px;
    min-height: 100vh;
    border-radius: 0;
    border-left: 1px solid #e0e0e0;
    box-shadow: -2px 0 15px rgba(0, 0, 0, 0.03);
    padding: 24px 20px;
    background: #ffffff;
    position: sticky;
    top: 64px;
}

.panel-actions {
    display: flex;
    gap: 8px;
    margin-bottom: 20px;
    align-items: center;
}

.panel-actions .btn-fill {
    flex: 1;
    background: linear-gradient(135deg, #1e40af 0%, #3b82f6 100%);
    border: none;
    border-radius: 8px;
    color: white;
    padding: 10px 16px;
    font-weight: 600;
    font-size: 0.85rem;
    cursor: pointer;
    transition: all 0.2s ease;
    box-shadow: 0 4px 12px rgba(30, 64, 175, 0.3);
}

.panel-actions .btn-fill:hover:not(:disabled) {
    transform: translateY(-1px);
    box-shadow: 0 6px 16px rgba(30, 64, 175, 0.4);
}

.panel-actions .btn-fill:disabled {
    opacity: 0.6;
    cursor: not-allowed;
}

.panel-actions .btn-export {
    padding: 10px 14px;
    border: 1.5px solid #e5e7eb;
    border-radius: 8px;
    background: transparent;
    color: #4b5563;
    font-size: 0.85rem;
    cursor: pointer;
    transition: all 0.2s ease;
}

.panel-actions .btn-export:hover {
    border-color: #3b82f6;
    color: #3b82f6;
}

.panel-actions .btn-reset {
    padding: 6px;
    background: none;
    border: none;
    color: #9ca3af;
    font-size: 0.75rem;
    cursor: pointer;
    transition: color 0.2s ease;
}

.panel-actions .btn-reset:hover {
    color: #ef4444;
}

.panel-title {
    font-weight: 600;
    font-size: 0.95rem;
    color: #1f2937;
    margin-bottom: 16px;
    padding-bottom: 12px;
    border-bottom: 1px solid #f3f4f6;
}

.panel-empty {
    font-size: 0.85rem;
    color: #9ca3af;
    text-align: center;
    padding: 30px 20px;
}

.panel-list {
    padding: 0;
    display: flex;
    flex-direction: column;
    gap: 10px;
}

.panel-list li {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 10px;
    padding: 14px;
    display: flex;
    flex-direction: column;
    gap: 6px;
    position: relative;
    overflow: hidden;
    transition: all 0.2s ease;
}

.panel-list li::before {
    content: '';
    position: absolute;
    left: 0;
    top: 0;
    bottom: 0;
    width: 3px;
    background: linear-gradient(180deg, #3b82f6 0%, #1d4ed8 100%);
    border-radius: 0 2px 2px 0;
}

.panel-list li:hover {
    border-color: #3b82f6;
    box-shadow: 0 4px 16px rgba(59, 130, 246, 0.15);
    transform: translateY(-2px);
}

.panel-item-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
}

.panel-name {
    font-weight: 600;
    font-size: 0.85rem;
    color: #1f2937;
}

.panel-status {
    font-size: 0.7rem;
    padding: 2px 6px;
    border-radius: 10px;
}

.panel-status.pending {
    color: #9ca3af;
    background: #f3f4f6;
}

.panel-status.filling {
    color: #3b82f6;
    background: #eff6ff;
    animation: pulse 1.5s ease-in-out infinite;
}

.panel-status.completed {
    color: #10b981;
    background: #ecfdf5;
}

@keyframes pulse {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.5; }
}

.panel-value {
    color: #6b7280;
    font-size: 0.85rem;
    line-height: 1.5;
    word-break: break-all;
}

.panel-placeholder {
    color: #9ca3af;
    font-style: italic;
}

.identity-switcher-tab {
    padding: 8px 12px 28px;
    max-width: 1100px;
}

.identity-switcher-lead {
    display: flex;
    flex-wrap: wrap;
    align-items: flex-start;
    justify-content: space-between;
    gap: 12px 16px;
    margin-bottom: 18px;
}

.identity-switcher-lead p {
    margin: 0;
    flex: 1;
    min-width: 220px;
    font-size: 0.88rem;
    color: #475569;
    line-height: 1.55;
}

.identity-persona-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
    gap: 14px;
}

.identity-persona-card {
    cursor: pointer;
    transition: border-color 0.15s ease, box-shadow 0.15s ease;
    border-width: 2px;
}

.identity-persona-card:focus-visible {
    outline: 2px solid #6366f1;
    outline-offset: 2px;
}

.identity-persona-card--active {
    border-color: #6366f1;
    box-shadow: 0 6px 18px rgba(99, 102, 241, 0.18);
}

.identity-persona-emoji {
    font-size: 2rem;
    line-height: 1;
    margin-bottom: 6px;
}

.identity-persona-card h3 {
    margin: 0 0 6px;
    font-size: 1rem;
}

.identity-persona-card p {
    margin: 0 0 10px;
    font-size: 0.82rem;
    color: var(--text-muted);
    line-height: 1.45;
}

.identity-custom-snippet {
    font-size: 0.76rem !important;
    color: #64748b !important;
}

.identity-persona-pill {
    display: inline-block;
    font-size: 0.72rem;
    padding: 3px 10px;
    border-radius: 999px;
    background: #eef2ff;
    color: #4338ca;
    font-weight: 600;
}

.identity-persona-card--active .identity-persona-pill {
    background: #4338ca;
    color: #fff;
}

.identity-persona-actions {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    flex-wrap: wrap;
}

.templates-layout {
    display: flex;
    flex-direction: column;
    gap: 12px;
    height: 100%;
}

.templates-layout.has-file {
    height: calc(100vh - 64px);
}

#tabTemplates {
    height: calc(100vh - 64px);
    display: flex;
    flex-direction: column;
    overflow: hidden;
    background-color: rgb(245, 243, 240);
}

@media (max-width: 960px) {
    .templates-immersive {
        flex-direction: column;
        height: auto;
    }

    .word-preview.immersive {
        padding: 24px;
    }

    .word-variable-panel.immersive {
        width: 100%;
        min-height: auto;
        border-left: none;
        border-top: 1px solid #e2e8f0;
        position: relative;
        top: 0;
    }
}

#tabTemplates {
    height: calc(100vh - 64px);
}

.panel-error {
    margin-top: 8px;
    color: var(--danger);
    font-size: 0.8rem;
}

.dashboard-reminder-bar {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 8px 12px;
    padding: 10px 20px;
    margin: 0 20px 12px;
    border-radius: var(--radius-lg);
    border: 1px solid #e0e7ff;
    background: linear-gradient(92deg, #eef2ff, #fff);
    font-size: 0.78rem;
}

.dashboard-reminder-label {
    font-weight: 700;
    color: #4338ca;
}

.dashboard-reminder-label--overdue {
    color: #b91c1c;
}

.dashboard-reminder-chip {
    padding: 4px 10px;
    border-radius: 999px;
    background: #fff;
    border: 1px solid #c7d2fe;
    color: #334155;
    max-width: 280px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}

.dashboard-reminder-chip--overdue {
    border-color: #fecaca;
    background: #fef2f2;
    color: #991b1b;
}

.card-like {
    display: flex;
    gap: 14px;
    align-items: flex-start;
    padding: 16px 18px;
    border-radius: var(--radius-lg);
    border: 1px solid var(--border-light);
    background: var(--surface-alt);
    margin-bottom: 14px;
}

.template-syntax-guide {
    margin-top: 16px;
    padding: 14px 16px;
    border-radius: var(--radius);
    border: 1px dashed #cbd5e1;
    background: #f8fafc;
}

.template-syntax-title {
    margin: 0 0 10px;
    font-size: 0.92rem;
}

.template-syntax-list {
    margin: 0;
    padding-left: 18px;
    font-size: 0.8rem;
    color: #475569;
    line-height: 1.55;
}

.template-syntax-list code {
    font-size: 0.76rem;
    background: #fff;
    padding: 1px 6px;
    border-radius: 4px;
}

.my-templates-panel {
    margin-top: 16px;
    padding: 14px 16px;
    border-radius: var(--radius-lg);
    border: 1px solid var(--border-light);
    background: #fff;
}

.my-templates-head {
    display: flex;
    align-items: baseline;
    gap: 12px;
    margin-bottom: 10px;
}

.my-templates-head h4 {
    margin: 0;
    font-size: 0.95rem;
}

.my-templates-hint {
    font-size: 0.75rem;
    color: var(--text-muted);
}

.my-templates-empty {
    font-size: 0.82rem;
    color: var(--text-muted);
    list-style: none;
    padding: 0;
    margin: 0;
}

.my-templates-list {
    list-style: none;
    padding: 0;
    margin: 0;
    display: flex;
    flex-direction: column;
    gap: 8px;
}

.my-templates-row {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 8px;
    padding: 8px 10px;
    border-radius: 8px;
    border: 1px solid #e2e8f0;
    font-size: 0.82rem;
}

.my-templates-name {
    font-weight: 600;
    flex: 1;
    min-width: 120px;
}

.my-templates-meta {
    font-size: 0.72rem;
    color: #94a3b8;
}

.my-templates-actions {
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    margin-left: auto;
}

.ppt-workbench {
    padding: 0 8px 24px;
    max-width: 1100px;
}

.ppt-toolbar h3 {
    margin: 0 0 6px;
    font-size: 1.05rem;
}

.ppt-toolbar p {
    margin: 0;
    font-size: 0.82rem;
    color: var(--text-muted);
}

.ppt-fields {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    align-items: center;
    margin-bottom: 12px;
}

.ppt-input {
    padding: 10px 12px;
    border: 1px solid var(--border);
    border-radius: var(--radius);
    font-family: var(--font-sans);
    font-size: 0.85rem;
}

.ppt-input--grow {
    flex: 1;
    min-width: 200px;
}

.ppt-input--num {
    width: 72px;
}

.ppt-input--style {
    min-width: 140px;
}

.ppt-outline-editor {
    width: 100%;
    box-sizing: border-box;
    padding: 12px;
    border-radius: var(--radius);
    border: 1px solid var(--border);
    font-family: ui-monospace, monospace;
    font-size: 0.8rem;
    line-height: 1.45;
    resize: vertical;
}

.word-highlight {
    background: #fff3b0;
    border: 1px dashed #f59e0b;
    border-radius: 3px;
    padding: 0 4px;
}

.api-config-intro {
    color: var(--text-secondary);
    font-size: 0.85rem;
    margin-bottom: 16px;
    line-height: 1.55;
}

.api-whisper-asr-card {
    margin-bottom: 20px;
    padding: 18px 18px 16px;
    border-radius: var(--radius-lg);
    border: 1px solid #c7d2fe;
    background: linear-gradient(165deg, #f8fafc 0%, #eef2ff 100%);
}

.api-whisper-asr-title {
    margin: 0 0 8px;
    font-size: 1.05rem;
    font-weight: 700;
    color: #1e293b;
}

.api-whisper-asr-desc {
    margin: 0 0 14px;
    font-size: 0.82rem;
    color: #475569;
    line-height: 1.55;
}

.api-whisper-asr-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 14px;
}

@media (max-width: 768px) {
    .api-whisper-asr-grid {
        grid-template-columns: 1fr;
    }
}

.api-whisper-asr-field {
    display: flex;
    flex-direction: column;
    gap: 6px;
}

.api-whisper-asr-label {
    font-size: 0.78rem;
    font-weight: 600;
    color: #334155;
}

.api-whisper-asr-input {
    width: 100%;
    box-sizing: border-box;
    padding: 10px 12px;
    border-radius: var(--radius);
    border: 1px solid var(--border);
    font-size: 0.85rem;
    font-family: var(--font-sans, system-ui, sans-serif);
}

.api-whisper-asr-actions {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 10px 14px;
    margin-top: 14px;
}

.api-whisper-asr-note {
    font-size: 0.75rem;
    color: #64748b;
    line-height: 1.45;
    max-width: 520px;
}

.model-config-status,
.model-config-error {
    padding: 14px;
    border-radius: var(--radius);
    font-size: 0.88rem;
}

.model-config-error {
    background: var(--danger-light);
    color: var(--danger);
    border: 1px solid #fecaca;
}

.model-config-flow {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
    gap: 14px;
}

.model-config-card {
    background: linear-gradient(165deg, #ffffff 0%, #faf9f7 100%);
    border: 1px solid var(--border-light);
    border-radius: var(--radius-lg);
    padding: 16px;
    box-shadow: var(--shadow-sm);
    display: flex;
    flex-direction: column;
    gap: 12px;
    transition: var(--transition);
}

.model-config-card:hover {
    border-color: #c7d2fe;
    box-shadow: var(--shadow-md);
}

.model-config-card.is-off {
    opacity: 0.72;
}

.mcc-head {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 10px;
}

.mcc-title {
    font-weight: 700;
    font-size: 0.95rem;
    color: #1f2937;
}

.mcc-sub {
    font-size: 0.75rem;
    color: var(--text-muted);
    margin-top: 4px;
}

.mcc-body {
    display: flex;
    flex-direction: column;
    gap: 6px;
    font-size: 0.78rem;
}

.mcc-line {
    display: flex;
    gap: 8px;
    align-items: baseline;
}

.mcc-k {
    flex: 0 0 52px;
    color: var(--text-muted);
    font-weight: 600;
}

.mcc-v {
    color: #333;
    word-break: break-all;
}

.mcc-pri-line {
    align-items: center;
    margin-top: 4px;
}

.mcc-priority-input {
    width: 88px;
    padding: 4px 8px;
    border: 1px solid var(--border);
    border-radius: 8px;
    font-size: 0.8rem;
    background: #fff;
}

.mcc-foot {
    display: flex;
    justify-content: flex-end;
    gap: 8px;
    margin-top: auto;
    padding-top: 4px;
}

.mcc-del {
    color: var(--danger) !important;
}

.model-config-empty {
    grid-column: 1 / -1;
    text-align: center;
    padding: 28px 16px;
    color: var(--text-muted);
    font-size: 0.88rem;
    border: 1px dashed var(--border);
    border-radius: var(--radius-lg);
    background: var(--surface-alt);
}

.model-config-modal {
    max-width: 440px;
    width: 92vw;
}

.mcm-fields {
    display: flex;
    flex-direction: column;
    gap: 10px;
}

.mcm-label {
    font-size: 0.75rem;
    font-weight: 600;
    color: var(--text-secondary);
}

.mcm-input {
    width: 100%;
    padding: 9px 11px;
    border: 1.5px solid var(--border);
    border-radius: var(--radius-sm);
    font-size: 0.85rem;
    font-family: var(--font-sans);
    background: var(--surface-alt);
}

.mcm-check {
    display: flex;
    align-items: center;
    gap: 8px;
    font-size: 0.82rem;
    color: var(--text-primary);
    margin-top: 4px;
}

.mcm-test-row {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 10px;
    margin-top: 8px;
}

.mcm-alert {
    font-size: 0.82rem;
    line-height: 1.45;
    padding: 10px 12px;
    border-radius: var(--radius-sm);
    margin-bottom: 12px;
    border: 1px solid transparent;
}

.mcm-alert--ok {
    background: #ecfdf5;
    border-color: #a7f3d0;
    color: #065f46;
}

.mcm-alert--fail {
    background: var(--danger-light);
    border-color: #fecaca;
    color: #b91c1c;
}

.mcm-actions {
    display: flex;
    justify-content: flex-end;
    gap: 10px;
    margin-top: 18px;
}

.assistant-stream-bubble {
    display: flex;
    flex-direction: column;
    gap: 10px;
    text-align: left;
}

.chat-agent-status {
    font-size: 0.72rem;
    line-height: 1.4;
    color: #0f766e;
    background: linear-gradient(90deg, #ecfdf5, #f0fdfa);
    border: 1px solid #99f6e4;
    border-radius: 999px;
    padding: 6px 12px;
    align-self: flex-start;
    max-width: 100%;
}

.clarification-card {
    align-self: stretch;
    max-width: 100%;
    border-radius: 14px;
    padding: 14px 16px;
    border: 1px solid #fed7aa;
    background: linear-gradient(145deg, #fffbeb 0%, #fff7ed 45%, #eff6ff 100%);
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.06);
}

.clarification-card--resolved {
    border-color: #bfdbfe;
    background: linear-gradient(135deg, #f8fafc, #eff6ff);
}

.clarification-card-title {
    font-weight: 700;
    font-size: 0.95rem;
    color: #9a3412;
    margin-bottom: 8px;
}

.clarification-card-lead {
    font-size: 0.78rem;
    color: #57534e;
    margin: 0 0 12px;
    line-height: 1.45;
}

.clarification-row {
    margin-bottom: 12px;
}

.clarification-q {
    display: block;
    font-size: 0.8rem;
    font-weight: 600;
    color: #1c1917;
    margin-bottom: 6px;
    line-height: 1.35;
}

.clarification-input {
    width: 100%;
    box-sizing: border-box;
    border-radius: 10px;
    border: 1px solid #fdba74;
    padding: 8px 10px;
    font-size: 0.82rem;
    line-height: 1.45;
    resize: vertical;
    min-height: 44px;
    background: rgba(255, 255, 255, 0.92);
    font-family: inherit;
}

.clarification-input:focus {
    outline: none;
    border-color: #3b82f6;
    box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.2);
}

.clarification-actions {
    display: flex;
    justify-content: flex-end;
    margin-top: 6px;
}

.clarification-submit {
    font-size: 0.85rem;
    padding: 8px 18px;
}

.clarification-resolved-hint {
    margin: 0;
    font-size: 0.78rem;
    color: #0369a1;
}

.chat-header {
    display: flex;
    align-items: flex-start;
    gap: 10px;
    flex-wrap: wrap;
}

.chat-header-actions {
    display: flex;
    flex-shrink: 0;
    align-items: center;
    gap: 8px;
    flex-wrap: wrap;
    margin-left: auto;
}

.skill-pack-btn {
    max-width: min(100%, 240px);
}

.chat-active-skill-pill {
    margin-top: 8px;
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 4px 10px;
    border-radius: 999px;
    font-size: 0.72rem;
    background: linear-gradient(90deg, #fef3c7, #fef9c3);
    border: 1px solid #fcd34d;
    color: #92400e;
}

.chat-active-skill-clear {
    border: none;
    background: transparent;
    cursor: pointer;
    padding: 0 4px;
    font-size: 0.85rem;
    line-height: 1;
    color: #78350f;
}

.chat-input-area--skills {
    position: relative;
    z-index: 2;
}

.chat-input-toolbar {
    display: flex;
    align-items: stretch;
    gap: 0;
    width: 100%;
    min-height: 48px;
}

.chat-input-textarea {
    flex: 1;
    min-width: 0;
    border: 1px solid #d6d3d1;
    border-radius: 12px 0 0 12px;
    border-right: none;
    padding: 10px 12px;
    font-size: 0.9rem;
    resize: vertical;
    min-height: 48px;
    max-height: 200px;
    font-family: inherit;
    box-sizing: border-box;
}

.chat-input-textarea:focus {
    outline: none;
    border-color: #93c5fd;
    box-shadow: 0 0 0 2px rgba(59, 130, 246, 0.15);
}

.chat-input-trailing {
    display: flex;
    flex-direction: column;
    justify-content: stretch;
    border: 1px solid #d6d3d1;
    border-left: none;
    border-radius: 0 12px 12px 0;
    overflow: hidden;
    background: #fafaf9;
}

.chat-mic-btn {
    flex: 1;
    min-height: 24px;
    border: none;
    background: #f5f5f4;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 4px 10px;
    transition: background 0.2s;
}

.chat-mic-btn:hover:not(:disabled) {
    background: #e7e5e4;
}

.chat-mic-btn--listening {
    background: linear-gradient(180deg, #fff1f2, #ffe4e6);
    animation: micPulse 1.1s ease-in-out infinite;
}

.chat-mic-btn--listening .chat-mic-icon {
    filter: saturate(1.2);
}

.chat-mic-btn--disabled,
.chat-mic-btn:disabled {
    opacity: 0.4;
    cursor: not-allowed;
}

.chat-mic-btn--warm.chat-mic-btn--listening {
    background: linear-gradient(180deg, #fdf2f8, #ffe4e6);
}

.chat-mic-icon {
    font-size: 1.1rem;
    line-height: 1;
}

@keyframes micPulse {
    0%,
    100% {
        box-shadow: inset 0 0 0 0 rgba(244, 63, 94, 0.15);
    }
    50% {
        box-shadow: inset 0 0 12px 0 rgba(244, 63, 94, 0.25);
    }
}

.chat-input-trailing .send-btn {
    border-radius: 0 !important;
    border-top: 1px solid #e7e5e4 !important;
    flex: 1;
    min-height: 28px;
}

.chat-speech-hint {
    font-size: 0.72rem;
    color: #64748b;
    margin: 0 0 6px;
    display: flex;
    align-items: center;
    gap: 6px;
}

.chat-speech-hint--warm {
    color: #be185d;
}

.chat-speech-dot {
    width: 8px;
    height: 8px;
    border-radius: 50%;
    background: #ef4444;
    animation: speechDot 1s ease-in-out infinite;
}

@keyframes speechDot {
    0%,
    100% {
        opacity: 0.35;
        transform: scale(0.9);
    }
    50% {
        opacity: 1;
        transform: scale(1.1);
    }
}

.companion-input-toolbar .chat-input-textarea {
    border-radius: 12px 0 0 12px;
}

.skill-at-popup {
    position: absolute;
    bottom: 100%;
    left: 0;
    right: 52px;
    margin-bottom: 8px;
    max-height: min(40vh, 240px);
    overflow-y: auto;
    background: #fff;
    border: 1px solid #e5e7eb;
    border-radius: 12px;
    box-shadow: 0 16px 48px rgba(15, 23, 42, 0.18);
    z-index: 6000;
}

.skill-at-item {
    display: flex;
    flex-direction: column;
    align-items: flex-start;
    gap: 2px;
    width: 100%;
    padding: 10px 14px;
    border: none;
    border-bottom: 1px solid #f3f4f6;
    background: #fff;
    cursor: pointer;
    text-align: left;
    font: inherit;
}

.skill-at-item:last-of-type {
    border-bottom: none;
}

.skill-at-item:hover {
    background: #f8fafc;
}

.skill-at-name {
    font-weight: 600;
    font-size: 0.85rem;
    color: #0f172a;
}

.skill-at-desc {
    font-size: 0.72rem;
    color: #64748b;
    line-height: 1.35;
}

.skill-at-empty {
    margin: 0;
    padding: 12px 14px;
    font-size: 0.78rem;
    color: #64748b;
}

.chat-header-text {
    min-width: 0;
    flex: 1;
}

.chat-context-pill {
    margin-top: 6px;
    font-size: 0.68rem;
    line-height: 1.35;
    color: #1e40af;
    background: #eff6ff;
    border: 1px solid #bfdbfe;
    border-radius: 8px;
    padding: 4px 10px;
    max-width: 100%;
    word-break: break-all;
}

.tab-chat-workbench {
    padding: 0 8px 16px;
    max-width: 100%;
    box-sizing: border-box;
    flex: 1;
    min-height: 0;
    display: flex;
    flex-direction: column;
}

.tab-chat-gemini-root {
    padding: 0 !important;
    min-height: min(82vh, 820px);
}

/* —— 全屏沉浸：仅对话助手 / 情感陪伴（根节点挂 dashboard--immersive） —— */
.main-content.workspace-dashboard-body.dashboard--immersive {
    padding: 0 !important;
    margin-top: 0 !important;
    margin-right: 0 !important;
    margin-bottom: 0 !important;
    position: relative;
    height: 100vh;
    min-height: 100vh;
    max-height: 100vh;
    box-sizing: border-box;
    display: flex;
    flex-direction: column;
    overflow: hidden;
    background: #ffffff;
}

.main-content.workspace-dashboard-body.dashboard--immersive.dashboard--companion {
    background: linear-gradient(180deg, #fffafc 0%, #faf9f7 38%, #f7f7f5 100%);
}

.main-content.workspace-dashboard-body.dashboard--immersive > .tab-chat-gemini-root,
.main-content.workspace-dashboard-body.dashboard--immersive > .tab-companion {
    flex: 1 1 auto;
    min-height: 0;
    display: flex;
    flex-direction: column;
    overflow: hidden;
}

.main-content.workspace-dashboard-body.dashboard--immersive .tab-chat-gemini-root {
    min-height: 0 !important;
}

.main-content.workspace-dashboard-body.dashboard--immersive .chat-gemini-layout {
    flex: 1;
    min-height: 0 !important;
    height: 100%;
    max-height: 100%;
    background: #ffffff;
}

.main-content.workspace-dashboard-body.dashboard--immersive .chat-gemini-main {
    min-height: 0;
}

.main-content.workspace-dashboard-body.dashboard--immersive .chat-messages--gemini {
    background: #ffffff;
}

.main-content.workspace-dashboard-body.dashboard--immersive .chat-suggestions.chat-suggestions--gemini {
    background: transparent !important;
}

.main-content.workspace-dashboard-body.dashboard--immersive .chat-floating-composer-wrap {
    background: transparent !important;
}

.main-content.workspace-dashboard-body.dashboard--immersive .chat-gemini-bottom {
    background: transparent;
}

.main-content.workspace-dashboard-body.dashboard--immersive .chat-history-sidebar {
    background: #ffffff;
}

.main-content.workspace-dashboard-body.dashboard--immersive .tab-companion {
    position: relative;
}

.main-content.workspace-dashboard-body.dashboard--immersive .tab-companion .chat-container {
    flex: 1;
    min-height: 0;
    height: auto !important;
}

.main-content.workspace-dashboard-body.dashboard--immersive .tab-companion .chat-header {
    padding-right: 64px;
}

.immersive-todo-floating {
    position: absolute;
    top: 16px;
    right: 24px;
    z-index: 9999;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 44px;
    height: 44px;
    padding: 0;
    border-radius: 12px;
    border: 1px solid #e2e8f0;
    background: rgba(255, 255, 255, 0.96);
    box-shadow: 0 2px 14px rgba(15, 23, 42, 0.1);
    cursor: pointer;
    color: #475569;
}

.immersive-todo-floating--root {
    position: fixed !important;
    top: auto;
    right: auto;
    bottom: auto;
    z-index: 9999;
    cursor: grab;
    touch-action: none;
    user-select: none;
}

.immersive-todo-floating--root.immersive-todo-floating--dragging {
    cursor: grabbing;
}

.immersive-todo-floating:hover {
    background: #f8fafc;
    color: #0f172a;
    border-color: #cbd5e1;
}

.topbar-todo-bell {
    position: relative;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 38px;
    height: 38px;
    margin-right: 4px;
    border: none;
    border-radius: 10px;
    background: transparent;
    cursor: pointer;
    color: #475569;
}

.topbar-todo-bell:hover {
    background: #f1f5f9;
    color: #0f172a;
}

.topbar-todo-bell-icon {
    font-size: 1.15rem;
    line-height: 1;
}

.topbar-todo-badge {
    position: absolute;
    top: 2px;
    right: 2px;
    min-width: 16px;
    height: 16px;
    padding: 0 4px;
    border-radius: 999px;
    background: #ef4444;
    color: #fff;
    font-size: 0.65rem;
    font-weight: 700;
    line-height: 16px;
    text-align: center;
}

.todo-drawer-list {
    list-style: none;
    margin: 0;
    padding: 0;
    display: flex;
    flex-direction: column;
    gap: 10px;
}

.todo-drawer-item {
    display: flex;
    gap: 10px;
    align-items: flex-start;
    padding: 12px;
    border-radius: 10px;
    border: 1px solid #e2e8f0;
    background: #fafafa;
}

.todo-drawer-item--done {
    opacity: 0.55;
}

.todo-drawer-item-main {
    flex: 1;
    min-width: 0;
}

.todo-drawer-title {
    font-weight: 600;
    font-size: 0.88rem;
    color: #0f172a;
}

.todo-drawer-meta {
    display: block;
    font-size: 0.72rem;
    color: #64748b;
    margin-top: 4px;
}

.todo-drawer-notes {
    margin: 6px 0 0;
    font-size: 0.78rem;
    color: #475569;
    line-height: 1.45;
}

.todo-drawer-actions {
    display: flex;
    flex-direction: column;
    gap: 6px;
    flex-shrink: 0;
}

.todo-drawer-guest,
.todo-drawer-empty,
.todo-drawer-loading {
    font-size: 0.85rem;
    color: #64748b;
}

.chat-gemini-layout {
    display: flex;
    flex: 1;
    min-height: min(82vh, 820px);
    max-width: 100%;
    position: relative;
    background: #f8fafc;
    border-radius: 0;
    overflow: hidden;
}

.chat-gemini-backdrop {
    display: none;
}

.chat-history-sidebar {
    width: 260px;
    flex-shrink: 0;
    border-right: 1px solid #e2e8f0;
    background: #fff;
    display: flex;
    flex-direction: column;
    padding: 12px 10px;
    gap: 8px;
    min-height: 0;
}

.chat-new-session-btn {
    width: 100%;
    padding: 10px 12px;
    border-radius: 12px;
    border: 1px dashed #94a3b8;
    background: #f8fafc;
    font-weight: 600;
    font-size: 0.85rem;
    cursor: pointer;
    color: #334155;
}

.chat-new-session-btn:hover {
    border-color: #3b82f6;
    color: #1d4ed8;
    background: #eff6ff;
}

.chat-history-hint,
.chat-history-loading {
    font-size: 0.78rem;
    color: #94a3b8;
    padding: 4px 6px;
}

.chat-history-list {
    list-style: none;
    margin: 0;
    padding: 0;
    overflow-y: auto;
    flex: 1;
    min-height: 0;
}

.chat-history-item {
    position: relative;
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 10px 32px 10px 10px;
    border-radius: 10px;
    cursor: pointer;
    font-size: 0.82rem;
    color: #334155;
    margin-bottom: 4px;
}

.chat-history-item:hover {
    background: #f1f5f9;
}

.chat-history-item--active {
    background: #e0f2fe;
    color: #0c4a6e;
}

.chat-history-title {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    flex: 1;
    min-width: 0;
}

.chat-history-delete {
    position: absolute;
    right: 6px;
    top: 50%;
    transform: translateY(-50%);
    opacity: 0;
    border: none;
    background: transparent;
    cursor: pointer;
    font-size: 0.85rem;
    padding: 4px;
    line-height: 1;
}

.chat-history-item:hover .chat-history-delete {
    opacity: 1;
}

.chat-clear-all-btn {
    margin-top: auto;
    padding: 8px;
    font-size: 0.72rem;
    color: #94a3b8;
    border: none;
    background: transparent;
    cursor: pointer;
    text-decoration: underline;
    text-underline-offset: 2px;
}

.chat-clear-all-btn:hover {
    color: #b91c1c;
}

.chat-gemini-main {
    flex: 1;
    min-width: 0;
    display: flex;
    flex-direction: column;
    min-height: 0;
    background: #fafafa;
    position: relative;
}

/** 快捷 Chip + 输入区：与主聊背景一体、加宽容器 */
.chat-gemini-bottom {
    flex-shrink: 0;
    width: 100%;
    max-width: 1200px;
    margin: 0 auto;
    padding: 0 clamp(12px, 5vw, 56px);
    box-sizing: border-box;
    background: transparent;
    border-top: none !important;
    display: flex;
    flex-direction: column;
    align-items: stretch;
    gap: 0;
    position: relative;
    z-index: 4;
}

.chat-gemini-toolbar {
    display: flex;
    align-items: center;
    gap: 10px;
    padding: 10px 14px;
    border-bottom: 1px solid #e5e7eb;
    background: #fff;
    flex-shrink: 0;
}

.chat-history-menu-btn {
    display: none;
    width: 40px;
    height: 40px;
    border-radius: 10px;
    border: 1px solid #e2e8f0;
    background: #fff;
    cursor: pointer;
    font-size: 1.1rem;
    flex-shrink: 0;
}

.chat-gemini-toolbar-center {
    flex: 1;
    min-width: 0;
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 8px;
}

.chat-gemini-brand {
    font-weight: 700;
    font-size: 0.95rem;
    color: #0f172a;
}

.chat-gemini-pill {
    font-size: 0.72rem;
    padding: 3px 8px;
    border-radius: 999px;
    background: #eef2ff;
    color: #4338ca;
    max-width: 100%;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}

.chat-gemini-pill--skill {
    display: inline-flex;
    align-items: center;
    gap: 4px;
}

.chat-gemini-pill-x {
    border: none;
    background: transparent;
    cursor: pointer;
    padding: 0 2px;
    font-size: 0.9rem;
    line-height: 1;
    color: inherit;
}

.chat-gemini-toolbar-actions {
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    flex-shrink: 0;
}

.chat-messages--gemini {
    flex: 1;
    min-height: 0;
    overflow-y: auto;
    padding: 16px 18px 120px;
    background: #fafafa;
    border: none;
    border-bottom: none !important;
}

.chat-suggestions.chat-suggestions--gemini {
    padding: 6px 0 8px;
    margin: 0;
    width: 100%;
    max-width: none;
    border-top: none !important;
    border-bottom: none !important;
    background: transparent !important;
    box-sizing: border-box;
}

.chat-floating-composer-wrap {
    position: sticky;
    bottom: 0;
    padding: 8px 0 16px;
    margin: 0;
    width: 100%;
    max-width: none;
    background: transparent !important;
    border-top: none !important;
    box-shadow: none !important;
    z-index: 5;
}

.chat-floating-composer {
    position: relative;
    display: flex;
    align-items: flex-end;
    gap: 8px;
    width: 100%;
    max-width: 100%;
    margin: 0;
    padding: 10px 12px 10px 14px;
    border-radius: 24px;
    background: #fff;
    border: 1px solid #e5e7eb;
    box-shadow: 0 4px 24px rgba(15, 23, 42, 0.08);
    box-sizing: border-box;
}

.chat-composer-textarea {
    flex: 1;
    min-width: 0;
    min-height: 44px;
    max-height: 200px;
    border: none;
    resize: none;
    font-size: 0.95rem;
    line-height: 1.45;
    font-family: inherit;
    padding: 6px 4px 6px 0;
    background: transparent;
    overflow-y: auto;
}

.chat-composer-textarea:focus {
    outline: none;
}

.chat-composer-actions {
    display: flex;
    flex-direction: row;
    align-items: center;
    gap: 6px;
    flex-shrink: 0;
    padding-bottom: 2px;
}

.chat-composer-hint {
    font-size: 0.78rem;
    font-weight: 700;
    color: #94a3b8;
    padding: 0 4px;
    user-select: none;
}

.chat-mic-btn--inset,
.send-btn--inset {
    width: 40px;
    height: 40px;
    border-radius: 12px;
    border: none;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    cursor: pointer;
}

.chat-mic-btn--inset {
    background: #f1f5f9;
}

.send-btn--inset {
    background: linear-gradient(135deg, #3b82f6, #2563eb);
    color: #fff;
}

.skill-at-popup--floating {
    position: absolute;
    left: 0;
    right: 0;
    bottom: calc(100% - 6px);
    max-width: 100%;
    width: 100%;
    margin: 0;
    box-sizing: border-box;
    z-index: 12;
}

.chat-speech-hint--composer {
    text-align: center;
    margin-bottom: 6px;
}

@media (max-width: 900px) {
    .chat-history-menu-btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
    }

    .chat-history-sidebar {
        position: fixed;
        left: 0;
        top: 0;
        bottom: 0;
        z-index: 30;
        width: min(280px, 86vw);
        transform: translateX(-102%);
        transition: transform 0.22s ease;
        box-shadow: 8px 0 32px rgba(15, 23, 42, 0.12);
    }

    .chat-gemini-layout--history-open .chat-history-sidebar {
        transform: translateX(0);
    }

    .chat-gemini-backdrop {
        display: block;
        position: fixed;
        inset: 0;
        z-index: 25;
        background: rgba(15, 23, 42, 0.35);
        opacity: 0;
        pointer-events: none;
        transition: opacity 0.2s ease;
    }

    .chat-gemini-layout--history-open .chat-gemini-backdrop {
        opacity: 1;
        pointer-events: auto;
    }
}

.chat-immersive-shell {
    display: flex;
    flex-direction: column;
    flex: 1;
    min-height: min(82vh, 780px);
    max-width: 100%;
    background: linear-gradient(180deg, #eef1f6 0%, #e8ecf3 100%);
    border-radius: 16px;
    padding: 12px;
    box-sizing: border-box;
}

.chat-container--immersive {
    border: none !important;
    box-shadow: none !important;
    background: transparent !important;
    flex: 1;
    min-height: 0;
}

.chat-container--split {
    flex: 1;
    min-height: 0;
    display: flex;
    flex-direction: column;
}

.impression-drawer-body {
    display: flex;
    flex-direction: column;
    gap: 10px;
    padding-bottom: 8px;
}

.impression-cards--drawer {
    max-height: calc(100vh - 240px);
    overflow-y: auto;
}

.paper-tab-empty {
    display: flex;
    justify-content: center;
    align-items: center;
    min-height: min(52vh, 520px);
    padding: 20px 16px 32px;
}

.paper-empty-hint {
    text-align: center;
    color: var(--text-muted, #64748b);
    font-size: 0.85rem;
    line-height: 1.55;
    margin: 0 0 16px;
    max-width: 440px;
    margin-inline: auto;
}

.paper-empty-upload {
    display: flex;
    flex-direction: column;
    gap: 14px;
    align-items: stretch;
    max-width: 520px;
    margin-inline: auto;
}

.paper-empty-toolbar {
    justify-content: center;
    flex-wrap: wrap;
}

.paper-empty-url {
    width: 100%;
}

.tab-api-config .api-config-toolbar {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 12px;
    margin-bottom: 14px;
}

.api-config-btn-group {
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.08);
    border-radius: 10px;
    overflow: hidden;
}

.mcc-btn-group {
    border-radius: 8px;
    overflow: hidden;
}

.impression-aside-head {
    flex-shrink: 0;
}

.impression-aside-title {
    margin: 0 0 6px;
    font-size: 0.95rem;
    font-weight: 700;
    color: #1e293b;
    letter-spacing: 0.02em;
}

.impression-aside-lead {
    margin: 0;
    font-size: 0.72rem;
    line-height: 1.45;
    color: #64748b;
}

.impression-guest-hint,
.impression-loading {
    margin: 0;
    font-size: 0.78rem;
    color: #94a3b8;
}

.impression-cards {
    flex: 1;
    overflow-y: auto;
    display: flex;
    flex-direction: column;
    gap: 10px;
    padding-right: 4px;
    min-height: 0;
}

.impression-card {
    background: #fff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 10px 10px 8px;
}

.impression-card-head {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 6px;
}

.impression-card-label {
    font-size: 0.72rem;
    font-weight: 700;
    color: #334155;
}

.impression-card-edit {
    font-size: 0.68rem;
    padding: 2px 8px;
    border-radius: 6px;
    border: 1px solid #cbd5e1;
    background: #f8fafc;
    color: #475569;
    cursor: pointer;
}

.impression-card-edit:hover {
    border-color: #94a3b8;
    color: #0f172a;
}

.impression-card-body {
    margin: 0;
    font-size: 0.76rem;
    line-height: 1.5;
    color: #1e293b;
    white-space: pre-wrap;
    word-break: break-word;
    min-height: 2.6em;
    cursor: text;
}

.impression-card-body--empty {
    color: #94a3b8;
    font-style: italic;
}

.impression-card-input {
    width: 100%;
    box-sizing: border-box;
    font-size: 0.76rem;
    line-height: 1.45;
    padding: 8px;
    border-radius: 8px;
    border: 1px solid #cbd5e1;
    resize: vertical;
    min-height: 64px;
    font-family: var(--font-sans, system-ui, sans-serif);
}

.impression-emotion-entry-btn {
    width: 100% !important;
    margin-top: 8px !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
}

.impression-save-btn {
    width: 100% !important;
    margin-top: 4px !important;
    border-radius: 10px !important;
}

.chat-thought-fold {
    border: 1px solid #e5e7eb;
    border-radius: 10px;
    background: #f3f4f6;
    padding: 0 10px 8px;
}

.chat-thought-summary {
    cursor: pointer;
    font-size: 0.78rem;
    font-weight: 600;
    color: #6b7280;
    padding: 8px 4px;
    list-style: none;
}

.chat-thought-summary::-webkit-details-marker {
    display: none;
}

.chat-thought-pre {
    margin: 0;
    padding: 8px 10px;
    max-height: 220px;
    overflow: auto;
    font-size: 0.74rem;
    line-height: 1.45;
    white-space: pre-wrap;
    word-break: break-word;
    color: #57534e;
    background: #fff;
    border-radius: 8px;
    border: 1px solid #e7e5e4;
}

.chat-reply-line {
    font-size: 0.85rem;
    line-height: 1.55;
    color: #333333;
    white-space: pre-wrap;
    word-break: break-word;
    min-height: 1.35em;
}

.chat-reply-line--md {
    white-space: normal;
}

.markdown-body {
    font-size: 0.85rem;
    line-height: 1.62;
    color: #1f2937;
    word-break: break-word;
    overflow-wrap: anywhere;
}

.markdown-body:first-child {
    margin-top: 0;
}

.markdown-body :deep(p) {
    margin: 0.4em 0;
}

.markdown-body :deep(p:first-child) {
    margin-top: 0;
}

.markdown-body :deep(p:last-child) {
    margin-bottom: 0;
}

.markdown-body :deep(strong) {
    font-weight: 700;
    color: #0f172a;
}

.markdown-body :deep(em) {
    font-style: italic;
    color: #334155;
}

.markdown-body :deep(a) {
    color: #2563eb;
    text-decoration: underline;
    text-underline-offset: 2px;
}

.markdown-body :deep(ul),
.markdown-body :deep(ol) {
    margin: 0.45em 0;
    padding-left: 20px;
    line-height: 1.55;
}

.markdown-body :deep(li) {
    margin: 0.2em 0;
}

.markdown-body :deep(ul) {
    list-style-type: disc;
}

.markdown-body :deep(ol) {
    list-style-type: decimal;
}

.markdown-body :deep(blockquote) {
    margin: 0.6em 0;
    padding: 0.35em 0 0.35em 12px;
    border-left: 3px solid #93c5fd;
    background: linear-gradient(90deg, rgba(239, 246, 255, 0.95), transparent);
    color: #475569;
}

.markdown-body :deep(hr) {
    border: none;
    border-top: 1px solid #e5e7eb;
    margin: 0.85rem 0;
}

.markdown-body :deep(h1),
.markdown-body :deep(h2),
.markdown-body :deep(h3),
.markdown-body :deep(h4) {
    margin: 0.65em 0 0.35em;
    font-weight: 700;
    color: #0f172a;
    line-height: 1.35;
}

.markdown-body :deep(h1) {
    font-size: 1.15rem;
}

.markdown-body :deep(h2) {
    font-size: 1.05rem;
}

.markdown-body :deep(h3) {
    font-size: 0.98rem;
}

.markdown-body :deep(pre) {
    margin: 0.55em 0;
    padding: 12px 14px;
    border-radius: 12px;
    background: linear-gradient(145deg, #f1f5f9 0%, #eef2ff 55%, #f8fafc 100%);
    border: 1px solid #e2e8f0;
    box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.75), 0 2px 8px rgba(15, 23, 42, 0.06);
    overflow-x: auto;
    font-size: 0.78rem;
    line-height: 1.5;
}

.markdown-body :deep(code) {
    font-family: ui-monospace, 'Cascadia Code', 'Segoe UI Mono', monospace;
    font-size: 0.82em;
}

.markdown-body :deep(p code),
.markdown-body :deep(li code) {
    padding: 2px 6px;
    border-radius: 6px;
    background: #f1f5f9;
    border: 1px solid #e2e8f0;
    color: #0f172a;
}

.markdown-body :deep(pre code) {
    padding: 0;
    border: none;
    background: transparent;
    font-size: inherit;
}

.markdown-body :deep(table) {
    width: 100%;
    border-collapse: collapse;
    margin: 0.55em 0;
    font-size: 0.82rem;
}

.markdown-body :deep(th),
.markdown-body :deep(td) {
    border: 1px solid #e5e7eb;
    padding: 6px 10px;
    text-align: left;
}

.markdown-body :deep(th) {
    background: #f8fafc;
    font-weight: 600;
}

.markdown-body--companion :deep(strong) {
    color: #9d174d;
}

.markdown-body--companion :deep(blockquote) {
    border-left-color: #f9a8d4;
    background: linear-gradient(90deg, rgba(253, 242, 248, 0.9), transparent);
}

.chat-thinking-placeholder {
    color: #9ca3af;
    font-size: 0.82rem;
    font-style: italic;
}

.mcm-field-hint {
    margin: -4px 0 8px;
    font-size: 0.72rem;
    line-height: 1.45;
    color: var(--text-muted);
}

.mcm-field-hint code {
    font-size: 0.7rem;
    padding: 1px 5px;
    border-radius: 4px;
    background: #f3f4f6;
    color: #374151;
}

/* —— 情感陪伴：暖色侧栏 / 主区底 —— */
.sidebar.sidebar--warm {
    border-right-color: #fbcfe8;
    background: linear-gradient(180deg, #fffbfb 0%, #fff7ed 52%, #fafafa 100%);
}

.sidebar.sidebar--warm .sidebar-header .logo-icon-sm {
    background: linear-gradient(135deg, #f472b6, #fb923c);
}

.sidebar.sidebar--warm .nav-item.active {
    background: linear-gradient(90deg, #ffe4f0, #ffedd5);
    color: #be185d;
}

.sidebar.sidebar--warm .nav-item.active::before {
    background: linear-gradient(180deg, #ec4899, #f97316);
}

.dashboard-layout.dashboard--companion .main-content {
    background: linear-gradient(180deg, #fffafc 0%, #faf9f7 35%, #f7f7f5 100%);
}

/* —— 陪伴对话区 —— */
.companion-chat-shell {
    border-color: #fbcfe8;
    box-shadow: 0 10px 40px rgba(244, 114, 182, 0.14);
}

.companion-chat-shell .chat-header {
    background: linear-gradient(92deg, #fff1f2, #fff7ed);
    border-bottom-color: #fbcfe8;
}

.companion-persona-panel {
    padding: 10px 14px;
    border-bottom: 1px solid #fce7f3;
    background: linear-gradient(180deg, #fffdfb, #fffafc);
}

.companion-persona-row {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 8px;
}

.companion-persona-label {
    font-size: 0.76rem;
    font-weight: 600;
    color: #9d174d;
    white-space: nowrap;
}

.companion-persona-select,
.companion-persona-input {
    font-size: 0.78rem;
    border-radius: 8px;
    border: 1px solid #fbcfe8;
    padding: 6px 10px;
    background: #fff;
    color: #1f2937;
}

.companion-persona-select {
    min-width: 140px;
}

.companion-persona-input {
    flex: 1;
    min-width: 140px;
}

.companion-persona-textarea {
    width: 100%;
    box-sizing: border-box;
    margin-top: 8px;
    font-size: 0.78rem;
    line-height: 1.45;
    border-radius: 8px;
    border: 1px solid #fbcfe8;
    padding: 8px 10px;
    resize: vertical;
    font-family: var(--font-sans, system-ui, sans-serif);
}

.activity-placeholder {
    margin: 0;
    padding: 8px 0 4px;
    font-size: 0.82rem;
    color: #64748b;
    line-height: 1.55;
}

.companion-chat-shell .chat-header .chat-avatar {
    background: linear-gradient(135deg, #f472b6, #fb923c);
    font-size: 1rem;
}

.companion-chat-shell .companion-status {
    color: #db2777;
}

.companion-chat-shell .companion-status::before {
    background: #f472b6;
}

.companion-messages {
    background: linear-gradient(180deg, #fffafb, #faf8f6) !important;
}

.companion-chat-shell .companion-bubble,
.companion-chat-shell .companion-bubble-user {
    border-radius: 22px !important;
    padding: 12px 16px !important;
}

.companion-chat-shell .chat-msg.assistant .companion-bubble {
    border-bottom-left-radius: 12px !important;
    border: 1px solid #fce7f3;
    background: #fff;
    box-shadow: 0 2px 12px rgba(244, 114, 182, 0.08);
}

.companion-chat-shell .chat-msg.user .companion-bubble-user {
    border-bottom-right-radius: 12px !important;
    background: linear-gradient(135deg, #ec4899, #f97316);
    color: #fff;
    border: none;
    box-shadow: 0 4px 14px rgba(236, 72, 153, 0.25);
}

.companion-chat-shell .chat-msg.assistant .msg-avatar {
    background: linear-gradient(135deg, #f9a8d4, #fdba74);
    font-size: 0.65rem;
}

.companion-thought-fold {
    border-color: #fce7f3;
    background: #fff5f7;
}

.companion-thinking-hint {
    color: #db2777;
    opacity: 0.85;
}

.companion-mood-chips {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    padding: 10px 16px 0;
    margin-bottom: 12px;
    background: #fffafb;
    border-top: 1px solid #fce7f3;
}

.companion-mood-chip {
    font-size: 0.72rem;
    padding: 6px 12px;
    border-radius: 999px;
    border: 1px solid #fbcfe8;
    background: #fff;
    color: #9d174d;
    cursor: pointer;
    font-family: var(--font-sans);
    transition: transform 0.15s ease, box-shadow 0.15s ease, background 0.15s ease;
}

.companion-mood-chip:hover {
    background: linear-gradient(135deg, #fdf2f8, #fff7ed);
    box-shadow: 0 2px 10px rgba(244, 114, 182, 0.18);
    transform: translateY(-1px);
}

.companion-input-area {
    border-top-color: #fce7f3;
    background: #fff;
}

.companion-textarea {
    border-radius: 20px !important;
    border-color: #fbcfe8 !important;
}

.companion-send {
    background: linear-gradient(135deg, #f472b6, #fb923c) !important;
    border: none !important;
    color: #fff !important;
}

/* —— 文献摘要 · 全屏学术工作站（米白底 + 墨绿强调） —— */
.tab-paper {
    width: 100%;
    max-width: none;
    margin: 0;
    padding: 0 24px 28px;
    box-sizing: border-box;
    background: rgb(245, 243, 240);
    min-height: calc(100vh - 120px);
    overflow-x: hidden;
}

.paper-lead {
    font-size: 0.82rem;
    color: #5c7268;
    margin: 0 0 16px;
    line-height: 1.5;
    max-width: 72rem;
}

.paper-workbench {
    width: 100%;
    max-width: 100%;
    overflow-x: hidden;
    box-sizing: border-box;
}

.paper-split {
    display: grid;
    grid-template-columns: minmax(0, 9fr) minmax(0, 11fr);
    align-items: stretch;
    gap: 20px;
    width: 100%;
    max-width: 100%;
    min-height: calc(100vh - 168px);
    box-sizing: border-box;
    overflow-x: hidden;
}

.paper-col {
    min-width: 0;
    display: flex;
    flex-direction: column;
}

.paper-col--source,
.paper-col--notes {
    width: 100%;
    max-width: 100%;
}

.paper-card {
    border-radius: 12px !important;
    border: 1px solid #c5ddd2 !important;
    background: linear-gradient(165deg, #f6faf8 0%, #f0f7f4 100%) !important;
    box-shadow: 0 6px 22px rgba(15, 61, 50, 0.07) !important;
}

.paper-card--stretch {
    height: 100%;
    display: flex;
    flex-direction: column;
    min-height: 0;
}

.paper-card--stretch :deep(.el-card__body) {
    flex: 1;
    display: flex;
    flex-direction: column;
    min-height: 0;
    padding: 14px 16px 16px;
    overflow-x: hidden;
    overflow-y: auto;
}

.paper-card :deep(.el-card__header) {
    background: linear-gradient(90deg, #0f3d32, #1a5c4a);
    color: #ecfdf5;
    font-weight: 600;
    padding: 12px 16px;
    border-bottom: 1px solid #0a2e26;
    border-radius: 12px 12px 0 0 !important;
}

.paper-card-head {
    letter-spacing: 0.02em;
}

.paper-card--notes :deep(.el-card__header) {
    background: linear-gradient(90deg, #134e4a, #1e6b5e);
}

.paper-source-body {
    display: flex;
    flex-direction: column;
    flex: 1;
    min-height: 0;
    gap: 0;
}

.paper-preview-shell {
    position: relative;
    width: 100%;
    height: calc(100vh - 180px);
    max-height: calc(100vh - 180px);
    min-height: 280px;
    border: 1px solid #b8d4c9;
    border-radius: 12px;
    background: #fff;
    box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.85), 0 4px 18px rgba(15, 61, 50, 0.06);
    overflow: hidden;
    flex-shrink: 0;
    transition: box-shadow 0.2s ease, border-color 0.2s ease;
}

.paper-preview-shell--drag {
    border-color: #1a5c4a;
    box-shadow: inset 0 0 0 2px rgba(26, 92, 74, 0.25), 0 6px 24px rgba(15, 61, 50, 0.12);
}

.paper-pdf-frame {
    display: block;
    width: 100%;
    height: 100%;
    min-height: 0;
    border: 0;
    background: #fff;
}

.paper-cornell-pre {
    margin: 0;
    padding: 14px;
    height: 100%;
    max-height: none;
    min-height: 0;
    overflow: auto;
    font-size: 0.8rem;
    line-height: 1.55;
    white-space: pre-wrap;
    word-break: break-word;
    background: #fff;
    border: none;
    border-radius: 0;
    color: #2d3b36;
    box-sizing: border-box;
}

.paper-toolbar-below {
    flex-shrink: 0;
    margin-top: 12px;
    padding-top: 12px;
    border-top: 1px solid rgba(197, 221, 210, 0.85);
    display: flex;
    flex-direction: column;
    gap: 10px;
}

.paper-toolbar-actions {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 10px;
    width: 100%;
    min-width: 0;
}

.paper-url-row {
    display: flex;
    flex-direction: row;
    flex-wrap: wrap;
    align-items: center;
    gap: 8px;
    width: 100%;
    min-width: 0;
    box-sizing: border-box;
}

.paper-url-row--toolbar {
    margin-bottom: 0;
}

.paper-url-row .paper-input {
    flex: 1 1 200px;
    min-width: 0;
}

.paper-mini-label {
    font-size: 0.72rem;
    font-weight: 600;
    color: #3d5a50;
    margin-top: 6px;
}

.paper-mini-label:first-child {
    margin-top: 0;
}

.paper-input {
    width: 100%;
    max-width: 100%;
    padding: 8px 10px;
    border: 1px solid #b8d4c9;
    border-radius: 10px;
    font-size: 0.82rem;
    font-family: var(--font-sans);
    background: #fff;
    box-sizing: border-box;
}

.paper-file-native {
    display: none;
}

.paper-drop-zone {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 14px 16px;
    border: 2px dashed #8fbcb3;
    border-radius: 12px;
    cursor: pointer;
    background: #f3faf7;
    transition: border-color 0.15s ease, background 0.15s ease;
}

.paper-drop-zone--compact {
    flex: 1 1 160px;
    padding: 8px 12px;
    margin: 0;
    font-size: 0.78rem;
    gap: 8px;
}

.paper-drop-zone--active {
    border-color: #1a5c4a;
    background: #e6f4ef;
}

.paper-drop-icon {
    font-size: 1.35rem;
    color: #1a5c4a;
    flex-shrink: 0;
}

.paper-stat {
    font-size: 0.72rem;
    color: #5c7268;
    margin-left: auto;
}

.paper-details {
    border: 1px solid #c5ddd2;
    border-radius: 12px;
    background: rgba(255, 255, 255, 0.65);
    padding: 0 12px;
    box-shadow: 0 2px 10px rgba(15, 61, 50, 0.04);
}

.paper-details-summary {
    cursor: pointer;
    font-size: 0.78rem;
    font-weight: 600;
    color: #1a5c4a;
    padding: 10px 0;
    list-style: none;
}

.paper-details-summary::-webkit-details-marker {
    display: none;
}

.paper-details-body {
    padding-bottom: 12px;
    display: flex;
    flex-direction: column;
    gap: 6px;
}

.paper-paste-area {
    width: 100%;
    max-width: 100%;
    padding: 8px 10px;
    border: 1px solid #b8d4c9;
    border-radius: 10px;
    font-size: 0.8rem;
    resize: vertical;
    font-family: var(--font-sans);
    box-sizing: border-box;
}

.paper-paste-area--adaptive {
    min-height: 4.5rem;
    max-height: 14rem;
    field-sizing: content;
}

@supports not (field-sizing: content) {
    .paper-paste-area--adaptive {
        min-height: 5.5rem;
    }
}

.paper-merge-btn {
    align-self: flex-start;
}

.paper-err {
    font-size: 0.78rem;
    color: #b91c1c;
    margin: 0;
}

.paper-notes-scroll {
    flex: 1;
    min-height: 0;
    overflow-y: auto;
    display: flex;
    flex-direction: column;
    gap: 12px;
    padding-right: 2px;
}

.paper-notes-empty {
    font-size: 0.82rem;
    color: #5c7268;
    line-height: 1.55;
    padding: 8px 4px 12px;
}

.paper-nb-wrap {
    flex: 1;
    min-height: 0;
    display: flex;
    flex-direction: column;
    --paper-dim-accent: #1a5c4a;
}

.paper-nb-wrap.paper-tone--academic {
    --paper-dim-accent: #0f3d32;
}

.paper-nb-wrap.paper-tone--news {
    --paper-dim-accent: #0284c7;
}

.paper-nb-wrap.paper-tone--meeting {
    --paper-dim-accent: #b45309;
}

.paper-nb-wrap.paper-tone--product {
    --paper-dim-accent: #6d28d9;
}

.paper-nb-wrap.paper-tone--legal {
    --paper-dim-accent: #4338ca;
}

.paper-nb-wrap.paper-tone--default {
    --paper-dim-accent: #0d9488;
}

.paper-type-chip {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    flex-wrap: wrap;
    padding: 6px 12px;
    margin-bottom: 10px;
    border-radius: 999px;
    background: linear-gradient(90deg, rgba(15, 61, 50, 0.08), rgba(13, 148, 136, 0.08));
    border: 1px solid rgba(15, 61, 50, 0.12);
    font-size: 0.78rem;
}

.paper-type-chip__label {
    font-weight: 700;
    color: #3d5a50;
    letter-spacing: 0.04em;
}

.paper-type-chip__value {
    font-weight: 600;
    color: var(--paper-dim-accent);
}

.paper-nb-grid {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 14px;
    align-content: start;
    align-items: stretch;
}

.paper-nb-grid--dynamic {
    grid-template-columns: repeat(auto-fill, minmax(min(100%, 240px), 1fr));
}

.paper-nb-grid--fill {
    flex: 1;
    min-height: min(42vh, 520px);
    align-content: stretch;
}

.paper-nb-grid--dynamic .nb-block--dynamic {
    min-height: 9.5rem;
    border-top: 3px solid var(--paper-dim-accent);
}

.nb-block {
    display: flex;
    flex-direction: column;
    min-height: 0;
    height: 100%;
    background: #fff;
    border: 1px solid #c5ddd2;
    border-radius: 12px;
    padding: 12px 14px;
    box-shadow: 0 4px 14px rgba(15, 61, 50, 0.07);
    box-sizing: border-box;
}

.nb-block--span {
    grid-column: 1 / -1;
}

.nb-title {
    margin: 0 0 8px;
    font-size: 0.72rem;
    text-transform: none;
    letter-spacing: 0.04em;
    color: #1a5c4a;
    font-weight: 700;
    line-height: 1.35;
    word-break: break-word;
}

.nb-title.nb-title--mid {
    font-size: 0.66rem;
}

.nb-title.nb-title--long {
    font-size: 0.6rem;
    letter-spacing: 0.02em;
}

.nb-body {
    flex: 1 1 auto;
    margin: 0;
    font-size: 0.82rem;
    line-height: 1.55;
    color: #2d3b36;
    white-space: pre-wrap;
    word-break: break-word;
    min-height: 0;
}

.nb-cite-list {
    flex: 1 1 auto;
    margin: 0;
    padding-left: 1.1rem;
    font-size: 0.8rem;
    line-height: 1.55;
    color: #2d3b36;
    min-height: 0;
}

.nb-block--citations {
    background: linear-gradient(180deg, #fff 0%, #f3faf7 100%);
}

.paper-chat-actions {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin-top: 8px;
    width: 100%;
}

.paper-deep-btn {
    width: 100%;
    margin-top: 4px;
    background: linear-gradient(90deg, #0f3d32, #1e6b5e) !important;
    border: none !important;
    border-radius: 10px !important;
}

.paper-deep-btn--half {
    flex: 1 1 160px;
    width: auto !important;
    margin-top: 0 !important;
}

.paper-deep-btn--plain {
    background: #fff !important;
    color: #0f3d32 !important;
    border: 1.5px solid #1a5c4a !important;
}

.paper-gen-btn {
    width: 100%;
    flex-shrink: 0;
    margin-top: auto;
    background: linear-gradient(90deg, #166534, #15803d) !important;
    border: none !important;
    border-radius: 10px !important;
}

.paper-phase-msg {
    margin: 0;
    font-size: 0.8rem;
    color: #1e6b5e;
    text-align: center;
    line-height: 1.45;
}

.paper-openclaw-command {
    margin-top: 14px;
    padding: 12px 12px 10px;
    border-radius: 10px;
    background: linear-gradient(180deg, rgba(255, 255, 255, 0.92) 0%, rgba(236, 253, 245, 0.75) 100%);
    border: 1px solid rgba(26, 92, 74, 0.22);
}

.paper-openclaw-command__title {
    font-size: 0.72rem;
    font-weight: 700;
    color: #0f3d32;
    letter-spacing: 0.04em;
    margin-bottom: 8px;
}

.paper-openclaw-command__grid {
    display: flex;
    flex-direction: column;
    gap: 8px;
    width: 100%;
}

.paper-openclaw-command__btn {
    width: 100% !important;
    margin: 0 !important;
    justify-content: center !important;
    border-radius: 8px !important;
    border-color: #1a5c4a !important;
    color: #0f3d32 !important;
    background: #fff !important;
}

.paper-openclaw-command__btn:hover {
    background: rgba(26, 92, 74, 0.08) !important;
}

@media (max-width: 767px) {
    .paper-split {
        grid-template-columns: minmax(0, 1fr);
        min-height: auto;
    }

    .paper-col--source,
    .paper-col--notes {
        max-width: 100%;
        width: 100%;
    }

    .paper-preview-shell {
        height: min(58vh, 520px);
        max-height: min(58vh, 520px);
    }

    .paper-stat {
        margin-left: 0;
        width: 100%;
    }

    .paper-nb-grid {
        grid-template-columns: 1fr;
    }

    .paper-nb-grid--dynamic {
        grid-template-columns: 1fr;
    }

    .nb-block--span {
        grid-column: 1;
    }
}

.sr-only {
    position: absolute;
    width: 1px;
    height: 1px;
    padding: 0;
    margin: -1px;
    overflow: hidden;
    clip: rect(0, 0, 0, 0);
    white-space: nowrap;
    border-width: 0;
}

.mcc-k.mcc-k--label {
    margin: 0;
    cursor: pointer;
}

.nav-sub-link {
    display: block;
    width: 100%;
    text-align: left;
    padding: 8px 12px 10px 36px;
    margin: -4px 0 6px;
    font-size: 0.76rem;
    line-height: 1.35;
    color: #64748b;
    background: transparent;
    border: none;
    cursor: pointer;
    border-radius: 8px;
}

.nav-sub-link:hover {
    color: #334155;
    background: rgba(15, 23, 42, 0.05);
}

.impression-live-pulse {
    font-size: 0.68rem;
    color: #0369a1;
    margin: 8px 0 4px;
    line-height: 1.4;
    animation: impressionPulse 2s ease-in-out infinite;
}

@keyframes impressionPulse {
    0%,
    100% {
        opacity: 0.35;
    }
    50% {
        opacity: 1;
    }
}

.impression-lib-inline {
    margin-top: 8px;
    width: 100%;
}

.skill-smart-suggest {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 8px;
    padding: 8px 12px;
    margin-bottom: 8px;
    border-radius: 12px;
    background: linear-gradient(92deg, rgba(238, 242, 255, 0.96), rgba(254, 249, 231, 0.92));
    border: 1px solid rgba(129, 140, 248, 0.35);
    font-size: 0.78rem;
    color: #334155;
}

.skill-smart-suggest-text {
    flex: 1;
    min-width: 0;
}

.skill-suggest-fade-enter-active,
.skill-suggest-fade-leave-active {
    transition: opacity 0.38s ease;
}

.skill-suggest-fade-enter-from,
.skill-suggest-fade-leave-to {
    opacity: 0;
}

@media (max-width: 640px) {
    .chat-header-actions {
        width: 100%;
        margin-left: 0 !important;
        justify-content: flex-start;
    }

    .nav-sub-link {
        padding-left: 28px;
        font-size: 0.74rem;
    }
}

.demo-reset-link {
    display: block;
    margin-top: 6px;
    padding: 0;
    border: none;
    background: none;
    cursor: pointer;
    font-size: 0.65rem;
    letter-spacing: 0.02em;
    color: #94a3b8;
    opacity: 0.55;
    text-decoration: underline;
    text-underline-offset: 2px;
}

.demo-reset-link:hover {
    color: #64748b;
    opacity: 0.95;
}

.stat-subhint {
    margin-top: 4px;
    font-size: 0.68rem;
    line-height: 1.35;
    color: #64748b;
    font-weight: 500;
}

.chat-assistant-foot {
    margin-top: 8px;
    padding-top: 6px;
    border-top: 1px dashed #e5e7eb;
    display: flex;
    align-items: center;
    gap: 8px;
}

.chat-like-btn {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    padding: 4px 10px;
    font-size: 0.78rem;
    border-radius: 8px;
    border: 1px solid #e2e8f0;
    background: #f8fafc;
    color: #334155;
    cursor: pointer;
    transition: background 0.15s ease, border-color 0.15s ease;
}

.chat-like-btn:hover:not(:disabled) {
    border-color: #93c5fd;
    background: #eff6ff;
    color: #1d4ed8;
}

.chat-like-btn:disabled {
    opacity: 0.45;
    cursor: not-allowed;
}

.chat-like-label {
    font-size: 0.72rem;
}

.welcome-wizard-lead {
    margin: 0 0 14px;
    font-size: 0.82rem;
    line-height: 1.5;
    color: #475569;
}

.welcome-wizard-label {
    display: block;
    margin: 10px 0 4px;
    font-size: 0.78rem;
    font-weight: 600;
    color: #334155;
}

.welcome-wizard-input {
    width: 100%;
    box-sizing: border-box;
    padding: 8px 10px;
    border-radius: 8px;
    border: 1px solid #cbd5e1;
    font-size: 0.82rem;
}

.welcome-wizard-textarea {
    resize: vertical;
    min-height: 52px;
}
</style>
