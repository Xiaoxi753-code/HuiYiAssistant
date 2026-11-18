<script setup>
import { ref, computed, nextTick, watch, onBeforeUnmount, onMounted } from 'vue'
import { useRouter, useRoute } from 'vue-router'
import { ElMessage, ElCard, ElButton, ElMessageBox, ElDialog, ElDrawer, ElEmpty, ElButtonGroup, ElInput, ElDatePicker, ElDropdown, ElDropdownMenu, ElDropdownItem, ElCheckbox } from 'element-plus'
import 'element-plus/es/components/message/style/css'
import 'element-plus/es/components/card/style/css'
import 'element-plus/es/components/button/style/css'
import 'element-plus/es/components/message-box/style/css'
import 'element-plus/es/components/dialog/style/css'
import 'element-plus/es/components/drawer/style/css'
import 'element-plus/es/components/empty/style/css'
import 'element-plus/es/components/input/style/css'
import 'element-plus/es/components/date-picker/style/css'
import 'element-plus/es/components/dropdown/style/css'
import 'element-plus/es/components/checkbox/style/css'
import { useUserStore } from '../store/user.js'
import { useOpenclawContextStore } from '../store/openclawContext.js'
import { renderAsync } from 'docx-preview'
import Docxtemplater from 'docxtemplater'
import PizZip from 'pizzip'
import { saveAs } from 'file-saver'
import { marked } from 'marked'
import DOMPurify from 'dompurify'
import * as echarts from 'echarts'
import MeetingMinutesView from '../components/MeetingMinutesView.vue'
import { formatLocalDateTime, formatLocalDateTimeWithRelative } from '../utils/dateDisplay.js'
import { ElIcon } from 'element-plus'
import 'element-plus/es/components/icon/style/css'
import {
    AlarmClock,
    Briefcase,
    Calendar,
    ChatDotRound,
    ChatLineRound,
    Comment,
    Delete,
    Document,
    DocumentChecked,
    EditPen,
    Folder,
    Histogram,
    Key,
    Lightning,
    DArrowLeft,
    DArrowRight,
    Link,
    MagicStick,
    Medal,
    Menu,
    Microphone,
    Pointer,
    Promotion,
    Reading,
    Setting,
    TrendCharts,
    UserFilled,
    Plus,
} from '@element-plus/icons-vue'
import TodoBell from '../components/TodoBell.vue'

marked.setOptions({
    gfm: true,
    breaks: true,
})

const PURIFY_MD_OPTS = {
    ADD_TAGS: ['details', 'summary', 'div', 'pre'],
    ADD_ATTR: ['class', 'open'],
}

function escapeHtmlForThink(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
}

/**
 * 流式阶段：若最后一个 <think> 尚未闭合，截断到该开标签之前，避免半段标签进 Markdown。
 * 多段场景中只看「最后一次」开标签是否已配对闭合。
 */
function stripIncompleteRedactedThinking(raw) {
    const s = String(raw ?? '')
    const openRe = /<redacted_thinking\b[^>]*>/gi
    let last = null
    let m
    while ((m = openRe.exec(s)) !== null) {
        last = { idx: m.index, len: m[0].length }
    }
    if (!last) {
        return s
    }
    const tail = s.slice(last.idx + last.len)
    if (!/<\/redacted_thinking>/i.test(tail)) {
        return s.slice(0, last.idx)
    }
    return s
}

function transformRedactedThinkingBlocks(s) {
    return String(s).replace(/<redacted_thinking\b[^>]*>([\s\S]*?)<\/redacted_thinking>/gi, (_full, inner) => {
        const body = escapeHtmlForThink(inner)
        return `\n\n<details class="deep-think-box"><summary>深度思考过程</summary><div class="think-content"><pre class="think-pre">${body}</pre></div></details>\n\n`
    })
}

function preprocessAssistantMarkdownSource(text) {
    let s = stripIncompleteRedactedThinking(String(text ?? ''))
    s = transformRedactedThinkingBlocks(s)
    return s
}

/** 助手正文 Markdown → 安全 HTML（流式过程中多次调用，失败时回退转义） */
function renderMarkdown(text) {
    if (text == null || text === '') {
        return ''
    }
    try {
        const src = preprocessAssistantMarkdownSource(text)
        const rawHtml = marked.parse(src)
        return DOMPurify.sanitize(rawHtml, PURIFY_MD_OPTS)
    } catch {
        return DOMPurify.sanitize(escapeHtmlForThink(String(text)))
    }
}

/** 已是 HTML 的消息（欢迎语、历史回放）同样要过净化，避免直接注入原始 HTML */
function sanitizeHtml(html) {
    if (html == null || html === '') {
        return ''
    }
    try {
        return DOMPurify.sanitize(String(html), PURIFY_MD_OPTS)
    } catch {
        return escapeHtmlForThink(String(html))
    }
}

const router = useRouter()
const route = useRoute()
const userStore = useUserStore()
const openclawCtx = useOpenclawContextStore()
const currentTab = ref('overview')
const wordContainer = ref(null)
const isPreviewLoading = ref(false)
const previewError = ref('')
const extractedVariables = ref([])
const currentTemplateId = ref(null)
const isFilling = ref(false)
const fillError = ref('')
const fillStreamOutput = ref('')
const currentFile = ref(null)
const templateArrayBuffer = ref(null)
const lastWordBlob = ref(null)
const templateFillRecords = ref([])
const templateFolderUploading = ref(false)
const libraryActivePane = ref('templates')
const libraryItems = ref([])
const libraryLoading = ref(false)
const libraryUploading = ref(false)
const librarySearch = ref('')
const libraryFileInputRef = ref(null)
const libraryRenamingId = ref(null)
const libraryRenameDraft = ref('')
let fillAbortController = null
let syncTimer = null

const tabTitles = {
  overview: '工作台概览',
  templates: '智能模板引擎',
  library: '个人资料库',
  chat: '对话助手',
  meeting: '智能会议纪要',
  paper: '智能长文速读',
  ppt: 'PPT大纲生成',
  companion: '情感陪伴式对话',
  identity: '多身份切换',
  reminders: '主动提醒推送',
  apiConfig: 'API & 模型配置',
  webhookConfig: '外联推送配置',
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

/** 工作台：点赞/评分折线图 */
const ratingsTrendLoading = ref(false)
const ratingsTrendData = ref([])
const ratingsTrendChartRef = ref(null)
let ratingsChartInstance = null

function disposeRatingsChart() {
    if (ratingsChartInstance) {
        try {
            ratingsChartInstance.dispose()
        } catch {
            /* ignore */
        }
        ratingsChartInstance = null
    }
}

function mountRatingsChart() {
    disposeRatingsChart()
    const el = ratingsTrendChartRef.value
    if (!el || !ratingsTrendData.value.length) {
        return
    }
    ratingsChartInstance = echarts.init(el)
    const items = ratingsTrendData.value
    const dates = items.map((x) => x.date)
    const counts = items.map((x) => Number(x.count || 0))
    ratingsChartInstance.setOption({
        animationDuration: 320,
        grid: { left: 34, right: 8, top: 20, bottom: 28 },
        tooltip: { trigger: 'axis' },
        xAxis: {
            type: 'category',
            boundaryGap: false,
            data: dates,
            axisLabel: { fontSize: 10, color: '#64748b', rotate: dates.length > 12 ? 32 : 0 },
        },
        yAxis: {
            type: 'value',
            minInterval: 1,
            splitLine: { lineStyle: { color: 'rgba(148,163,184,0.35)' } },
            axisLabel: { fontSize: 10, color: '#64748b' },
        },
        series: [
            {
                name: '点赞/评分',
                type: 'line',
                smooth: true,
                showSymbol: dates.length <= 20,
                areaStyle: { color: 'rgba(124, 58, 237, 0.12)' },
                lineStyle: { width: 2.4, color: '#7c3aed' },
                itemStyle: { color: '#7c3aed' },
                data: counts,
            },
        ],
    })
}

function onRatingsChartResize() {
    if (ratingsChartInstance) {
        try {
            ratingsChartInstance.resize()
        } catch {
            /* ignore */
        }
    }
}

async function loadRatingsTrend() {
    ratingsTrendLoading.value = true
    try {
        const r = await fetch('/api/ratings/trend?days=90')
        const d = await r.json().catch(() => ({}))
        if (r.ok && d.success && Array.isArray(d.items)) {
            ratingsTrendData.value = d.items
        } else {
            ratingsTrendData.value = []
        }
    } catch {
        ratingsTrendData.value = []
    } finally {
        ratingsTrendLoading.value = false
    }
    await nextTick()
    mountRatingsChart()
}

const companionSettings = ref({
    persona: 'default',
    address: '',
    customTraits: '',
})

/** 主对话工作台身份（与「情感陪伴」人格分离）：personaId 随 /api/chat 上报，服务端写入 system */
const CHAT_PERSONA_STORAGE_KEY = 'openclaw_chat_persona_v1'
const BUILTIN_CHAT_PERSONAS = [
    { id: 'meeting_assistant', title: '会议助理', iconComponent: Calendar, desc: '干练结构化，议程与行动项优先' },
    { id: 'strict_mentor', title: '严厉导师', iconComponent: Medal, desc: '高标准反馈，先风险后步骤' },
    { id: 'geek_blaze', title: '爆裂极客', iconComponent: Lightning, desc: '快节奏技术梗，结论必须落地' },
    { id: 'workplace_helper', title: '职场助手', iconComponent: Briefcase, desc: '商务礼貌，结论先行' },
    { id: 'caring_partner', title: '知心大姐', iconComponent: ChatDotRound, desc: '先共情再建议，短句柔和' },
]

const chatActivePersonaId = ref('workplace_helper')
const chatCustomPersonas = ref([])
const customPersonaDialogOpen = ref(false)
const customPersonaForm = ref({
    name: '',
    emoji: '',
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
    customPersonaForm.value = { name: '', emoji: '', systemPrompt: '' }
    customPersonaDialogOpen.value = true
}

async function submitCustomPersona() {
    const name = String(customPersonaForm.value.name || '').trim().slice(0, 80)
    const emoji = String(customPersonaForm.value.emoji || '').trim().slice(0, 2) || (name ? name.charAt(0) : '自')
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

function resolveChatSessionMeta(chatMode = 'default') {
    const payload = resolvePersonaPayloadForChat()
    const builtin = BUILTIN_CHAT_PERSONAS.find((p) => p.id === payload.personaId)
    const custom = chatCustomPersonas.value.find((p) => p.id === payload.personaId)
    return {
        chatMode,
        personaId: payload.personaId || 'workplace_helper',
        personaLabel: payload.personaLabel || builtin?.title || custom?.name || '职场助手',
        personaEmoji: String(payload.personaEmoji || custom?.emoji || '').trim().slice(0, 8) || '',
        ...(activeChatSkill.value?.id
            ? { skillId: activeChatSkill.value.id, skillName: activeChatSkill.value.name || '' }
            : {}),
    }
}

/** 会话列表左侧：按 personaId 选择 Element Plus 图标（不依赖历史 emoji 字段） */
function chatSessionPersonaIcon(session) {
    const id = String(session?.personaId || '').trim()
    if (id) {
        const b = BUILTIN_CHAT_PERSONAS.find((p) => p.id === id)
        if (b?.iconComponent) {
            return b.iconComponent
        }
        if (id.startsWith('custom_')) {
            return UserFilled
        }
    }
    return Briefcase
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

/** 聊天流式请求状态：chatBusy 防并发发送，chatAbortController 供「停止生成」中断 */
const chatBusy = ref(false)
let chatAbortController = null

function stopChatStream() {
    if (!chatAbortController) {
        return
    }
    chatAbortController.abort()
    chatAbortController = null
}

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
const TODO_ALERT_STORAGE_KEY = 'openclaw_todo_alerted_v1'
let todoPollTimer = null
const dueTodoAlerts = ref([])
const dueTodoAlertVisible = ref(false)
const reminderForm = ref({
    title: '',
    remindAt: '',
    notes: '',
    repeatMode: 'once',
    repeatCustomDays: /** @type {number[]} */ ([]),
})
const reminderSaving = ref(false)

const WEEKDAY_CHECKBOXES = [
    { value: 1, label: '周一' },
    { value: 2, label: '周二' },
    { value: 3, label: '周三' },
    { value: 4, label: '周四' },
    { value: 5, label: '周五' },
    { value: 6, label: '周六' },
    { value: 7, label: '周日' },
]

const remindersBoardTodos = ref([])
const remindersBoardLoading = ref(false)
const remindersBoardTab = ref('pending')

const reminderTimeEditOpen = ref(false)
const reminderTimeEditRow = ref(null)
/** @type {import('vue').Ref<Date | null>} */
const reminderTimeEditDate = ref(null)
const reminderTimeEditSaving = ref(false)

function todoDrawerBucket(t) {
    if (Number(t.done) === 1) {
        return 2
    }
    const st = String(t.status || 'pending').toLowerCase()
    if (st === 'notified') {
        return 1
    }
    return 0
}

const sortedTodosDrawerList = computed(() => {
    const arr = [...todosDrawerList.value]
    arr.sort((a, b) => {
        const ba = todoDrawerBucket(a)
        const bb = todoDrawerBucket(b)
        if (ba !== bb) {
            return ba - bb
        }
        const ta = Date.parse(a.remindAt) || 0
        const tb = Date.parse(b.remindAt) || 0
        if (ba === 0) {
            return ta - tb
        }
        return tb - ta
    })
    return arr
})

function hasTodoBeenPushed(t) {
    const st = String(t.status || '').toLowerCase()
    if (st === 'notified') {
        return true
    }
    const w = String(t.pushedWecomAt || '').trim()
    const f = String(t.pushedFeishuAt || '').trim()
    return !!(w || f)
}

function sortRemindersForBoard(list, mode) {
    const arr = [...list]
    arr.sort((a, b) => {
        const ta = Date.parse(a.remindAt) || 0
        const tb = Date.parse(b.remindAt) || 0
        return mode === 'pending' ? ta - tb : tb - ta
    })
    return arr
}

const remindersPendingBoard = computed(() => {
    const list = remindersBoardTodos.value.filter((t) => Number(t.done) === 0 && !hasTodoBeenPushed(t))
    return sortRemindersForBoard(list, 'pending')
})

const remindersPushedBoard = computed(() => {
    const list = remindersBoardTodos.value.filter((t) => hasTodoBeenPushed(t))
    return sortRemindersForBoard(list, 'pushed')
})

const remindersRecurringBoard = computed(() => {
    const list = remindersBoardTodos.value.filter((t) => {
        if (Number(t.done) === 1) {
            return false
        }
        const rm = String(t.repeatMode || 'once').toLowerCase()
        return rm !== 'once'
    })
    return sortRemindersForBoard(list, 'pending')
})

const remindersBoardListActive = computed(() => {
    if (remindersBoardTab.value === 'pending') {
        return remindersPendingBoard.value
    }
    if (remindersBoardTab.value === 'pushed') {
        return remindersPushedBoard.value
    }
    return remindersRecurringBoard.value
})

/** 对话会话（左侧历史） */
const chatSessions = ref([])
const chatSessionsLoading = ref(false)
const currentChatSessionId = ref(null)
/** 窄屏时历史栏以抽屉覆盖显示，由汉堡按钮切换 */
const chatHistoryOverlayOpen = ref(false)
const chatInputTextareaRef = ref(null)

const CHAT_HISTORY_W_KEY = 'huiyi_chat_history_width_px'
const CHAT_HISTORY_C_KEY = 'huiyi_chat_history_collapsed'
const chatHistoryWidthPx = ref(260)
const chatHistoryCollapsed = ref(false)
const chatLayoutIsNarrow = ref(false)

let chatHistoryResizeActive = false
let chatHistoryResizeStartX = 0
let chatHistoryResizeStartW = 260

function updateChatLayoutNarrow() {
    if (typeof window === 'undefined') {
        return
    }
    chatLayoutIsNarrow.value = window.innerWidth <= 900
}

function onChatHistoryResizeMouseMove(e) {
    if (!chatHistoryResizeActive || chatHistoryCollapsed.value || chatLayoutIsNarrow.value) {
        return
    }
    const delta = e.clientX - chatHistoryResizeStartX
    const nw = Math.min(420, Math.max(200, chatHistoryResizeStartW + delta))
    chatHistoryWidthPx.value = nw
}

function onChatHistoryResizeMouseUp() {
    if (!chatHistoryResizeActive) {
        return
    }
    chatHistoryResizeActive = false
    document.removeEventListener('mousemove', onChatHistoryResizeMouseMove)
    document.removeEventListener('mouseup', onChatHistoryResizeMouseUp)
    document.body.style.cursor = ''
    document.body.style.userSelect = ''
    try {
        localStorage.setItem(CHAT_HISTORY_W_KEY, String(chatHistoryWidthPx.value))
    } catch {
        /* ignore */
    }
}

function startChatHistoryResize(e) {
    if (chatHistoryCollapsed.value || chatLayoutIsNarrow.value) {
        return
    }
    chatHistoryResizeActive = true
    chatHistoryResizeStartX = e.clientX
    chatHistoryResizeStartW = chatHistoryWidthPx.value
    document.addEventListener('mousemove', onChatHistoryResizeMouseMove)
    document.addEventListener('mouseup', onChatHistoryResizeMouseUp)
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
}

function toggleChatHistoryCollapse() {
    chatHistoryCollapsed.value = !chatHistoryCollapsed.value
    try {
        localStorage.setItem(CHAT_HISTORY_C_KEY, chatHistoryCollapsed.value ? '1' : '0')
    } catch {
        /* ignore */
    }
}

const chatHistorySidebarStyle = computed(() => {
    if (chatLayoutIsNarrow.value) {
        return {}
    }
    return {
        width: `${chatHistoryWidthPx.value}px`,
        flexShrink: 0,
    }
})

watch(chatLayoutIsNarrow, (narrow) => {
    if (narrow) {
        chatHistoryCollapsed.value = false
    }
})

/** 我的 Word 模板库 */
const myTemplates = ref([])
const myTemplatesLoading = ref(false)

const libraryFilteredItems = computed(() => {
    const q = String(librarySearch.value || '').trim().toLowerCase()
    if (!q) {
        return libraryItems.value
    }
    return libraryItems.value.filter((item) => `${item.title || ''} ${item.originalName || ''} ${item.folderPath || ''} ${(item.tags || []).join(' ')}`.toLowerCase().includes(q))
})

const libraryFolders = computed(() => {
    const map = new Map()
    for (const item of libraryItems.value) {
        const key = item.folderPath || '未分组'
        map.set(key, (map.get(key) || 0) + 1)
    }
    return Array.from(map.entries()).map(([name, count]) => ({ name, count }))
})

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

/** 实时 ASR（Qwen3）：仅存浏览器 localStorage，会议 WebSocket 握手时由前端发送给网关 */
const WHISPER_ASR_STORAGE_KEY = 'openclaw_whisper_asr_v1'
const DEFAULT_REALTIME_ASR_WS =
    'wss://dashscope.aliyuncs.com/api-ws/v1/realtime?model=qwen3-asr-flash-realtime-2026-02-10'
const whisperAsrRealtimeWsUrl = ref(DEFAULT_REALTIME_ASR_WS)
const whisperAsrRealtimeApiKey = ref('')

function loadWhisperAsrFromStorage() {
    if (typeof localStorage === 'undefined') {
        return
    }
    try {
        const raw = localStorage.getItem(WHISPER_ASR_STORAGE_KEY)
        if (!raw) {
            whisperAsrRealtimeWsUrl.value = DEFAULT_REALTIME_ASR_WS
            whisperAsrRealtimeApiKey.value = ''
            return
        }
        const o = JSON.parse(raw)
        whisperAsrRealtimeWsUrl.value =
            String(o.realtimeAsrWsUrl || o.asrWsUrl || '').trim() || DEFAULT_REALTIME_ASR_WS
        whisperAsrRealtimeApiKey.value = String(o.realtimeAsrApiKey ?? o.asrApiKey ?? '').trim()
    } catch {
        whisperAsrRealtimeWsUrl.value = DEFAULT_REALTIME_ASR_WS
        whisperAsrRealtimeApiKey.value = ''
    }
}

function saveWhisperAsrToStorage() {
    if (typeof localStorage === 'undefined') {
        ElMessage.warning('当前环境无法写入本地存储')
        return
    }
    const payload = {
        realtimeAsrWsUrl: String(whisperAsrRealtimeWsUrl.value || '').trim() || DEFAULT_REALTIME_ASR_WS,
        realtimeAsrApiKey: String(whisperAsrRealtimeApiKey.value || '').trim(),
    }
    localStorage.setItem(WHISPER_ASR_STORAGE_KEY, JSON.stringify(payload))
    ElMessage.success('实时 ASR 配置已保存到本机；会议纪要页将通过 WebSocket 使用该配置。')
}

/** 外联推送：企业微信 / 飞书 Webhook，存服务端 SQLite system_configs */
const webhookWecomUrl = ref('')
const webhookFeishuUrl = ref('')
const webhookConfigLoading = ref(false)
const webhookSaveLoading = ref(false)
const webhookTestWecomLoading = ref(false)
const webhookTestFeishuLoading = ref(false)

async function fetchWebhookConfigs() {
    webhookConfigLoading.value = true
    try {
        const r = await fetch('/api/configs/webhooks')
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        webhookWecomUrl.value = String(d.wecom_url || '')
        webhookFeishuUrl.value = String(d.feishu_url || '')
    } catch (e) {
        ElMessage.error(String(e?.message || e || '加载外联配置失败'))
    } finally {
        webhookConfigLoading.value = false
    }
}

async function saveWebhookConfigs() {
    webhookSaveLoading.value = true
    try {
        const r = await fetch('/api/configs/webhooks', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                wecom_url: String(webhookWecomUrl.value || '').trim(),
                feishu_url: String(webhookFeishuUrl.value || '').trim(),
            }),
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        ElMessage.success('保存成功')
    } catch (e) {
        ElMessage.error(String(e?.message || e || '保存失败'))
    } finally {
        webhookSaveLoading.value = false
    }
}

/** 外联：POST /api/configs/test-push，校验当前输入框中的 URL 是否可用 */
async function testWebhookPush(channel) {
    const ch = channel === 'feishu' ? 'feishu' : 'wecom'
    const url = ch === 'feishu' ? String(webhookFeishuUrl.value || '').trim() : String(webhookWecomUrl.value || '').trim()
    if (ch === 'feishu') {
        webhookTestFeishuLoading.value = true
    } else {
        webhookTestWecomLoading.value = true
    }
    try {
        const r = await fetch('/api/configs/test-push', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ url, channel: ch }),
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        ElMessage.success('会意助手链路测试消息已发送，请到群内查看')
    } catch (e) {
        ElMessage.error(String(e?.message || e || '测试推送失败'))
    } finally {
        webhookTestFeishuLoading.value = false
        webhookTestWecomLoading.value = false
    }
}

function getInitialChatMessages() {
    return [
        {
            role: 'assistant',
            isHtml: true,
            content:
                '你好！我是会意助手对话助手。我可以帮你：<br><br>' +
                '&#8226; <strong>直接操作工作台</strong> — 比如"帮我创建会议纪要"、"打开模板引擎"<br>' +
                '&#8226; <strong>智能问答</strong> — 回答关于功能使用的问题<br>' +
                '&#8226; <strong>快捷跳转</strong> — 说"去API配置"即可跳转到配置页<br><br>' +
                '试试输入：<em>"帮我创建会议纪要"</em> 或 <em>"打开模板引擎"</em>',
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
    () => route.query.openTodo,
    (v) => {
        if (v !== '1' && v !== 'true') {
            return
        }
        if (!userStore.user?.email) {
            return
        }
        openTodoDrawer()
        const q = { ...route.query }
        delete q.openTodo
        router.replace({ name: route.name, query: q, params: route.params })
    },
    { immediate: true },
)

watch(
    () => userStore.user?.email,
    () => {
        void loadDashboardStats()
        void loadTodosUpcoming()
        void loadMyTemplates()
        void loadTemplateFillRecords()
        void loadLibraryItems()
        void loadChatSessions()
        void loadTodoPendingCount()
    },
)

/** 自定义 Skill：打包提炼、@ 唤醒、随请求注入 system */
const savedSkills = ref([])
const activeChatSkill = ref(null)
const skillExtractLoading = ref(false)
/** 本条助手气泡是否已提交评分（会话内记忆） */
const likedAssistantTurnIds = ref({})
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

/** 长文速读 Agent 嗅探的潜在待办（容错：无字段或空数组） */
const paperPotentialTodosList = computed(() => {
    const s = paperSummary.value
    if (!s || typeof s !== 'object') {
        return []
    }
    const raw = s.potential_todos
    if (!Array.isArray(raw)) {
        return []
    }
    return raw
        .filter((x) => x && typeof x === 'object')
        .map((x) => ({
            title: String(x.title || x.content || x.task || '').trim(),
            remind_time: x.remind_time != null ? String(x.remind_time).trim() : '',
            notes: x.notes != null ? String(x.notes).trim() : '',
        }))
        .filter((x) => x.title.length > 0)
})

const paperPotentialTodoPick = ref(/** @type {Record<number, boolean>} */({}))
const paperPotentialSubmitting = ref(false)

watch(paperSummary, () => {
    paperPotentialTodoPick.value = {}
})

function setPaperPotentialPick(idx, checked) {
    paperPotentialTodoPick.value = { ...paperPotentialTodoPick.value, [idx]: !!checked }
}

/** 供后端合并进单条 system 的附加块（JSON，便于模型解析） */
function buildPaperSystemContextBlock() {
    const ct = paperSummaryContentType.value || '未分类'
    const dim = paperSummaryDimensions.value
    const payload = { content_type: ct, ...dim }
    return `【会意助手·材料结构化摘要】以下 JSON 为工作台提炼的键值要点，供本轮对话使用。\n${JSON.stringify(payload)}`
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

/** 将当前文献卡片同步到会意助手工作台全局上下文（Pinia），供其他 Tab / 对话读取 */
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

/** 指挥会意助手：快捷意图 → 对话助手 + 文献 system 上下文 + 低温严谨参数 */
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
    if (tab === 'overview') {
        nextTick(() => onRatingsChartResize())
    }
    if (tab === 'apiConfig') {
        loadWhisperAsrFromStorage()
        fetchModelConfigs()
    }
    if (tab === 'webhookConfig') {
        void fetchWebhookConfigs()
    }
    if (tab === 'chat') {
        loadUserImpressionProfile()
        fetchSavedSkills()
        void loadChatSessions()
        nextTick(() => syncChatComposerHeight())
    }
    if (tab === 'templates') {
        void loadMyTemplates()
        void loadTemplateFillRecords()
    }
    if (tab === 'library') {
        void loadMyTemplates()
        void loadTemplateFillRecords(30)
        void loadLibraryItems()
    }
    if (tab === 'reminders') {
        void loadRemindersBoard()
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

function getTodoAlertedIds() {
    try {
        const raw = localStorage.getItem(TODO_ALERT_STORAGE_KEY)
        const a = raw ? JSON.parse(raw) : []
        return new Set(Array.isArray(a) ? a.map((x) => Number(x)).filter((n) => Number.isFinite(n)) : [])
    } catch {
        return new Set()
    }
}

function markTodoAlerted(id) {
    const s = getTodoAlertedIds()
    s.add(Number(id))
    localStorage.setItem(TODO_ALERT_STORAGE_KEY, JSON.stringify([...s].slice(-400)))
}

function updateDueTodoAlerts() {
    const now = Date.now()
    const alerted = getTodoAlertedIds()
    const due = [...upcomingOverdue.value, ...upcomingSoon.value].filter((t) => {
        const st = String(t.status || 'pending').toLowerCase()
        if (st === 'notified') {
            return false
        }
        const dueAt = Date.parse(t.remindAt)
        if (!Number.isFinite(dueAt) || dueAt > now) {
            return false
        }
        return !alerted.has(t.id)
    })
    dueTodoAlerts.value = due
    dueTodoAlertVisible.value = due.length > 0
}

function formatTodoClock(remindAt) {
    const d = new Date(remindAt)
    if (!Number.isFinite(d.getTime())) {
        return '—'
    }
    return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`
}

function parseTodoRepeatCustomDays(raw) {
    if (Array.isArray(raw)) {
        return raw.map((x) => Number.parseInt(String(x), 10)).filter((n) => Number.isFinite(n) && n >= 1 && n <= 7)
    }
    try {
        const a = JSON.parse(String(raw || '[]'))
        return Array.isArray(a)
            ? a.map((x) => Number.parseInt(String(x), 10)).filter((n) => Number.isFinite(n) && n >= 1 && n <= 7)
            : []
    } catch {
        return []
    }
}

function todoRepeatTag(mode, repeatCustomDaysRaw) {
    const m = String(mode || 'once').toLowerCase()
    if (m === 'daily') {
        return { label: '每天', kind: 'daily' }
    }
    if (m === 'workday') {
        return { label: '工作日', kind: 'workday' }
    }
    if (m === 'custom_days') {
        const days = parseTodoRepeatCustomDays(repeatCustomDaysRaw)
        const map = { 1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '日' }
        const label =
            days.length > 0 ? `每周${days.map((d) => `周${map[d] || d}`).join('、')}` : '自定义'
        return { label, kind: 'custom' }
    }
    return { label: '单次', kind: 'once' }
}

function getTodayYmdShanghaiClient() {
    try {
        const parts = new Intl.DateTimeFormat('en-CA', {
            timeZone: 'Asia/Shanghai',
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
        }).formatToParts(new Date())
        const y = parts.find((p) => p.type === 'year')?.value
        const m = parts.find((p) => p.type === 'month')?.value
        const d = parts.find((p) => p.type === 'day')?.value
        return y && m && d ? `${y}-${m}-${d}` : ''
    } catch {
        return ''
    }
}

function todoStatusDisplay(t) {
    if (Number(t.done) === 1) {
        return { label: '已完成', kind: 'done' }
    }
    const rm = String(t.repeatMode || 'once').toLowerCase()
    if (rm !== 'once') {
        const today = getTodayYmdShanghaiClient()
        const lcd = String(t.lastCompletedDate || '').trim()
        if (today && lcd && lcd === today) {
            return { label: '✓ 今日已处理', kind: 'repeat_cleared' }
        }
        return { label: '🔁 循环提醒', kind: 'repeat' }
    }
    const st = String(t.status || 'pending').toLowerCase()
    if (st === 'notified') {
        return { label: '已推送', kind: 'notified' }
    }
    return { label: '未完成', kind: 'pending' }
}

function isTodoPastDue(t) {
    if (Number(t.done) === 1) {
        return false
    }
    const st = String(t.status || 'pending').toLowerCase()
    if (st === 'notified') {
        return false
    }
    const dueAt = Date.parse(t.remindAt)
    return Number.isFinite(dueAt) && dueAt <= Date.now()
}

function dismissDueTodoAlerts() {
    for (const t of dueTodoAlerts.value) {
        markTodoAlerted(t.id)
    }
    dueTodoAlerts.value = []
    dueTodoAlertVisible.value = false
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
    updateDueTodoAlerts()
}

async function loadTodoPendingCount() {
    const email = userStore.user?.email
    if (!email) {
        todoPendingCount.value = 0
        return
    }
    try {
        const r = await fetch(`/api/todos/attention-count?email=${encodeURIComponent(email)}`)
        const d = await r.json().catch(() => ({}))
        if (r.ok && d.success) {
            todoPendingCount.value = Number(d.count || 0)
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

async function loadRemindersBoard() {
    const email = userStore.user?.email
    if (!email) {
        remindersBoardTodos.value = []
        return
    }
    remindersBoardLoading.value = true
    try {
        const r = await fetch(`/api/todos?email=${encodeURIComponent(email)}`)
        const d = await r.json().catch(() => ({}))
        if (r.ok && d.success) {
            remindersBoardTodos.value = d.items || []
        }
    } catch {
        remindersBoardTodos.value = []
    } finally {
        remindersBoardLoading.value = false
    }
}

async function addSelectedPaperPotentialTodos() {
    const email = userStore.user?.email
    if (!email) {
        ElMessage.warning('请先登录后再加入待办')
        return
    }
    const list = paperPotentialTodosList.value
    const pickedIdx = list.map((_, i) => i).filter((i) => paperPotentialTodoPick.value[i])
    if (!pickedIdx.length) {
        ElMessage.warning('请先勾选要加入的待办')
        return
    }
    paperPotentialSubmitting.value = true
    const defaultRemind = () => new Date(Date.now() + 2 * 3600000).toISOString()
    let ok = 0
    try {
        for (const i of pickedIdx) {
            const item = list[i]
            if (!item) {
                continue
            }
            let remindAt = item.remind_time
            if (!remindAt || Number.isNaN(Date.parse(remindAt))) {
                remindAt = defaultRemind()
            } else {
                const parsed = Date.parse(remindAt)
                remindAt = Number.isFinite(parsed) ? new Date(parsed).toISOString() : defaultRemind()
            }
            const r = await fetch('/api/todos', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    userEmail: email,
                    title: item.title.slice(0, 500),
                    remindAt,
                    notes: item.notes ? `[长文速读] ${item.notes}`.slice(0, 2000) : '[长文速读] 来自智能长文速读嗅探',
                    source: 'paper_summary',
                    repeatMode: 'once',
                }),
            })
            const d = await r.json().catch(() => ({}))
            if (r.ok && d.success) {
                ok += 1
            }
        }
        ElMessage.success(`已加入 ${ok} 条待办并进入主动提醒调度`)
        paperPotentialTodoPick.value = {}
        void loadTodosUpcoming()
        void loadTodoPendingCount()
        void loadRemindersBoard()
        void loadTodosForDrawer()
        updateDueTodoAlerts()
    } catch (e) {
        ElMessage.error(String(e?.message || e || '加入失败'))
    } finally {
        paperPotentialSubmitting.value = false
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
        if (d.mode === 'repeat') {
            ElMessage.success('已标记本日完成；循环任务将在明日重新计入提醒数字')
        } else {
            ElMessage.success('已完成')
        }
        void loadTodosForDrawer()
        void loadRemindersBoard()
        void loadTodoPendingCount()
        void loadTodosUpcoming()
        updateDueTodoAlerts()
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
        ElMessage.success('已删除')
        void loadTodosForDrawer()
        void loadRemindersBoard()
        void loadTodoPendingCount()
        void loadTodosUpcoming()
        updateDueTodoAlerts()
    } catch (e) {
        ElMessage.error(String(e?.message || e || '删除失败'))
    }
}

async function confirmDeleteTodoRow(row) {
    try {
        await ElMessageBox.confirm('确定要删除该待办吗？删除后不可恢复。', '删除待办', {
            confirmButtonText: '删除',
            cancelButtonText: '取消',
            type: 'warning',
        })
    } catch {
        return
    }
    await deleteTodoRow(row)
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
        if (String(t.status || 'pending').toLowerCase() === 'notified') {
            continue
        }
        const due = Date.parse(t.remindAt)
        if (!Number.isFinite(due) || due > now) {
            continue
        }
        if (notified.has(t.id)) {
            continue
        }
        try {
            new Notification('会意助手待办提醒', { body: t.title })
            markTodoNotified(t.id)
        } catch {
            /* ignore */
        }
    }
    updateDueTodoAlerts()
}

function toggleReminderCustomDay(day) {
    const arr = [...(reminderForm.value.repeatCustomDays || [])]
    const i = arr.indexOf(day)
    if (i >= 0) {
        arr.splice(i, 1)
    } else {
        arr.push(day)
    }
    arr.sort((a, b) => a - b)
    reminderForm.value.repeatCustomDays = arr
}

function openReminderTimeEdit(row) {
    if (!row?.id) {
        ElMessage.error('待办数据异常，请刷新后重试')
        return
    }
    reminderTimeEditRow.value = row
    const d = new Date(row.remindAt || '')
    if (Number.isFinite(d.getTime())) {
        reminderTimeEditDate.value = d
    } else {
        /** 没有提醒时间的待办：默认给 5 分钟后，避免点按钮毫无反应 */
        reminderTimeEditDate.value = new Date(Date.now() + 5 * 60 * 1000)
        ElMessage.info('这条待办还没有提醒时间，已默认填为 5 分钟后')
    }
    reminderTimeEditOpen.value = true
}

async function saveReminderTimeEdit() {
    const email = userStore.user?.email
    const row = reminderTimeEditRow.value
    if (!email || !row?.id) {
        return
    }
    const dt = reminderTimeEditDate.value
    if (!dt || !Number.isFinite(dt.getTime())) {
        ElMessage.warning('请选择时间')
        return
    }
    const remindAt = dt.toISOString()
    reminderTimeEditSaving.value = true
    try {
        const r = await fetch(`/api/todos/${row.id}/remind-at`, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email, remindAt }),
        })
        const resJson = await r.json().catch(() => ({}))
        if (!r.ok || !resJson.success) {
            throw new Error(resJson.message || `HTTP ${r.status}`)
        }
        ElMessage.success('已更新提醒时间')
        reminderTimeEditOpen.value = false
        void loadRemindersBoard()
        void loadTodosUpcoming()
        void loadTodoPendingCount()
        void loadTodosForDrawer()
    } catch (e) {
        ElMessage.error(String(e?.message || e || '更新失败'))
    } finally {
        reminderTimeEditSaving.value = false
    }
}

async function submitReminder() {
    const email = userStore.user?.email
    if (!email) {
        ElMessage.warning('请先登录后新增提醒')
        return
    }
    const title = String(reminderForm.value.title || '').trim()
    /** datetime-local 给的是无时区的本地时间，统一转成带时区的 ISO，避免服务端按自身时区解释 */
    const remindAtRaw = String(reminderForm.value.remindAt || '').trim()
    const remindAtDate = new Date(remindAtRaw)
    const remindAt = Number.isFinite(remindAtDate.getTime()) ? remindAtDate.toISOString() : ''
    const notes = String(reminderForm.value.notes || '').trim()
    const repeatMode = reminderForm.value.repeatMode || 'once'
    if (!title) {
        ElMessage.warning('请填写待办标题')
        return
    }
    if (!remindAt) {
        ElMessage.warning('请选择提醒时间')
        return
    }
    if (repeatMode === 'custom_days' && !(reminderForm.value.repeatCustomDays || []).length) {
        ElMessage.warning('自定义重复请至少选择一个星期')
        return
    }
    reminderSaving.value = true
    try {
        const payload = {
            userEmail: email,
            title,
            remindAt,
            notes,
            source: 'manual',
            repeatMode,
        }
        if (repeatMode === 'custom_days') {
            payload.repeatCustomDays = reminderForm.value.repeatCustomDays || []
        }
        const r = await fetch('/api/todos', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        reminderForm.value = {
            title: '',
            remindAt: '',
            notes: '',
            repeatMode: 'once',
            repeatCustomDays: [],
        }
        ElMessage.success('已新增提醒')
        void loadTodosUpcoming()
        void loadTodoPendingCount()
        void loadRemindersBoard()
    } catch (e) {
        ElMessage.error(String(e?.message || e || '新增提醒失败'))
    } finally {
        reminderSaving.value = false
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

async function loadTemplateFillRecords(limit = 8) {
    const email = userStore.user?.email
    if (!email) {
        templateFillRecords.value = []
        return
    }
    try {
        const r = await fetch(`/api/templates/fill-records?email=${encodeURIComponent(email)}&limit=${encodeURIComponent(String(limit))}`)
        const d = await r.json().catch(() => ({}))
        if (r.ok && d.success) {
            templateFillRecords.value = d.items || []
        }
    } catch {
        templateFillRecords.value = []
    }
}


async function loadLibraryItems() {
    const email = userStore.user?.email
    if (!email) {
        libraryItems.value = []
        return
    }
    libraryLoading.value = true
    try {
        const r = await fetch(`/api/library/items?email=${encodeURIComponent(email)}`)
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        libraryItems.value = d.items || []
    } catch (e) {
        libraryItems.value = []
        ElMessage.error(String(e?.message || e || '资料库加载失败'))
    } finally {
        libraryLoading.value = false
    }
}

function triggerLibraryUpload() {
    if (!userStore.user?.email) {
        ElMessage.warning('请先登录后上传资料')
        return
    }
    libraryFileInputRef.value?.click?.()
}

async function handleLibraryFileUpload(event) {
    const file = event?.target?.files?.[0]
    if (!file) {
        return
    }
    const email = userStore.user?.email
    if (!email) {
        ElMessage.warning('请先登录')
        return
    }
    libraryUploading.value = true
    try {
        const fd = new FormData()
        fd.append('file', file)
        fd.append('userEmail', email)
        fd.append('folderPath', '手动上传')
        const r = await fetch('/api/library/upload', { method: 'POST', body: fd })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        ElMessage.success('资料已上传并入库')
        void loadLibraryItems()
    } catch (e) {
        ElMessage.error(String(e?.message || e || '资料上传失败'))
    } finally {
        libraryUploading.value = false
        if (event?.target) {
            event.target.value = ''
        }
    }
}

function downloadLibraryItem(row) {
    const email = userStore.user?.email
    if (!email || !row?.id) {
        ElMessage.warning('请先登录')
        return
    }
    if (!row.canDownload) {
        ElMessage.info('这条记录是文本纪要/摘要，没有原始文件可下载')
        return
    }
    window.open(`/api/library/items/${row.id}/file?email=${encodeURIComponent(email)}`, '_blank')
}

function startLibraryRename(row) {
    libraryRenamingId.value = row.id
    libraryRenameDraft.value = row.title || ''
}

function cancelLibraryRename() {
    libraryRenamingId.value = null
}

async function commitLibraryRename(row) {
    if (libraryRenamingId.value !== row.id) {
        return
    }
    const draft = String(libraryRenameDraft.value || '').trim().slice(0, 240)
    const prev = String(row.title || '').trim()
    if (!draft || draft === prev) {
        cancelLibraryRename()
        return
    }
    const email = userStore.user?.email
    if (!email) {
        ElMessage.warning('请先登录')
        cancelLibraryRename()
        return
    }
    try {
        const r = await fetch(`/api/files/${row.id}/rename`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email, newName: draft }),
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        ElMessage.success('已重命名')
        cancelLibraryRename()
        await loadLibraryItems()
    } catch (e) {
        ElMessage.error(String(e?.message || e || '重命名失败'))
    }
}

async function editLibraryItem(row) {
    const email = userStore.user?.email
    if (!email || !row?.id) {
        ElMessage.warning('请先登录')
        return
    }
    try {
        const { value: folder } = await ElMessageBox.prompt('文件夹路径（可用 / 分层，例如：项目A/会议）', '整理资料', {
            inputValue: row.folderPath || '',
        })
        const { value: tagsInput } = await ElMessageBox.prompt('标签（用逗号分隔）', '资料标签', {
            inputValue: Array.isArray(row.tags) ? row.tags.join(', ') : '',
        })
        const tags = String(tagsInput || '').split(/[，,;；\s]+/).map((x) => x.trim()).filter(Boolean)
        const r = await fetch(`/api/library/items/${row.id}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email, title: row.title, folderPath: String(folder || '').trim(), tags }),
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        ElMessage.success('资料整理信息已更新')
        void loadLibraryItems()
    } catch (e) {
        if (e !== 'cancel') {
            ElMessage.error(String(e?.message || e || '整理失败'))
        }
    }
}

async function deleteLibraryItem(row) {
    const email = userStore.user?.email
    if (!email || !row?.id) {
        return
    }
    try {
        await ElMessageBox.confirm(`确定删除资料「${row.title || row.originalName || ''}」？`, '删除资料', { type: 'warning' })
    } catch {
        return
    }
    try {
        const r = await fetch(`/api/library/items/${row.id}?email=${encodeURIComponent(email)}`, { method: 'DELETE' })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        ElMessage.success('已删除资料')
        void loadLibraryItems()
    } catch (e) {
        ElMessage.error(String(e?.message || e || '删除失败'))
    }
}

async function agentOrganizeLibrary() {
    const email = userStore.user?.email
    if (!email) {
        ElMessage.warning('请先登录')
        return
    }
    libraryLoading.value = true
    try {
        const r = await fetch('/api/library/agent-organize', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email }),
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        libraryItems.value = d.items || []
        ElMessage.success(`Agent 已整理 ${d.updated || 0} 条资料`)
    } catch (e) {
        ElMessage.error(String(e?.message || e || 'Agent 整理失败'))
    } finally {
        libraryLoading.value = false
    }
}

function libraryItemTypeLabel(type) {
    const t = String(type || '')
    if (t === 'meeting') return '会议纪要'
    if (t === 'summary') return '摘要记录'
    return '文件资料'
}

function libraryItemTypeIconComponent(type) {
    const t = String(type || '')
    if (t === 'meeting') {
        return Microphone
    }
    if (t === 'summary') {
        return DocumentChecked
    }
    return Document
}

function downloadTemplateRecord(row) {
    const email = userStore.user?.email
    if (!email || !row?.id) {
        ElMessage.warning('请先登录')
        return
    }
    window.open(`/api/templates/fill-records/${row.id}/file?email=${encodeURIComponent(email)}`, '_blank')
}

async function applyTemplateRecordVariables(row) {
    if (!currentFile.value || !Array.isArray(extractedVariables.value) || extractedVariables.value.length === 0) {
        ElMessage.warning('请先打开对应模板，再复用历史变量')
        return
    }
    const variables = row?.variables && typeof row.variables === 'object' ? row.variables : {}
    Object.entries(variables).forEach(([name, value]) => updateVariableValue(name, String(value ?? '')))
    await syncToWord({ render: true })
    ElMessage.success('已复用该次生成的变量')
}

async function editTemplateOrganization(row) {
    const email = userStore.user?.email
    if (!email || !row?.id) {
        ElMessage.warning('请先登录')
        return
    }
    try {
        const { value: folder } = await ElMessageBox.prompt('文件夹路径（可用 / 分层，例如：合同/采购）', '整理模板', {
            inputValue: row.folderPath || '',
        })
        const { value: tagsInput } = await ElMessageBox.prompt('标签（用逗号分隔）', '模板标签', {
            inputValue: Array.isArray(row.tags) ? row.tags.join(', ') : '',
        })
        const tags = String(tagsInput || '').split(/[，,;；\s]+/).map((x) => x.trim()).filter(Boolean)
        const name = row.originalName || row.original_name || 'template.docx'
        const r = await fetch(`/api/templates/${row.id}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ email, original_name: name, folder_path: String(folder || '').trim(), tags }),
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        ElMessage.success('模板整理信息已更新')
        void loadMyTemplates()
    } catch (e) {
        if (e !== 'cancel') {
            ElMessage.error(String(e?.message || e || '整理失败'))
        }
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
            body: JSON.stringify({ email, original_name: name, folder_path: row.folderPath || '', tags: row.tags || [] }),
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
        window.addEventListener('resize', onRatingsChartResize, { passive: true })
        loadWhisperAsrFromStorage()
        loadChatPersonaLocalStorage()
        updateChatLayoutNarrow()
        window.addEventListener('resize', updateChatLayoutNarrow, { passive: true })
        try {
            chatHistoryCollapsed.value = localStorage.getItem(CHAT_HISTORY_C_KEY) === '1'
            const rw = Number.parseInt(String(localStorage.getItem(CHAT_HISTORY_W_KEY) || ''), 10)
            if (Number.isFinite(rw) && rw >= 200 && rw <= 420) {
                chatHistoryWidthPx.value = rw
            }
        } catch {
            /* ignore */
        }
    }
    void fetchWebhookConfigs()
    fetchSavedSkills()
    void loadDashboardStats()
    void loadRatingsTrend()
    void loadUserImpressionProfile()
    void loadTodosUpcoming()
    void loadTodoPendingCount()
    void loadChatSessions()
    void requestTodoNotificationPermission()
    todoPollTimer = window.setInterval(() => {
        void loadTodosUpcoming()
            .then(() => checkDueTodoNotifications())
            .then(() => loadTodoPendingCount())
    }, 10000)
    nextTick(() => {
        void maybeOpenWelcomeWizard()
    })
})

/** 统一提示：此前只写 console，用户侧完全看不到反馈 */
function showToast(type, msg) {
  const text = String(msg ?? '').trim()
  if (!text) {
    return
  }
  if (type === 'success') {
    ElMessage.success(text)
  } else if (type === 'error') {
    ElMessage.error(text)
  } else if (type === 'warning') {
    ElMessage.warning(text)
  } else {
    ElMessage.info(text)
  }
}

function handleLogout() {
  userStore.clearUser()
  router.push('/login')
}

function clearChat() {
    currentChatSessionId.value = null
    chatMessages.value = getInitialChatMessages()
    likedAssistantTurnIds.value = {}
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

/** 从当前对话提炼 Skill（顶栏与气泡底部「打包 Skill」共用） */
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

/** 对话助手 AI 气泡：仅正常完成流式输出后展示点赞入口 */
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

function assistantTurnKey(message) {
    return String(message?._clientTurnId || '')
}

function isAssistantLiked(message) {
    const k = assistantTurnKey(message)
    return Boolean(k && likedAssistantTurnIds.value[k])
}

async function onAssistantThumbsUp(message) {
    const k = assistantTurnKey(message)
    if (!k) {
        return
    }
    if (!userStore.user?.email) {
        ElMessage.warning('请先登录后再评分')
        return
    }
    if (likedAssistantTurnIds.value[k]) {
        return
    }
    try {
        const r = await fetch('/api/ratings', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ score: 1 }),
        })
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        likedAssistantTurnIds.value = { ...likedAssistantTurnIds.value, [k]: true }
        ElMessage.success('感谢您的评分！')
        void loadRatingsTrend()
    } catch (e) {
        ElMessage.error(String(e?.message || e || '评分失败'))
    }
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
        ElMessage.success('已保存，会意助手会更懂你的语境')
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
    if (e.key !== 'Enter' || e.shiftKey) {
        return
    }
    /** 中文输入法选词中的回车只用于确认候选词，不能当发送（isComposing / 229 为兼容旧 IME） */
    if (e.isComposing || e.keyCode === 229) {
        return
    }
    e.preventDefault()
    sendChatMessage()
}

function onCompanionComposerKeydown(e) {
    if (e.key !== 'Enter' || e.shiftKey) {
        return
    }
    if (e.isComposing || e.keyCode === 229) {
        return
    }
    e.preventDefault()
    sendChatMessage(undefined, { chatMode: CHAT_MODE_COMPANION })
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
        _clientTurnId: row.id != null ? `db-msg-${row.id}` : undefined,
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
            body: JSON.stringify({ email, title: '新对话', ...resolveChatSessionMeta('default') }),
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
                chatMode: 'default',
                sessionMeta: resolveChatSessionMeta('default'),
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
    /** 切换会话前先掐断正在生成的回复，否则上一轮的流会继续往新会话里追加内容 */
    stopChatStream()
    try {
        const r = await fetch(`/api/chat/sessions/${sessionId}/messages?email=${encodeURIComponent(email)}`)
        const d = await r.json().catch(() => ({}))
        if (!r.ok || !d.success) {
            throw new Error(d.message || `HTTP ${r.status}`)
        }
        currentChatSessionId.value = sessionId
        const sessionMeta = chatSessions.value.find((s) => s.id === sessionId)
        if (sessionMeta?.personaId) {
            chatActivePersonaId.value = sessionMeta.personaId
        }
        if (sessionMeta?.skillId) {
            activeChatSkill.value = { id: Number(sessionMeta.skillId) || sessionMeta.skillId, name: sessionMeta.skillName || '历史 Skill' }
        } else {
            activeChatSkill.value = null
        }
        const items = d.items || []
        if (items.length === 0) {
            chatMessages.value = getInitialChatMessages()
        } else {
            chatMessages.value = items.map(mapDbRowToChatMessage)
        }
        chatLoadedPaperLabel.value = ''
        await nextTick()
        scheduleScrollChatToBottom(false)
    } catch (e) {
        ElMessage.error(String(e?.message || e || '加载会话失败'))
    }
}

/**
 * 新建对话：只在本地开一张空白草稿，不立刻落库。
 * 会话记录由首次发送时的 ensureChatSessionBeforeSend 创建，
 * 否则每点一次「新对话」就往历史里堆一条空记录。
 */
async function createNewChatSession() {
    const email = userStore.user?.email
    if (!email) {
        ElMessage.warning('请先登录后再使用多会话')
        return
    }
    stopChatStream()
    currentChatSessionId.value = null
    chatMessages.value = getInitialChatMessages()
    chatInput.value = ''
    chatLoadedPaperLabel.value = ''
    activeChatSkill.value = null
    await nextTick()
    syncChatComposerHeight()
    scheduleScrollChatToBottom(false)
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
    if (chatBusy.value) {
        ElMessage.warning('上一条回复还在生成，请稍候，或点击「停止」后再发送')
        return
    }
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
        /** 先占锁再 await，否则两次极快的回车都能穿过上面的忙闲判断 */
        chatBusy.value = true
        try {
            await ensureChatSessionBeforeSend()
        } catch (e) {
            chatBusy.value = false
            ElMessage.error(String(e?.message || e || '会话创建失败，请稍后重试'))
            return
        }
    } else {
        chatBusy.value = true
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
        _clientTurnId: `turn-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`,
    }
    const assistantIndex = messagesRef.value.length
    messagesRef.value.push(assistantMessage)

    await nextTick()
    scheduleScrollChatToBottom(isCompanion)

    chatAbortController = new AbortController()

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
            signal: chatAbortController?.signal,
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
    } finally {
        chatBusy.value = false
        chatAbortController = null
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
    } else if (name.endsWith('.doc')) {
        paperFileKind.value = 'doc'
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
        paperTitle.value = file.name.replace(/\.(pdf|doc|docx)$/i, '')
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
        ElMessage.warning('请先有正文材料')
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
    document.removeEventListener('mousemove', onChatHistoryResizeMouseMove)
    document.removeEventListener('mouseup', onChatHistoryResizeMouseUp)
    if (typeof window !== 'undefined') {
        window.removeEventListener('resize', onRatingsChartResize)
        window.removeEventListener('resize', updateChatLayoutNarrow)
        disposeRatingsChart()
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
    const map = { aliyun: '阿里云百炼', openai: 'OpenAI 兼容', openclaw: '会意助手本地网关' }
    return map[type] || type
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

function createDocxTagParser() {
    return (tag) => {
        const key = String(tag ?? '').trim()
        return {
            get(scope) {
                if (!key) {
                    return undefined
                }
                if (scope && Object.prototype.hasOwnProperty.call(scope, key)) {
                    return scope[key]
                }
                return undefined
            },
        }
    }
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

function normalizeTemplateZipForDoubleBraces(zip) {
    const files = ['word/document.xml']
    zip.file(/word\/(header|footer)\d+\.xml/).forEach((file) => files.push(file.name))
    files.forEach((name) => {
        const file = zip.file(name)
        if (!file) {
            return
        }
        const xml = file.asText()
        const normalized = xml.replace(/\{\{\s*([^{}]+?)\s*\}\}/g, (_m, key) => `{${String(key).trim()}}`)
        if (normalized !== xml) {
            zip.file(name, normalized)
        }
    })
}

async function syncToWord({ render = true } = {}) {
    if (!templateArrayBuffer.value) {
        return
    }

    const zip = new PizZip(templateArrayBuffer.value)
    normalizeTemplateZipForDoubleBraces(zip)
    const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        delimiters: { start: '{', end: '}' },
        parser: createDocxTagParser(),
        nullGetter: (part) => `{{${part.value}}}`,
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

    const found = new Set()
    const addName = (raw) => {
        const value = String(raw || '').replace(/[{}]/g, '').trim()
        if (value && value.length <= 120 && !/[\r\n]/.test(value)) {
            found.add(value)
        }
    }
    let match
    const doubleRegex = /\{\{\s*([^{}]+?)\s*\}\}/g
    while ((match = doubleRegex.exec(text)) !== null) {
        addName(match[1])
    }
    const singleRegex = /(?<!\{)\{\s*([^{}]+?)\s*\}(?!\})/g
    while ((match = singleRegex.exec(text)) !== null) {
        addName(match[1])
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
        scheduleSyncToWord()
        return
    }

    extractedVariables.value = [
        ...extractedVariables.value,
        { name, value, isTyping: false },
    ]
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
    fillStreamOutput.value = ''
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
            buffer += chunk
            const parts = buffer.split('\n\n')
            buffer = parts.pop() || ''

            for (const part of parts) {
                if (!part.trim()) {
                    continue
                }

                const { event, data } = parseSseChunk(part)
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
                        if (typeof payload === 'string') {
                            payload = JSON.parse(payload)
                        }
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
                    fillStreamOutput.value += payload.text
                    if (fillStreamOutput.value.length > 8000) {
                        fillStreamOutput.value = fillStreamOutput.value.slice(-8000)
                    }
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

async function uploadTemplateFolderFiles(files) {
    const docxFiles = Array.from(files || []).filter((file) => String(file.name || '').toLowerCase().endsWith('.docx'))
    if (docxFiles.length === 0) {
        ElMessage.warning('文件夹中没有可上传的 .docx 模板')
        return
    }
    const formData = new FormData()
    docxFiles.forEach((file) => {
        formData.append('files', file, file.webkitRelativePath || file.name)
    })
    const em = userStore.user?.email
    if (em) {
        formData.append('userEmail', em)
    }
    templateFolderUploading.value = true
    try {
        const response = await fetch('/api/templates/upload-folder', {
            method: 'POST',
            body: formData,
        })
        const data = await response.json().catch(() => ({}))
        if (!response.ok || !data.success) {
            throw new Error(data.message || `HTTP ${response.status}`)
        }
        ElMessage.success(`已导入 ${data.count || docxFiles.length} 个模板`)
        void loadMyTemplates()
    } catch (error) {
        ElMessage.error(String(error?.message || error || '文件夹导入失败'))
    } finally {
        templateFolderUploading.value = false
    }
}

async function handleFolderUpload(event) {
    const files = event.target.files
    await uploadTemplateFolderFiles(files)
    event.target.value = ''
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
    fillStreamOutput.value = ''
    if (wordContainer.value) {
        wordContainer.value.innerHTML = ''
    }
}

async function handleExportWord() {
    if (!templateArrayBuffer.value) {
        showToast('error', '请先上传模板')
        return
    }

    try {
        await syncToWord({ render: false })
    } catch (error) {
        console.error('[docx export] failed', error)
        showToast('error', '导出失败，请稍后重试')
        return
    }
    if (!lastWordBlob.value) {
        showToast('error', '导出失败，请稍后重试')
        return
    }

    const originalName = String(currentFile.value?.name || 'huiyi-template.docx')
    const safeBase = originalName.replace(/\.docx$/i, '') || 'huiyi-template'
    saveAs(lastWordBlob.value, `${safeBase}-filled.docx`)
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
        <!-- 沉浸 Tab 的待办入口由各自头部（对话工具栏 / 陪伴页头部）内的 TodoBell 提供，
             此处不再额外放一个 fixed 按钮，否则会在内容顶部占出一截空白 -->
            <!-- Topbar（沉浸 Tab 隐藏，把高度还给聊天区） -->
            <div v-show="!isImmersiveWorkbenchTab" class="topbar">
                <div>
                    <h1 class="page-title">{{ pageTitle }}</h1>
                    <span class="breadcrumb">首页 / 工作台</span>
                </div>
                <div class="topbar-actions topbar-actions--end">
                    <!-- Teleport 挂载点：API 配置页的「新增配置」会传送至此，顺序为 [新增配置][🔔][头像] -->
                    <div id="header-page-actions" class="header-page-actions"></div>
                    <button
                        v-if="userStore.user?.email"
                        type="button"
                        class="topbar-todo-bell"
                        title="待办事项"
                        aria-label="待办事项"
                        @click="openTodoDrawer"
                    >
                        <span class="topbar-todo-bell-icon" aria-hidden="true">🔔</span>
                        <span v-if="todoPendingCount > 0" class="topbar-todo-badge">{{ todoPendingCount > 99 ? '99+' : todoPendingCount }}</span>
                    </button>
                    <el-dropdown v-if="userStore.user?.email" trigger="click" placement="bottom-end">
                        <button type="button" class="topbar-user-chip" title="个人中心" aria-label="个人中心">
                            <span class="topbar-user-avatar">{{ (userStore.user?.name || 'U').charAt(0).toUpperCase() }}</span>
                            <span class="topbar-user-name">{{ userStore.user?.name || '用户' }}</span>
                        </button>
                        <template #dropdown>
                            <el-dropdown-menu>
                                <el-dropdown-item @click="goTab('identity')">多身份切换</el-dropdown-item>
                                <el-dropdown-item @click="goTab('apiConfig')">API 与模型</el-dropdown-item>
                                <el-dropdown-item divided @click="handleLogout">退出登录</el-dropdown-item>
                            </el-dropdown-menu>
                        </template>
                    </el-dropdown>
                </div>
            </div>

            <Teleport v-if="currentTab === 'apiConfig' && userStore.user?.email" to="#header-page-actions">
                <el-button type="primary" round :icon="Plus" class="header-api-add-btn" @click="openModelModal(null)">新增配置</el-button>
            </Teleport>

            <div v-if="dueTodoAlertVisible" class="dashboard-reminder-bar" role="alert">
                <div
                    class="dashboard-reminder-main"
                    tabindex="0"
                    role="button"
                    @click="openTodoDrawer"
                    @keydown.enter.prevent="openTodoDrawer"
                >
                    <span class="dashboard-reminder-label dashboard-reminder-label--overdue">到点提醒</span>
                    <span class="dashboard-reminder-count">您有 {{ dueTodoAlerts.length }} 个待办已到期</span>
                    <div class="dashboard-reminder-marquee">
                        <span
                            v-for="t in dueTodoAlerts.slice(0, 6)"
                            :key="`due-${t.id}`"
                            class="dashboard-reminder-chip dashboard-reminder-chip--overdue"
                            :title="t.title"
                        >{{ t.title }}</span>
                        <span v-if="dueTodoAlerts.length > 6" class="dashboard-reminder-more">还有 {{ dueTodoAlerts.length - 6 }} 项…</span>
                    </div>
                    <span class="dashboard-reminder-hint">点击查看全部</span>
                </div>
                <div class="dashboard-reminder-actions">
                    <button type="button" class="btn btn-outline btn-xs" @click="openTodoDrawer">查看待办</button>
                    <button type="button" class="btn btn-ghost btn-xs" @click="dismissDueTodoAlerts">知道了</button>
                </div>
            </div>

            <el-drawer v-model="todoDrawerVisible" title="待办事项" direction="rtl" size="420px" append-to-body class="todo-drawer">
                <p v-if="!userStore.user?.email" class="todo-drawer-guest">登录后查看从对话与会议纪要中静默提取的待办。</p>
                <template v-else>
                    <div v-if="todosDrawerLoading" class="todo-drawer-loading">加载中…</div>
                    <ul v-else-if="sortedTodosDrawerList.length === 0" class="todo-drawer-empty">暂无待办</ul>
                    <ul v-else class="todo-drawer-list">
                        <li
                            v-for="t in sortedTodosDrawerList"
                            :key="t.id"
                            class="todo-drawer-item todo-alarm-card todo-drawer-item--card"
                            :class="{ 'todo-drawer-item--done': Number(t.done) === 1 }"
                        >
                            <div class="todo-alarm-card__time">{{ formatTodoClock(t.remindAt) }}</div>
                            <div class="todo-drawer-item-main todo-alarm-card__body">
                                <div class="todo-alarm-card__head">
                                    <span class="todo-drawer-title">{{ t.title }}</span>
                                    <span
                                        class="todo-drawer-status-tag"
                                        :class="{
                                            'todo-drawer-status-tag--pending': todoStatusDisplay(t).kind === 'pending',
                                            'todo-drawer-status-tag--notified': todoStatusDisplay(t).kind === 'notified',
                                            'todo-drawer-status-tag--done': todoStatusDisplay(t).kind === 'done',
                                            'todo-drawer-status-tag--repeat': todoStatusDisplay(t).kind === 'repeat',
                                            'todo-drawer-status-tag--repeat_cleared': todoStatusDisplay(t).kind === 'repeat_cleared',
                                        }"
                                    >{{ todoStatusDisplay(t).label }}</span>
                                    <span
                                        class="todo-mode-tag"
                                        :class="{
                                            'todo-mode-tag--daily': todoRepeatTag(t.repeatMode, t.repeatCustomDays).kind === 'daily',
                                            'todo-mode-tag--workday': todoRepeatTag(t.repeatMode, t.repeatCustomDays).kind === 'workday',
                                            'todo-mode-tag--custom': todoRepeatTag(t.repeatMode, t.repeatCustomDays).kind === 'custom',
                                        }"
                                    >{{ todoRepeatTag(t.repeatMode, t.repeatCustomDays).label }}</span>
                                </div>
                                <span class="todo-drawer-meta">{{ formatLocalDateTimeWithRelative(t.remindAt) }}{{ t.source ? ' · ' + t.source : '' }}</span>
                                <p v-if="t.notes" class="todo-drawer-notes">{{ t.notes }}</p>
                            </div>
                            <div class="todo-drawer-actions">
                                <button
                                    v-if="Number(t.done) === 0"
                                    type="button"
                                    class="btn btn-accent btn-xs todo-snooze-btn"
                                    @click="openReminderTimeEdit(t)"
                                >
                                    <span class="todo-snooze-inner">
                                        <el-icon :size="14"><AlarmClock /></el-icon>
                                        ⏰ 调整时间
                                    </span>
                                </button>
                                <button v-if="Number(t.done) === 0" type="button" class="btn btn-outline btn-xs" @click="markTodoDoneRow(t)">完成</button>
                                <button
                                    type="button"
                                    class="btn btn-ghost btn-xs todo-card-delete-btn"
                                    title="删除待办"
                                    aria-label="删除待办"
                                    @click="confirmDeleteTodoRow(t)"
                                >
                                    <el-icon :size="16"><Delete /></el-icon>
                                </button>
                            </div>
                        </li>
                    </ul>
                </template>
            </el-drawer>

            <!-- ===== TAB: OVERVIEW ===== -->
            <div v-show="currentTab === 'overview'" class="tab-content" id="tabOverview">
                <div class="stats-row">
                    <div class="stat-card stat-card--metric">
                        <div class="stat-metric-left" style="background: #eef2ff; color: #4a6cf7">
                            <el-icon class="stat-metric-icon" :size="30"><ChatLineRound /></el-icon>
                        </div>
                        <div class="stat-metric-body">
                            <div class="stat-metric-label">累计对话轮次</div>
                            <div class="stat-metric-value">{{ dashboardStats.conversation_turns }}</div>
                            <div class="stat-metric-foot">系统持续为您沉淀对话价值</div>
                        </div>
                    </div>
                    <div class="stat-card stat-card--metric">
                        <div class="stat-metric-left" style="background: #ecfdf5; color: #059669">
                            <el-icon class="stat-metric-icon" :size="30"><MagicStick /></el-icon>
                        </div>
                        <div class="stat-metric-body">
                            <div class="stat-metric-label">已生成技能数</div>
                            <div class="stat-metric-value">{{ dashboardStats.skills_count }}</div>
                            <div class="stat-metric-foot">技能与场景能力持续扩展</div>
                        </div>
                    </div>
                    <div class="stat-card stat-card--metric">
                        <div class="stat-metric-left" style="background: #fffbeb; color: #d97706">
                            <el-icon class="stat-metric-icon" :size="30"><Reading /></el-icon>
                        </div>
                        <div class="stat-metric-body">
                            <div class="stat-metric-label">长文处理次数</div>
                            <div class="stat-metric-value">{{ dashboardStats.papers_processed }}</div>
                            <div class="stat-metric-foot">长文速读与卡片化摘要累计次数</div>
                        </div>
                    </div>
                    <div class="stat-card stat-card--ratings-trend">
                        <div class="stat-ratings-head">
                            <div class="stat-metric-left stat-ratings-head-icon" style="background: #f5f3ff; color: #7c3aed">
                                <el-icon :size="26"><Histogram /></el-icon>
                            </div>
                            <div class="stat-ratings-head-text">
                                <div class="stat-label">认可即沉淀</div>
                                <div class="stat-subhint">历史满意度 / 点赞趋势（近 90 天）</div>
                            </div>
                        </div>
                        <div class="stat-ratings-chart-area">
                            <div v-if="ratingsTrendLoading" class="stat-ratings-empty">加载趋势中…</div>
                            <div v-else-if="!ratingsTrendData.length" class="stat-ratings-empty">暂无评分数据，可在对话中为回复点赞以形成趋势</div>
                            <div v-else ref="ratingsTrendChartRef" class="stat-ratings-chart-el"></div>
                        </div>
                    </div>
                </div>

                <div class="template-showcase">
                    <h2>智能模板引擎</h2>
                    <p class="subtitle">
                        在<strong>个人资料库 · 我的模板库</strong>中上传 <code>.docx</code> 模板；在本页使用
                        <strong>&#123;&#123;自动填充&#125;&#125;</strong> 与 <strong>{变量}</strong> 创建与编辑模板、一键调用填充并导出 Word。
                    </p>
                    <div class="template-code-block">
                        <span style="color:#6c7086;"># 示例：周报骨架</span><br>
                        <span style="color:#a6e3a1;">## 本周重点</span><br>
                        <span class="token-fill">&#123;&#123;自动填充本周关键进展&#125;&#125;</span><br>
                        <span style="color:#a6e3a1;">## 风险与依赖</span><br>
                        <span class="token-var">{风险项}</span> <span style="color:#6c7086;">（自行命名变量，由模型按上下文补全）</span><br>
                        <span style="color:#a6e3a1;">## 下周计划</span><br>
                        <span class="token-var">{下周计划要点}</span>
                    </div>
                    <div class="template-demo-row">
                        <div v-if="false" class="template-demo-card" @click="goTab('templates')">
                            <div class="demo-icon">&#9679;</div>
                            <div class="label">会议纪要模板</div>
                            <div class="hint">12个变量 · 自动填充</div>
                        </div>
                        <div class="template-demo-card" @click="goTab('library')">
                            <div class="demo-icon">&#9634;</div>
                            <div class="label">我的模板库</div>
                            <div class="hint">上传与管理 .docx</div>
                        </div>
                        <div class="template-demo-card" @click="goTab('templates')">
                            <div class="demo-icon demo-icon--ep" aria-hidden="true"><el-icon :size="22"><Document /></el-icon></div>
                            <div class="label">创建 / 编辑 / 调用</div>
                            <div class="hint">填充变量 · 导出 Word</div>
                        </div>
                        <div v-if="false" class="template-demo-card" @click="goTab('templates')">
                            <div class="demo-icon">&#9650;</div>
                            <div class="label">PPT大纲模板</div>
                            <div class="hint">逐页自动生成</div>
                        </div>
                    </div>
                </div>

                <div class="longread-showcase card-like">
                    <div class="longread-showcase__icon" aria-hidden="true">
                        <el-icon :size="28"><Reading /></el-icon>
                    </div>
                    <div class="longread-showcase__body">
                        <h2 class="longread-showcase__title">智能长文速读</h2>
                        <p class="longread-showcase__lead">
                            上传 PDF / Word 或拉取网页正文，一键生成<strong>类型识别 + 卡片化要点</strong>，节省通读时间；并可嗅探文中潜在待办，批量汇入「主动提醒」。
                        </p>
                        <ul class="longread-showcase__bullets">
                            <li>自适应维度：论文、新闻、手册、纪要等不同体裁自动换用不同卡片结构</li>
                            <li>可带入对话助手，基于摘要与原文节选深度追问</li>
                            <li>发现行动项时支持「批量加入我的待办」</li>
                        </ul>
                        <el-button type="primary" plain size="small" class="longread-showcase__cta" @click="goTab('paper')">进入智能长文速读</el-button>
                    </div>
                </div>

                <div class="section-header">
                    <h3>全部功能模块</h3>
                    <span class="see-all">拖拽排序</span>
                </div>
                <div class="cards-grid">
                    <div v-if="false" class="card" @click="goTab('meeting')">
                        <div class="card-icon blue">&#9679;</div>
                        <span class="card-badge badge-popular">热门</span>
                        <h3>智能会议纪要</h3>
                        <p>实时 WebSocket 转写→结构化纪要+日历待办自动推送</p>
                        <span class="card-arrow">&#8594;</span>
                    </div>
                    <div class="card" @click="goTab('paper')">
                        <div class="card-icon green">&#9633;</div>
                        <span class="card-badge badge-pro">PRO</span>
                        <h3>智能长文速读</h3>
                        <p>PDF / Word / 网页正文 → 卡片化核心要点；嗅探潜在待办并一键加入提醒引擎</p>
                        <span class="card-arrow">&#8594;</span>
                    </div>
                    <div v-if="false" class="card" @click="goTab('ppt')">
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
                    暂无集中的操作流水；上方数字来自服务端统计。在各功能 Tab 中完成的文档、长文速读与对话会反映在数据中。
                </p>
            </div>

            <!-- ===== TAB: CHAT（Gemini 式双栏 + 悬浮输入） ===== -->
            <div v-show="currentTab === 'chat'" class="tab-content tab-chat-workbench tab-chat-gemini-root" id="tabChat">
                <div class="chat-gemini-layout" :class="{ 'chat-gemini-layout--history-open': chatHistoryOverlayOpen }">
                    <div class="chat-gemini-backdrop" @click="chatHistoryOverlayOpen = false"></div>
                    <aside
                        v-if="chatLayoutIsNarrow || !chatHistoryCollapsed"
                        class="chat-history-sidebar"
                        aria-label="会话历史"
                        :style="chatHistorySidebarStyle"
                    >
                        <div class="chat-history-sidebar-toolbar">
                            <button
                                v-if="!chatLayoutIsNarrow"
                                type="button"
                                class="chat-history-collapse-btn"
                                title="收起会话列表"
                                aria-label="收起会话列表"
                                @click="toggleChatHistoryCollapse"
                            >
                                <el-icon :size="16"><DArrowLeft /></el-icon>
                            </button>
                            <button type="button" class="chat-new-session-btn chat-new-session-btn--toolbar" @click="createNewChatSession">
                                + 新对话
                            </button>
                        </div>
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
                                <span v-if="s.personaLabel || s.skillName" class="chat-history-meta">
                                    <span class="chat-history-meta-line">
                                        <el-icon class="chat-history-meta-ico" :size="14"><component :is="chatSessionPersonaIcon(s)" /></el-icon>
                                        <span>{{ s.personaLabel || '默认身份' }}</span>
                                    </span>
                                    <span v-if="s.skillName" class="chat-history-meta-line chat-history-meta-line--skill">
                                        <el-icon class="chat-history-meta-ico" :size="14"><Lightning /></el-icon>
                                        <span>{{ s.skillName }}</span>
                                    </span>
                                </span>
                                <button
                                    type="button"
                                    class="chat-history-delete"
                                    title="删除会话"
                                    aria-label="删除会话"
                                    @click.stop="deleteChatSessionRow(s, $event)"
                                >
                                    <el-icon :size="16"><Delete /></el-icon>
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
                    <div
                        v-if="!chatLayoutIsNarrow && !chatHistoryCollapsed"
                        class="chat-history-resizer"
                        aria-hidden="true"
                        @mousedown.prevent="startChatHistoryResize"
                    />
                    <button
                        v-if="!chatLayoutIsNarrow && chatHistoryCollapsed"
                        type="button"
                        class="chat-history-expand-rail"
                        title="展开会话列表"
                        aria-label="展开会话列表"
                        @click="toggleChatHistoryCollapse"
                    >
                        <el-icon :size="18"><DArrowRight /></el-icon>
                    </button>
                    <div class="chat-gemini-main">
                        <header class="chat-gemini-toolbar">
                            <button
                                type="button"
                                class="chat-history-menu-btn"
                                :aria-label="chatLayoutIsNarrow ? '展开或收起会话列表' : (chatHistoryCollapsed ? '展开会话列表' : '收起会话列表')"
                                @click="
                                    chatLayoutIsNarrow ? (chatHistoryOverlayOpen = !chatHistoryOverlayOpen) : toggleChatHistoryCollapse()
                                "
                            >
                                <el-icon :size="20"><Menu /></el-icon>
                            </button>
                            <div class="chat-gemini-toolbar-center">
                                <span class="chat-gemini-brand">会意助手</span>
                                <span v-if="chatLoadedPaperLabel" class="chat-gemini-pill">{{ chatLoadedPaperLabel }}</span>
                                <span v-if="activeChatSkill" class="chat-gemini-pill chat-gemini-pill--skill">
                                    <el-icon class="chat-gemini-pill-ico" :size="13"><Lightning /></el-icon>
                                    {{ activeChatSkill.name }}
                                    <button type="button" class="chat-gemini-pill-x" title="取消 Skill" @click="clearActiveChatSkill">×</button>
                                </span>
                            </div>
                            <div class="chat-gemini-toolbar-actions">
                                <TodoBell
                                    v-if="userStore.user?.email"
                                    :count="todoPendingCount"
                                    @open="openTodoDrawer"
                                />
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
                                    <div class="msg-avatar">{{ message.role === 'assistant' ? '会' : (userStore.user?.name || 'U').charAt(0).toUpperCase() }}</div>
                                    <div v-if="message.isHtml" class="msg-bubble" v-html="sanitizeHtml(message.content)"></div>
                                    <div v-else-if="message.role === 'assistant'" class="msg-bubble assistant-stream-bubble">
                                        <div v-if="message.statusBanner" class="chat-agent-status" role="status">{{ message.statusBanner }}</div>
                                        <div
                                            v-if="message.clarificationCard"
                                            class="clarification-card"
                                            :class="message.clarificationCard.resolved ? 'clarification-card--resolved' : 'clarification-card--active'"
                                        >
                                            <div class="clarification-card-title">还需要一点细节</div>
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
                                                :class="{ 'chat-like-btn--on': isAssistantLiked(message) }"
                                                title="为本条回复评分（点赞）"
                                                :disabled="!userStore.user?.email"
                                                @click="onAssistantThumbsUp(message)"
                                            >
                                                <el-icon class="chat-like-ico" :size="16"><Pointer /></el-icon>
                                                <span class="chat-like-label">赞</span>
                                            </button>
                                            <button
                                                type="button"
                                                class="chat-pack-skill-btn"
                                                title="将当前对话提炼并打包为 Skill"
                                                :disabled="skillExtractLoading || !userStore.user?.email || !canPackSkill"
                                                @click="extractSkillFromConversation"
                                            >
                                                <el-icon class="chat-like-ico" :size="16"><Lightning /></el-icon>
                                                <span class="chat-like-label">打包 Skill</span>
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
                                    placeholder="问问会意助手…（Shift+Enter 换行，Enter 发送；输入 @ 唤醒技能）"
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
                                        <span class="chat-mic-icon" aria-hidden="true"><el-icon :size="18"><Microphone /></el-icon></span>
                                    </button>
                                    <button
                                        v-if="chatBusy"
                                        type="button"
                                        class="send-btn send-btn--inset send-btn--stop"
                                        title="停止生成"
                                        aria-label="停止生成"
                                        @click="stopChatStream"
                                    >
                                        &#9632;
                                    </button>
                                    <button v-else type="button" class="send-btn send-btn--inset" title="发送" aria-label="发送" @click="sendChatMessage()">&#9654;</button>
                                </div>
                            </div>
                            </div>
                        </div>
                    </div>
                </div>
                <el-drawer
                    v-model="impressionDrawerVisible"
                    title="会意助手眼中的你"
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
                            <span class="btn-ico-text">
                                <el-icon :size="16"><TrendCharts /></el-icon>
                                查看我的情绪仪表盘
                            </span>
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
                            <input
                                id="templates-folder-upload"
                                type="file"
                                accept=".docx"
                                multiple
                                webkitdirectory
                                directory
                                class="word-upload-input word-upload-input--folder"
                                aria-label="上传模板文件夹"
                                @change="handleFolderUpload"
                            />
                            <div class="word-upload-hint">
                                <span class="word-upload-icon">&#8679;</span>
                                <div>
                                    <div style="font-weight:600;">点击或拖拽上传 .docx 文件</div>
                                    <div style="font-size:0.75rem;color:var(--text-muted);">单文件最大 20MB；也可导入模板文件夹，自动入库</div>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="template-folder-actions">
                        <label class="btn btn-outline btn-sm" for="templates-folder-upload">
                            {{ templateFolderUploading ? '导入中…' : '上传文件夹作为模板库' }}
                        </label>
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
                                <button class="btn-export" :disabled="!templateArrayBuffer" @click="handleExportWord">导出 Word</button>
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
                            <div v-if="isFilling || fillStreamOutput" class="panel-stream">
                                <div class="panel-stream-title">流式输出</div>
                                <pre class="panel-stream-body">{{ fillStreamOutput || '等待模型输出…' }}</pre>
                            </div>
                            <div v-if="fillError" class="panel-error">{{ fillError }}</div>
                        </div>
                    </div>
                    <div v-if="!currentFile" class="template-syntax-guide">
                        <h4 class="template-syntax-title">占位符怎么用（真实生效规则）</h4>
                        <ul class="template-syntax-list">
                            <li v-pre>
                                在 Word 正文里优先用<strong>半角双花括号</strong>写变量名，例如 <code>{{报告主题}}</code>、<code>{{汇报人}}</code>。系统会自动识别并在渲染前转换为 Docxtemplater 可填充槽位；历史单花括号 <code>{报告主题}</code> 仍兼容。
                            </li>
                            <li>
                                上传模板后，右侧面板会列出识别到的变量。点击「AI 自动填充」时，模型按上下文填入；也可手动改每条变量再导出。
                            </li>
                            <li>
                                需要 AI 扩写一段说明时，在花括号内写清楚指令，例如 <code>{请用三句话写本周风险与对策}</code>，描述越具体，填充越稳。
                            </li>
                        </ul>
                    </div>
                </div>
            </div>


            <!-- ===== TAB: LIBRARY ===== -->
            <div v-show="currentTab === 'library'" class="tab-content tab-library" id="tabLibrary">
                <div class="library-hero">
                    <div>
                        <p class="library-kicker">Personal Knowledge Base</p>
                        <h2>个人资料库</h2>
                        <p>把「模板库」和「文件资料库」放在一个入口：上次写过的模板结果、上传过的 PDF / Docx、会议纪要都会在这里沉淀，方便复用和继续写作。</p>
                    </div>
                    <div class="library-hero-actions">
                        <button type="button" class="btn btn-outline" @click="goTab('templates')">去模板引擎写文档</button>
                        <button type="button" class="btn btn-primary" :disabled="libraryLoading" @click="agentOrganizeLibrary">Agent 自动整理</button>
                    </div>
                </div>

                <div class="library-tabs" role="tablist" aria-label="资料库类型">
                    <button type="button" :class="{ active: libraryActivePane === 'templates' }" @click="libraryActivePane = 'templates'">模板库</button>
                    <button type="button" :class="{ active: libraryActivePane === 'files' }" @click="libraryActivePane = 'files'">文件资料库</button>
                </div>

                <section v-show="libraryActivePane === 'templates'" class="library-section">
                    <div class="library-section-head">
                        <div>
                            <h3>模板库：复用模板与上次填写结果</h3>
                            <p>这里展示你上传过的 Word 模板，以及 AI 自动填充后保存的历史记录。写类似材料时，可以先下载上次结果、复用变量，或打开原模板继续生成。</p>
                        </div>
                        <button type="button" class="btn btn-outline btn-sm" @click="loadMyTemplates(); loadTemplateFillRecords(30)">刷新</button>
                    </div>
                    <div class="library-grid-two">
                        <div class="library-card-panel">
                            <div class="library-card-panel-head">
                                <h4>已上传模板</h4>
                                <span>{{ myTemplates.length }} 个</span>
                            </div>
                            <div v-if="!userStore.user?.email" class="library-empty">登录后可查看你的模板库。</div>
                            <div v-else-if="myTemplatesLoading" class="library-empty">模板加载中…</div>
                            <div v-else-if="myTemplates.length === 0" class="library-empty">暂无模板。请到「智能模板引擎」上传 .docx。</div>
                            <ul v-else class="library-list">
                                <li v-for="row in myTemplates" :key="row.id" class="library-row">
                                    <div class="library-row-main">
                                        <strong>{{ row.originalName }}</strong>
                                        <span><template v-if="row.folderPath"><span class="lib-folder-inline"><el-icon :size="13"><Folder /></el-icon> {{ row.folderPath }} · </span></template>{{ row.uploadedAt || row.uploaded_at }}</span>
                                        <em v-if="row.tags?.length">{{ row.tags.join(' / ') }}</em>
                                    </div>
                                    <div class="library-row-actions">
                                        <button type="button" class="btn btn-outline btn-xs" @click="openMyTemplate(row); goTab('templates')">打开</button>
                                        <button type="button" class="btn btn-outline btn-xs" @click="editTemplateOrganization(row)">整理</button>
                                        <button type="button" class="btn btn-ghost btn-xs" @click="deleteMyTemplate(row)">删除</button>
                                    </div>
                                </li>
                            </ul>
                        </div>
                        <div class="library-card-panel">
                            <div class="library-card-panel-head">
                                <h4>填充历史</h4>
                                <span>{{ templateFillRecords.length }} 条</span>
                            </div>
                            <div v-if="!userStore.user?.email" class="library-empty">登录后可回溯生成历史。</div>
                            <div v-else-if="templateFillRecords.length === 0" class="library-empty">暂无历史。模板 AI 填充完成后会自动存档。</div>
                            <ul v-else class="library-list">
                                <li v-for="row in templateFillRecords" :key="row.id" class="library-row">
                                    <div class="library-row-main">
                                        <strong>{{ row.templateName || '未命名模板' }}</strong>
                                        <span>{{ formatLocalDateTime(row.createdAt) }}</span>
                                    </div>
                                    <div class="library-row-actions">
                                        <button type="button" class="btn btn-outline btn-xs" @click="applyTemplateRecordVariables(row)">复用变量</button>
                                        <button type="button" class="btn btn-outline btn-xs" @click="downloadTemplateRecord(row)">下载</button>
                                    </div>
                                </li>
                            </ul>
                        </div>
                    </div>
                </section>

                <section v-show="libraryActivePane === 'files'" class="library-section">
                    <div class="library-section-head">
                        <div>
                            <h3>文件资料库：文档、论文与会议纪要</h3>
                            <p>支持手动上传 PDF / Word（.docx，旧版 .doc 仅存文件）/ TXT / MD；文献解析、结构化摘要和会议纪要会自动归档到这里。你可手动建文件夹，也可用 Agent 整理。</p>
                        </div>
                        <div class="library-section-actions">
                            <input ref="libraryFileInputRef" class="sr-only" type="file" accept=".pdf,.doc,.docx,.txt,.md,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document" @change="handleLibraryFileUpload" />
                            <button type="button" class="btn btn-outline btn-sm" :disabled="libraryUploading" @click="triggerLibraryUpload">上传资料</button>
                            <button type="button" class="btn btn-primary btn-sm" :disabled="libraryLoading" @click="agentOrganizeLibrary">Agent 整理</button>
                        </div>
                    </div>
                    <div class="library-toolbar">
                        <label class="sr-only" for="librarySearchInput">搜索资料库</label>
                        <input id="librarySearchInput" v-model="librarySearch" type="search" placeholder="搜索标题、文件夹或标签…" />
                        <button type="button" class="btn btn-outline btn-sm" @click="loadLibraryItems">刷新</button>
                    </div>
                    <div v-if="libraryFolders.length" class="library-folder-cloud">
                        <span v-for="f in libraryFolders" :key="f.name" class="lib-folder-chip"><el-icon :size="13"><Folder /></el-icon> {{ f.name }} <b>{{ f.count }}</b></span>
                    </div>
                    <div v-if="!userStore.user?.email" class="library-empty library-empty--wide">登录后可查看个人资料库。</div>
                    <div v-else-if="libraryLoading" class="library-empty library-empty--wide">资料库加载中…</div>
                    <div v-else-if="libraryFilteredItems.length === 0" class="library-empty library-empty--wide">暂无资料。上传文件、处理会议纪要或解析文献后会自动出现在这里。</div>
                    <ul v-else class="library-file-list">
                        <li v-for="row in libraryFilteredItems" :key="row.id" class="library-file-row">
                            <div class="library-file-icon library-file-icon--ep" aria-hidden="true">
                                <el-icon :size="20"><component :is="libraryItemTypeIconComponent(row.itemType)" /></el-icon>
                            </div>
                            <div class="library-file-main">
                                <div class="library-file-title">
                                    <template v-if="libraryRenamingId === row.id">
                                        <el-input
                                            v-model="libraryRenameDraft"
                                            size="small"
                                            class="library-rename-input"
                                            maxlength="240"
                                            @blur="commitLibraryRename(row)"
                                            @keyup.enter="commitLibraryRename(row)"
                                            @keydown.esc.prevent="cancelLibraryRename"
                                        />
                                    </template>
                                    <template v-else>
                                        <strong>{{ row.title }}</strong>
                                        <button type="button" class="library-rename-ico" title="重命名" aria-label="重命名" @click.stop="startLibraryRename(row)">
                                            <el-icon :size="16"><EditPen /></el-icon>
                                        </button>
                                        <span>{{ libraryItemTypeLabel(row.itemType) }}</span>
                                    </template>
                                </div>
                                <p v-if="row.summaryText" class="library-file-excerpt">{{ String(row.summaryText).slice(0, 160) }}</p>
                                <p v-else-if="row.contentText" class="library-file-excerpt">{{ String(row.contentText).slice(0, 160) }}</p>
                                <div class="library-file-meta">
                                    <span class="lib-folder-inline"><el-icon :size="13"><Folder /></el-icon> {{ row.folderPath || '未分组' }}</span>
                                    <span>{{ formatLocalDateTime(row.createdAt) }}</span>
                                    <span v-if="row.tags?.length">{{ row.tags.join(' / ') }}</span>
                                </div>
                            </div>
                            <div class="library-row-actions">
                                <button v-if="row.canDownload" type="button" class="btn btn-outline btn-xs" @click="downloadLibraryItem(row)">下载</button>
                                <button type="button" class="btn btn-outline btn-xs" @click="editLibraryItem(row)">整理</button>
                                <button type="button" class="btn btn-ghost btn-xs" @click="deleteLibraryItem(row)">删除</button>
                            </div>
                        </li>
                    </ul>
                </section>
            </div>

            <!-- ===== TAB: MEETING ===== -->
            <div v-show="currentTab === 'meeting'" class="tab-content tab-meeting-minutes" id="tabMeeting">
                <MeetingMinutesView @processed="onMeetingProcessed" />
            </div>

    <!-- ===== TAB: PAPER（全屏学术工作站 · 双栏） ===== -->
    <div v-show="currentTab === 'paper'" class="tab-content tab-paper" id="tabPaper">
        <div v-if="!paperWorkbenchHasSource" class="paper-tab-empty">
            <el-empty description="请先上传 PDF / Word 或拉取链接正文，再生成智能长文速读卡片" :image-size="88">
                <template #default>
                    <p class="paper-empty-hint">支持 PDF / Word（推荐 .docx；旧版 .doc 请先另存为 .docx），或通过链接拉取正文后再生成摘要。</p>
                    <div class="paper-empty-upload">
                        <label class="sr-only" for="tab-paper-upload-file-empty">上传 PDF 或 Word 文献</label>
                        <input
                            id="tab-paper-upload-file-empty"
                            ref="paperFileInputRef"
                            name="paper_upload_file_empty"
                            type="file"
                            accept=".pdf,.doc,.docx,application/pdf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
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
                                    accept=".pdf,.doc,.docx,application/pdf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
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
                                    <summary class="paper-details-summary">标题、摘录与合并（可选展开）</summary>
                                    <div class="paper-details-body">
                                        <label class="paper-mini-label" for="tab-paper-doc-title">文档标题（可选）</label>
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
                                <div v-if="paperPotentialTodosList.length" class="paper-potential-todos" role="status">
                                    <p class="paper-potential-todos__tip">
                                        💡 发现 {{ paperPotentialTodosList.length }} 条潜在待办，请勾选后加入
                                    </p>
                                    <ul class="paper-potential-todos__list">
                                        <li
                                            v-for="(pt, pidx) in paperPotentialTodosList"
                                            :key="`pt-${pidx}-${pt.title.slice(0, 24)}`"
                                            class="paper-potential-todos__row"
                                        >
                                            <el-checkbox
                                                :model-value="!!paperPotentialTodoPick[pidx]"
                                                :disabled="!userStore.user?.email || paperPotentialSubmitting"
                                                @change="(v) => setPaperPotentialPick(pidx, v)"
                                            >
                                                <span class="paper-potential-todos__title">{{ pt.title }}</span>
                                                <span v-if="pt.remind_time" class="paper-potential-todos__time">{{ pt.remind_time }}</span>
                                            </el-checkbox>
                                        </li>
                                    </ul>
                                    <el-button
                                        type="primary"
                                        size="small"
                                        :loading="paperPotentialSubmitting"
                                        :disabled="!userStore.user?.email || paperPotentialSubmitting"
                                        @click="addSelectedPaperPotentialTodos"
                                    >
                                        加入我的待办
                                    </el-button>
                                </div>
                                <div class="paper-chat-actions">
                                    <el-button type="primary" class="paper-deep-btn paper-deep-btn--half" @click="openPaperContextChat">
                                        带入上下文对话
                                    </el-button>
                                    <el-button type="primary" plain class="paper-deep-btn paper-deep-btn--half paper-deep-btn--plain" @click="openPaperDeepDialog">
                                        要点 + 原文讨论
                                    </el-button>
                                </div>
                                <div class="paper-openclaw-command" role="navigation" aria-label="指挥会意助手处理此摘要">
                                    <div class="paper-openclaw-command__title">指挥会意助手处理此内容</div>
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
                            <p v-else-if="paperAiReading" class="paper-phase-msg">AI 正在通读材料并提取要点…</p>
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
        <div class="chat-header companion-chat-header-row">
            <div class="chat-avatar companion-avatar">&#9825;</div>
            <div class="companion-chat-header-text">
                <div class="chat-title">情感陪伴 · 暖心扣字</div>
                <div class="chat-status companion-status">在这儿陪你</div>
            </div>
            <div class="companion-header-actions">
                <TodoBell v-if="userStore.user?.email" :count="todoPendingCount" @open="openTodoDrawer" />
                <button type="button" class="btn btn-ghost btn-xs" @click="clearCompanionChat" title="清空并重置开场">清空</button>
            </div>
        </div>
        <div class="companion-persona-panel">
            <div class="companion-persona-row">
                <label class="companion-persona-label" for="companion-persona-select">陪伴人格</label>
                <select id="companion-persona-select" v-model="companionSettings.persona" class="companion-persona-select" name="companion_persona">
                    <option value="default">默认会意助手</option>
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
                    @keydown="onCompanionComposerKeydown"
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
                        <span class="chat-mic-icon" aria-hidden="true"><el-icon :size="18"><Microphone /></el-icon></span>
                    </button>
                    <button
                        v-if="chatBusy"
                        type="button"
                        class="send-btn companion-send send-btn--stop"
                        title="停止生成"
                        @click="stopChatStream"
                    >
                        &#9632;
                    </button>
                    <button v-else type="button" class="send-btn companion-send" @click="sendChatMessage(undefined, { chatMode: CHAT_MODE_COMPANION })" title="发送">&#9654;</button>
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
            <div class="identity-persona-icon-wrap" aria-hidden="true">
                <el-icon :size="28" class="identity-persona-elicon"><component :is="p.iconComponent" /></el-icon>
            </div>
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
            <div class="identity-persona-icon-wrap" aria-hidden="true">
                <el-icon :size="28" class="identity-persona-elicon"><UserFilled /></el-icon>
            </div>
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
    <label class="welcome-wizard-label" for="cp-emoji">角标（1–2 字符，可选）</label>
    <input
        id="cp-emoji"
        v-model="customPersonaForm.emoji"
        class="welcome-wizard-input"
        maxlength="2"
        name="custom_persona_emoji"
        placeholder="如 研"
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
        <div class="card" style="grid-column:1/-1;max-width:760px;">
            <div class="card-icon amber">&#9200;</div>
            <h3>主动提醒与推送</h3>
            <p>从会议纪要与对话自动提取待办，并在到点时推送（微信优先）。</p>
            <div class="reminder-form">
                <div class="reminder-form-row">
                    <label class="sr-only" for="reminder-title">待办标题</label>
                    <input
                        id="reminder-title"
                        v-model="reminderForm.title"
                        type="text"
                        class="reminder-input"
                        placeholder="待办标题，例如：提交周报"
                    />
                </div>
                <div class="reminder-form-row">
                    <label class="sr-only" for="reminder-time">提醒时间</label>
                    <input
                        id="reminder-time"
                        v-model="reminderForm.remindAt"
                        type="datetime-local"
                        class="reminder-input"
                    />
                </div>
                <div class="reminder-form-row">
                    <label class="sr-only" for="reminder-notes">备注</label>
                    <textarea
                        id="reminder-notes"
                        v-model="reminderForm.notes"
                        class="reminder-input reminder-textarea"
                        rows="2"
                        placeholder="备注（可选）"
                    ></textarea>
                </div>
                <div class="reminder-form-row">
                    <label class="reminder-inline-label" for="reminder-repeat">重复</label>
                    <select id="reminder-repeat" v-model="reminderForm.repeatMode" class="reminder-select" name="reminder_repeat">
                        <option value="once">单次</option>
                        <option value="daily">每天</option>
                        <option value="workday">工作日</option>
                        <option value="custom_days">自定义星期</option>
                    </select>
                </div>
                <div v-if="reminderForm.repeatMode === 'custom_days'" class="reminder-form-row reminder-custom-days">
                    <span class="reminder-inline-label">重复日</span>
                    <div class="reminder-weekday-grid">
                        <label v-for="w in WEEKDAY_CHECKBOXES" :key="w.value" class="reminder-weekday-item">
                            <input
                                type="checkbox"
                                :checked="(reminderForm.repeatCustomDays || []).includes(w.value)"
                                @change="toggleReminderCustomDay(w.value)"
                            />
                            {{ w.label }}
                        </label>
                    </div>
                </div>
                <button
                    class="btn btn-accent btn-sm"
                    style="width:100%;"
                    type="button"
                    :disabled="reminderSaving"
                    @click="submitReminder"
                >{{ reminderSaving ? '保存中…' : '+ 手动添加提醒' }}</button>
            </div>
            <div class="reminder-list reminder-board">
                <p v-if="!userStore.user?.email" class="todo-drawer-guest">登录后查看与新增提醒。</p>
                <template v-else>
                    <div class="reminders-board-tabs" role="tablist">
                        <button
                            type="button"
                            class="reminders-board-tab"
                            :class="{ 'reminders-board-tab--active': remindersBoardTab === 'pending' }"
                            role="tab"
                            :aria-selected="remindersBoardTab === 'pending'"
                            @click="remindersBoardTab = 'pending'"
                        >
                            待推送
                            <span class="reminders-board-tab-count">{{ remindersPendingBoard.length }}</span>
                        </button>
                        <button
                            type="button"
                            class="reminders-board-tab"
                            :class="{ 'reminders-board-tab--active': remindersBoardTab === 'recurring' }"
                            role="tab"
                            :aria-selected="remindersBoardTab === 'recurring'"
                            @click="remindersBoardTab = 'recurring'"
                        >
                            循环任务
                            <span class="reminders-board-tab-count">{{ remindersRecurringBoard.length }}</span>
                        </button>
                        <button
                            type="button"
                            class="reminders-board-tab"
                            :class="{ 'reminders-board-tab--active': remindersBoardTab === 'pushed' }"
                            role="tab"
                            :aria-selected="remindersBoardTab === 'pushed'"
                            @click="remindersBoardTab = 'pushed'"
                        >
                            已推送记录
                            <span class="reminders-board-tab-count">{{ remindersPushedBoard.length }}</span>
                        </button>
                    </div>
                    <div v-if="remindersBoardLoading" class="todo-drawer-loading">加载中…</div>
                    <div
                        v-else-if="remindersBoardTab === 'pending' && remindersPendingBoard.length === 0"
                        class="todo-drawer-empty"
                    >暂无待推送项（未到点或已全部推送的提醒见「已推送记录」）</div>
                    <div
                        v-else-if="remindersBoardTab === 'recurring' && remindersRecurringBoard.length === 0"
                        class="todo-drawer-empty"
                    >暂无进行中的循环任务</div>
                    <div
                        v-else-if="remindersBoardTab === 'pushed' && remindersPushedBoard.length === 0"
                        class="todo-drawer-empty"
                    >暂无已推送记录</div>
                    <ul v-else class="reminder-alarm-cards reminder-board-cards">
                        <li
                            v-for="t in remindersBoardListActive"
                            :key="`rb-${t.id}`"
                            class="todo-alarm-card reminder-alarm-card reminder-board-card"
                        >
                            <div class="todo-alarm-card__time">{{ formatTodoClock(t.remindAt) }}</div>
                            <div class="todo-alarm-card__body">
                                <div class="todo-alarm-card__head">
                                    <strong class="reminder-card-title">{{ t.title }}</strong>
                                    <span
                                        v-if="String(t.repeatMode || 'once').toLowerCase() !== 'once'"
                                        class="todo-mode-tag todo-mode-tag--repeat-banner"
                                    >🔁 循环提醒</span>
                                    <span v-else class="todo-mode-tag todo-mode-tag--once">单次</span>
                                </div>
                                <span class="reminder-card-meta">{{ formatLocalDateTimeWithRelative(t.remindAt) }}{{ t.source ? ' · ' + t.source : '' }}</span>
                                <p v-if="t.notes" class="todo-drawer-notes">{{ t.notes }}</p>
                                <div
                                    v-if="(remindersBoardTab === 'pending' || remindersBoardTab === 'recurring') && Number(t.done) === 0"
                                    class="reminder-board-actions"
                                >
                                    <button type="button" class="btn btn-accent btn-xs" title="调整提醒时间" @click="openReminderTimeEdit(t)">
                                        ⏰ 调整时间
                                    </button>
                                    <button type="button" class="btn btn-outline btn-xs" @click="markTodoDoneRow(t)">完成</button>
                                    <button
                                        type="button"
                                        class="btn btn-ghost btn-xs todo-card-delete-btn"
                                        title="删除待办"
                                        aria-label="删除待办"
                                        @click="confirmDeleteTodoRow(t)"
                                    >
                                        <el-icon :size="16"><Delete /></el-icon>
                                    </button>
                                </div>
                            </div>
                        </li>
                    </ul>
                </template>
            </div>
        </div>
    </div>
</div>

<el-dialog v-model="reminderTimeEditOpen" title="⏰ 调整提醒时间" width="440px" append-to-body destroy-on-close>
    <p class="reminder-time-edit-hint">选择新的日期与时间（可早于或晚于当前设定），保存后将写入主动提醒调度并重置外联推送状态。</p>
    <el-date-picker
        v-model="reminderTimeEditDate"
        type="datetime"
        placeholder="选择日期时间"
        style="width: 100%"
        format="YYYY-MM-DD HH:mm"
    />
    <template #footer>
        <el-button @click="reminderTimeEditOpen = false">取消</el-button>
        <el-button type="primary" :loading="reminderTimeEditSaving" @click="saveReminderTimeEdit">保存</el-button>
    </template>
</el-dialog>

<!-- ===== TAB: API CONFIG (新增API配置) ===== -->
<div v-show="currentTab === 'apiConfig'" class="tab-content tab-api-config" id="tabApiConfig">
    <p class="api-config-intro">对话与模板填充将自动选用「已启用且优先级最高」的一条配置；数字越大优先级越高。列表中的 Key 为脱敏展示。【模型 ID】须与百炼 Code 一致（小写，如 <code>glm-5.1</code>、<code>qwen-plus</code>）；保存与请求时会自动转为小写。展示名可另填「GLM-5.1」等便于识别。请使用顶部栏「新增配置」；列表会在打开本页时自动加载。</p>

    <!-- ASR：界面隐藏；配置仍由 loadWhisperAsrFromStorage / saveWhisperAsrToStorage 写入 localStorage，会议纪要页可读 -->
    <div v-if="false" class="api-config-panel-card">
        <h3 class="api-config-panel-title">语音识别 (ASR) 配置</h3>
        <p class="api-config-panel-desc">
            与上方「文本对话模型」<strong>完全分离</strong>。会议纪要使用 <strong>全双工 WebSocket</strong> 连接 DashScope Qwen3 实时 ASR；麦克风 WebM 由服务端
            <strong>FFmpeg</strong> 转为 16kHz PCM 后转发。以下 Key 与 URL 仅存本机，在会议页建立 <code>/api/ws/realtime-asr</code> 时由浏览器发给网关。
        </p>
        <div class="api-config-panel-grid">
            <div class="api-config-panel-field api-config-panel-field--full">
                <label class="api-config-panel-label" for="realtime-asr-ws">ASR WebSocket URL（实时 · Qwen3）</label>
                <input
                    id="realtime-asr-ws"
                    v-model="whisperAsrRealtimeWsUrl"
                    name="realtime_asr_ws_url"
                    type="url"
                    class="api-config-panel-input"
                    :placeholder="DEFAULT_REALTIME_ASR_WS"
                    autocomplete="url"
                />
            </div>
            <div class="api-config-panel-field">
                <label class="api-config-panel-label" for="realtime-asr-key">ASR API Key（实时 · 百炼）</label>
                <input
                    id="realtime-asr-key"
                    v-model="whisperAsrRealtimeApiKey"
                    name="realtime_asr_api_key"
                    type="password"
                    class="api-config-panel-input"
                    placeholder="填写 DashScope / 百炼 API Key"
                    autocomplete="new-password"
                />
            </div>
        </div>
        <div class="api-config-panel-actions">
            <button type="button" class="btn btn-primary btn-sm" @click="saveWhisperAsrToStorage">保存 ASR 配置到本机</button>
            <span class="api-config-panel-note">配置写入浏览器 localStorage；会议页实时录音依赖本配置。</span>
        </div>
    </div>

    <div v-if="modelConfigsLoading" class="model-config-status">加载配置中…</div>
    <div v-else-if="modelConfigsError" class="model-config-error">{{ modelConfigsError }}</div>
    <div v-else class="model-config-flow">
        <div
            v-for="row in modelConfigs"
            :key="row.id"
            class="api-config-panel-card model-config-card"
            :class="{ 'is-off': !row.is_enabled }"
        >
            <div class="api-config-panel-head mcc-head">
                <div class="api-config-panel-head-icon" aria-hidden="true">
                    <el-icon :size="22"><Setting /></el-icon>
                </div>
                <div class="api-config-panel-head-text">
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
                <div class="mcc-line">
                    <span class="mcc-k mcc-k--ico"><el-icon :size="14"><Link /></el-icon> Base</span>
                    <span class="mcc-v" :title="row.base_url">{{ row.base_url }}</span>
                </div>
                <div class="mcc-line">
                    <span class="mcc-k mcc-k--ico"><el-icon :size="14"><Key /></el-icon> Key</span>
                    <span class="mcc-v">{{ row.has_api_key ? row.api_key_masked : '未设置' }}</span>
                </div>
                <div class="mcc-line mcc-pri-line">
                    <label class="mcc-k mcc-k--label" :for="'mcc-priority-' + row.id">优先级</label>
                    <input
                        :id="'mcc-priority-' + row.id"
                        v-model.number="row.priority"
                        class="mcc-priority-input api-config-panel-input api-config-panel-input--inline"
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
        <div v-if="modelConfigs.length === 0" class="model-config-empty">暂无模型配置。请点击右上角「新增配置」，或在首次启动时编辑数据库中的默认占位项并填写密钥。</div>
    </div>
</div>

<!-- ===== TAB: 外联推送配置 ===== -->
<div v-show="currentTab === 'webhookConfig'" class="tab-content webhook-config-tab" id="tabWebhookConfig">
    <p class="webhook-config-lead">
        待办到达提醒时间后，服务端会按下方 Webhook 向<strong>企业微信</strong>与<strong>飞书</strong>推送（轮询间隔由服务器环境变量
        <code>TODO_PUSH_POLL_MS</code> 控制，默认约 10 秒）。地址保存在服务器数据库，无需 <code>.env</code>；未填写则静默跳过对应渠道。
    </p>
    <div v-if="webhookConfigLoading" class="webhook-config-loading">加载配置中…</div>
    <div v-else class="webhook-config-grid">
            <div class="webhook-card webhook-card--wecom">
            <div class="webhook-card-head">
                <span class="webhook-card-icon webhook-card-icon--ep" aria-hidden="true"><el-icon :size="22"><Briefcase /></el-icon></span>
                <div>
                    <h3 class="webhook-card-title">企业微信机器人 (WeCom)</h3>
                    <p class="webhook-card-sub">群机器人 Webhook，JSON 文本消息</p>
                </div>
            </div>
            <label class="webhook-field-label" for="webhook-wecom-url">Webhook URL</label>
            <div class="webhook-field-row">
                <input
                    id="webhook-wecom-url"
                    v-model="webhookWecomUrl"
                    class="webhook-field-input"
                    type="url"
                    name="webhook_wecom_url"
                    placeholder="https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=..."
                    autocomplete="off"
                />
                <button
                    type="button"
                    class="btn btn-outline btn-sm webhook-test-btn"
                    :disabled="!String(webhookWecomUrl || '').trim() || webhookTestWecomLoading"
                    @click="testWebhookPush('wecom')"
                >
                    <span v-if="webhookTestWecomLoading">发送中…</span>
                    <span v-else class="webhook-test-btn-inner"><el-icon :size="15"><Promotion /></el-icon> 测试发送</span>
                </button>
            </div>
        </div>
            <div class="webhook-card webhook-card--feishu">
            <div class="webhook-card-head">
                <span class="webhook-card-icon webhook-card-icon--ep" aria-hidden="true"><el-icon :size="22"><Comment /></el-icon></span>
                <div>
                    <h3 class="webhook-card-title">飞书机器人 (Feishu)</h3>
                    <p class="webhook-card-sub">自定义机器人 Webhook</p>
                </div>
            </div>
            <label class="webhook-field-label" for="webhook-feishu-url">Webhook URL</label>
            <div class="webhook-field-row">
                <input
                    id="webhook-feishu-url"
                    v-model="webhookFeishuUrl"
                    class="webhook-field-input"
                    type="url"
                    name="webhook_feishu_url"
                    placeholder="https://open.feishu.cn/open-apis/bot/v2/hook/..."
                    autocomplete="off"
                />
                <button
                    type="button"
                    class="btn btn-outline btn-sm webhook-test-btn"
                    :disabled="!String(webhookFeishuUrl || '').trim() || webhookTestFeishuLoading"
                    @click="testWebhookPush('feishu')"
                >
                    <span v-if="webhookTestFeishuLoading">发送中…</span>
                    <span v-else class="webhook-test-btn-inner"><el-icon :size="15"><Promotion /></el-icon> 测试发送</span>
                </button>
            </div>
        </div>
    </div>
    <div class="webhook-config-footer">
        <button type="button" class="btn btn-primary" :disabled="webhookSaveLoading" @click="saveWebhookConfigs">
            {{ webhookSaveLoading ? '保存中…' : '保存配置' }}
        </button>
    </div>
</div>

    </div>

<el-dialog
    v-model="welcomeWizardOpen"
    title="欢迎来到会意助手"
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
                <option value="openclaw">会意助手本地网关</option>
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
    height: 100%;
    min-height: 0;
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
    min-height: 0;
    height: 100%;
    overflow-y: auto;
    border-radius: 0;
    border-left: 1px solid #e0e0e0;
    box-shadow: -2px 0 15px rgba(0, 0, 0, 0.03);
    padding: 24px 20px;
    background: #ffffff;
    position: sticky;
    top: 0;
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

.identity-persona-icon-wrap {
    display: flex;
    align-items: center;
    justify-content: center;
    margin-bottom: 6px;
    min-height: 2rem;
}

.identity-persona-elicon {
    color: #6366f1;
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
    height: 100%;
    min-height: 0;
}

#tabTemplates {
    height: auto;
    flex: 1 1 auto;
    min-height: 0;
    display: flex;
    flex-direction: column;
    /* 未加载模板时内容可能高于一屏，允许本页内部滚动；加载后由子容器自己吃掉高度 */
    overflow-y: auto;
    overflow-x: hidden;
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
    height: auto;
    flex: 1 1 auto;
    min-height: 0;
}

.panel-error {
    margin-top: 8px;
    color: var(--danger);
    font-size: 0.8rem;
}

.dashboard-reminder-bar {
    display: flex;
    flex-wrap: wrap;
    align-items: flex-start;
    gap: 10px 14px;
    padding: 12px 20px;
    margin: 0 20px 12px;
    border-radius: var(--radius-lg);
    border: 1px solid #e0e7ff;
    background: linear-gradient(92deg, #eef2ff, #fff);
    font-size: 0.78rem;
}

.dashboard-reminder-main {
    flex: 1;
    min-width: 0;
    cursor: pointer;
    border-radius: 10px;
    padding: 4px 6px;
    margin: -4px -6px;
    transition: background 0.15s ease;
}

.dashboard-reminder-main:hover {
    background: rgba(255, 255, 255, 0.65);
}

.dashboard-reminder-count {
    display: block;
    font-weight: 700;
    font-size: 0.92rem;
    color: #1e1b4b;
    margin-top: 4px;
}

.dashboard-reminder-marquee {
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    margin-top: 8px;
    align-items: center;
}

.dashboard-reminder-more {
    font-size: 0.72rem;
    color: #64748b;
}

.dashboard-reminder-hint {
    display: block;
    margin-top: 6px;
    font-size: 0.7rem;
    color: #6366f1;
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

.dashboard-reminder-actions {
    margin-left: auto;
    display: flex;
    align-items: center;
    gap: 8px;
}

.reminder-form {
    margin-top: 12px;
    display: flex;
    flex-direction: column;
    gap: 10px;
    padding: 12px;
    border: 1px solid var(--border-light);
    border-radius: var(--radius);
    background: var(--surface-alt);
}

.reminder-inline-label {
    font-size: 0.78rem;
    font-weight: 600;
    color: #374151;
    margin-bottom: 4px;
}

.reminder-select {
    width: 100%;
    max-width: 280px;
    padding: 8px 10px;
    border-radius: 10px;
    border: 1px solid var(--border-light);
    background: #fff;
    font-size: 0.85rem;
}

.reminder-form-row {
    display: flex;
    flex-direction: column;
    gap: 6px;
}

.reminder-input {
    width: 100%;
    padding: 10px 12px;
    border-radius: 8px;
    border: 1px solid var(--border-light);
    background: #fff;
    font-size: 0.85rem;
}

.reminder-textarea {
    resize: vertical;
}

.reminder-list {
    margin-top: 14px;
}

.reminder-group {
    margin-top: 10px;
}

.reminder-group-title {
    font-weight: 700;
    font-size: 0.82rem;
    color: #1f2937;
    margin-bottom: 6px;
}

.reminder-items {
    display: flex;
    flex-direction: column;
    gap: 8px;
}

.reminder-items li {
    display: flex;
    flex-direction: column;
    gap: 4px;
    padding: 10px 12px;
    border-radius: 10px;
    border: 1px solid #e5e7eb;
    background: #fff;
}

.reminder-items li span {
    color: var(--text-muted);
    font-size: 0.78rem;
}

.reminder-alarm-cards {
    list-style: none;
    margin: 0;
    padding: 0;
    display: flex;
    flex-direction: column;
    gap: 12px;
}

.reminder-alarm-card {
    margin: 0;
}

.reminder-card-title {
    font-size: 0.95rem;
    font-weight: 650;
    color: #0f172a;
}

.reminder-card-meta {
    font-size: 0.72rem;
    color: #64748b;
    margin-top: 6px;
    display: block;
}

.todo-alarm-card {
    display: flex;
    gap: 12px;
    align-items: stretch;
    padding: 16px 16px 16px 14px;
    border-radius: 16px;
    border: 1px solid #f0f2f5;
    background: linear-gradient(145deg, #ffffff 0%, #f8fafc 100%);
    box-shadow: 0 2px 10px rgba(15, 23, 42, 0.05);
    transition: box-shadow 0.2s ease, border-color 0.2s ease, transform 0.2s ease;
}

.todo-alarm-card:hover {
    box-shadow: 0 12px 28px rgba(15, 23, 42, 0.12);
    border-color: #e8eaef;
}

.todo-alarm-card__time {
    flex-shrink: 0;
    min-width: 72px;
    font-size: 1.65rem;
    font-weight: 750;
    letter-spacing: -0.03em;
    font-variant-numeric: tabular-nums;
    color: #0f172a;
    line-height: 1.1;
    padding-top: 2px;
}

.todo-alarm-card__body {
    flex: 1;
    min-width: 0;
}

.todo-alarm-card__head {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 8px;
}

.todo-mode-tag {
    display: inline-flex;
    align-items: center;
    padding: 2px 10px;
    border-radius: 999px;
    font-size: 0.68rem;
    font-weight: 700;
    letter-spacing: 0.02em;
    background: #e2e8f0;
    color: #475569;
}

.todo-mode-tag--daily {
    background: #d1fae5;
    color: #047857;
}

.todo-mode-tag--workday {
    background: #fef3c7;
    color: #b45309;
}

.todo-mode-tag--custom {
    background: #e0e7ff;
    color: #4338ca;
}

/* 滚动条已全局隐藏，这里只保留稳定内边距，避免卡片右缘贴边 */
.todo-drawer :deep(.el-drawer__body) {
    padding-right: 18px;
}

.todo-drawer-status-tag {
    display: inline-flex;
    align-items: center;
    padding: 2px 10px;
    border-radius: 999px;
    font-size: 0.66rem;
    font-weight: 700;
    letter-spacing: 0.02em;
    background: #e2e8f0;
    color: #475569;
}

.todo-drawer-status-tag--pending {
    background: #dbeafe;
    color: #1d4ed8;
}

.todo-drawer-status-tag--notified {
    background: #ffedd5;
    color: #c2410c;
}

.todo-drawer-status-tag--done {
    background: #d1fae5;
    color: #047857;
}

.todo-drawer-status-tag--repeat {
    background: #ede9fe;
    color: #5b21b6;
}

.todo-drawer-status-tag--repeat_cleared {
    background: #ecfdf5;
    color: #047857;
    border: 1px solid rgba(4, 120, 87, 0.18);
}

.todo-mode-tag--repeat-banner {
    background: linear-gradient(135deg, #ede9fe 0%, #e9d5ff 100%);
    color: #6b21a8;
    border: 1px solid rgba(107, 33, 168, 0.2);
}

.todo-mode-tag--once {
    background: #f1f5f9;
    color: #64748b;
}

.paper-potential-todos {
    margin-top: 14px;
    padding: 12px 14px;
    border-radius: 12px;
    border: 1px dashed rgba(79, 70, 229, 0.35);
    background: linear-gradient(120deg, rgba(238, 242, 255, 0.95), rgba(250, 245, 255, 0.9));
    display: flex;
    flex-direction: column;
    align-items: stretch;
    gap: 10px;
}

.paper-potential-todos__list {
    list-style: none;
    margin: 0;
    padding: 0;
    display: flex;
    flex-direction: column;
    gap: 6px;
    max-height: 220px;
    overflow-y: auto;
}

.paper-potential-todos__row {
    padding: 6px 8px;
    border-radius: 8px;
    background: rgba(255, 255, 255, 0.72);
    border: 1px solid #f0f2f5;
}

.paper-potential-todos__title {
    font-weight: 600;
    color: #1e293b;
}

.paper-potential-todos__time {
    display: block;
    margin-top: 2px;
    font-size: 0.72rem;
    color: #64748b;
    font-weight: 500;
}

.paper-potential-todos__tip {
    margin: 0;
    font-size: 0.88rem;
    font-weight: 600;
    color: #4338ca;
}

.longread-showcase {
    margin-top: 28px;
    padding: 22px 24px;
    display: flex;
    gap: 18px;
    align-items: flex-start;
    border-radius: 16px;
    border: 1px solid rgba(99, 102, 241, 0.2);
    background: linear-gradient(135deg, #faf5ff 0%, #eef2ff 55%, #f0fdfa 100%);
    box-shadow: 0 10px 36px rgba(79, 70, 229, 0.08);
}

.longread-showcase__icon {
    flex-shrink: 0;
    width: 52px;
    height: 52px;
    border-radius: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: rgba(99, 102, 241, 0.12);
    color: #4f46e5;
}

.longread-showcase__title {
    margin: 0 0 8px;
    font-size: 1.15rem;
    font-weight: 800;
    letter-spacing: -0.02em;
    color: #1e1b4b;
}

.longread-showcase__lead {
    margin: 0 0 10px;
    font-size: 0.92rem;
    line-height: 1.55;
    color: #334155;
}

.longread-showcase__bullets {
    margin: 0 0 14px;
    padding-left: 1.1rem;
    font-size: 0.86rem;
    line-height: 1.55;
    color: #475569;
}

.longread-showcase__cta {
    border-radius: 10px !important;
}

.reminder-time-edit-hint {
    margin: 0 0 12px;
    font-size: 0.82rem;
    line-height: 1.5;
    color: #64748b;
}

.todo-drawer-item--card.todo-alarm-card {
    box-shadow: 0 8px 28px rgba(15, 23, 42, 0.09);
    border-color: #f0f2f5;
}

.todo-drawer-item--card.todo-alarm-card:hover {
    box-shadow: 0 14px 36px rgba(15, 23, 42, 0.14);
}

.reminder-custom-days {
    flex-direction: column;
    align-items: stretch;
    gap: 8px;
}

.reminder-weekday-grid {
    display: flex;
    flex-wrap: wrap;
    gap: 8px 14px;
}

.reminder-weekday-item {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    font-size: 0.82rem;
    color: #334155;
    cursor: pointer;
    user-select: none;
}

.reminders-board-tabs {
    display: flex;
    gap: 8px;
    margin-bottom: 14px;
    flex-wrap: wrap;
}

.reminders-board-tab {
    border: 1px solid #e2e8f0;
    background: #f8fafc;
    color: #475569;
    padding: 8px 14px;
    border-radius: 999px;
    font-size: 0.82rem;
    font-weight: 600;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 6px;
    transition: background 0.15s ease, border-color 0.15s ease, color 0.15s ease;
}

.reminders-board-tab--active {
    background: #eff6ff;
    border-color: #93c5fd;
    color: #1e40af;
}

.reminders-board-tab-count {
    font-size: 0.72rem;
    font-weight: 700;
    opacity: 0.85;
}

.reminder-board-cards {
    gap: 12px;
}

.reminder-board-card {
    flex-wrap: wrap;
}

.reminder-board-actions {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin-top: 12px;
}

.chat-history-sidebar-toolbar {
    display: flex;
    align-items: center;
    gap: 8px;
}

.chat-history-collapse-btn {
    flex-shrink: 0;
    width: 36px;
    height: 36px;
    border-radius: 10px;
    border: 1px solid #e2e8f0;
    background: #fff;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    color: #475569;
}

.chat-history-collapse-btn:hover {
    border-color: #93c5fd;
    color: #1d4ed8;
    background: #eff6ff;
}

.chat-new-session-btn--toolbar {
    flex: 1;
    min-width: 0;
}

.chat-history-resizer {
    width: 6px;
    flex-shrink: 0;
    cursor: col-resize;
    background: transparent;
    align-self: stretch;
    position: relative;
}

.chat-history-resizer:hover {
    background: rgba(59, 130, 246, 0.12);
}

.chat-history-expand-rail {
    flex-shrink: 0;
    width: 36px;
    align-self: stretch;
    min-height: 120px;
    border: none;
    border-right: 1px solid #e2e8f0;
    background: #fff;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #64748b;
    padding: 0;
}

.chat-history-expand-rail:hover {
    background: #f1f5f9;
    color: #1d4ed8;
}

.todo-snooze-inner {
    display: inline-flex;
    align-items: center;
    gap: 6px;
}

.todo-snooze-btn {
    font-weight: 650;
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

.panel-stream {
    margin-top: 12px;
    border: 1px dashed #cbd5e1;
    border-radius: 12px;
    padding: 10px 12px;
    background: #f8fafc;
}

.panel-stream-title {
    font-size: 0.78rem;
    font-weight: 700;
    color: #334155;
    margin-bottom: 6px;
}

.panel-stream-body {
    margin: 0;
    font-size: 0.78rem;
    line-height: 1.5;
    color: #475569;
    max-height: 180px;
    overflow: auto;
    white-space: pre-wrap;
    word-break: break-word;
}

.template-folder-actions {
    display: flex;
    justify-content: center;
    margin: 0.75rem 0 0.25rem;
}

.word-upload-input--folder {
    display: none;
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

.webhook-config-tab {
    padding: 8px 12px 32px;
    max-width: 960px;
}

.webhook-config-lead {
    font-size: 0.86rem;
    color: #475569;
    line-height: 1.55;
    margin: 0 0 18px;
}

.webhook-config-lead code {
    font-size: 0.78rem;
    background: #f1f5f9;
    padding: 1px 6px;
    border-radius: 4px;
}

.webhook-config-loading {
    padding: 24px;
    color: var(--text-muted);
    font-size: 0.9rem;
}

.webhook-config-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
    gap: 18px;
    margin-bottom: 22px;
}

.webhook-card {
    border-radius: var(--radius-lg);
    padding: 20px 18px 18px;
    border: 1px solid var(--border-light);
    box-shadow: var(--shadow-sm);
    display: flex;
    flex-direction: column;
    gap: 12px;
    transition: var(--transition);
}

.webhook-card--wecom {
    background: linear-gradient(160deg, #eff6ff 0%, #ffffff 55%);
    border-color: #bfdbfe;
}

.webhook-card--feishu {
    background: linear-gradient(160deg, #f0fdf4 0%, #ffffff 55%);
    border-color: #bbf7d0;
}

.webhook-card:hover {
    border-color: #94a3b8;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.08);
}

.webhook-card-head {
    display: flex;
    align-items: flex-start;
    gap: 12px;
}

.webhook-card-icon {
    font-size: 1.75rem;
    line-height: 1;
}

.webhook-card-icon--ep {
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    color: #6366f1;
}

.webhook-test-btn-inner {
    display: inline-flex;
    align-items: center;
    gap: 6px;
}

.webhook-card-title {
    margin: 0 0 4px;
    font-size: 1.05rem;
    font-weight: 700;
    color: #0f172a;
}

.webhook-card-sub {
    margin: 0;
    font-size: 0.78rem;
    color: #64748b;
}

.webhook-field-label {
    font-size: 0.76rem;
    font-weight: 600;
    color: #334155;
}

.webhook-field-row {
    display: flex;
    align-items: stretch;
    gap: 10px;
    flex-wrap: wrap;
}

.webhook-field-row .webhook-field-input {
    flex: 1;
    min-width: 0;
}

.webhook-field-input {
    width: 100%;
    box-sizing: border-box;
    padding: 10px 12px;
    border-radius: var(--radius);
    border: 1px solid var(--border);
    font-size: 0.82rem;
    font-family: ui-monospace, system-ui, sans-serif;
}

.webhook-field-row > .webhook-field-input {
    width: auto;
}

.webhook-test-btn {
    flex-shrink: 0;
    align-self: stretch;
    white-space: nowrap;
}

.webhook-config-footer {
    display: flex;
    flex-wrap: wrap;
    gap: 12px;
    align-items: center;
}

/** API 配置页：统一卡片表面（与旧 ASR 卡一致；模型列表与隐藏 ASR 共用） */
.api-config-panel-card {
    padding: 18px 18px 16px;
    border-radius: var(--radius-lg);
    border: 1px solid #c7d2fe;
    background: linear-gradient(165deg, #f8fafc 0%, #eef2ff 100%);
}

.api-config-panel-title {
    margin: 0 0 8px;
    font-size: 1.05rem;
    font-weight: 700;
    color: #1e293b;
}

.api-config-panel-desc {
    margin: 0 0 14px;
    font-size: 0.82rem;
    color: #475569;
    line-height: 1.55;
}

.api-config-panel-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 14px;
}

@media (max-width: 768px) {
    .api-config-panel-grid {
        grid-template-columns: 1fr;
    }
}

.api-config-panel-field {
    display: flex;
    flex-direction: column;
    gap: 6px;
}

.api-config-panel-field--full {
    grid-column: 1 / -1;
}

.api-config-panel-label {
    font-size: 0.78rem;
    font-weight: 600;
    color: #334155;
}

.api-config-panel-input {
    width: 100%;
    box-sizing: border-box;
    padding: 10px 12px;
    border-radius: var(--radius);
    border: 1px solid var(--border);
    font-size: 0.85rem;
    font-family: var(--font-sans, system-ui, sans-serif);
    background: #fff;
}

.api-config-panel-input--inline {
    width: auto;
    min-width: 88px;
}

.api-config-panel-actions {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 10px 14px;
    margin-top: 14px;
}

.api-config-panel-note {
    font-size: 0.75rem;
    color: #64748b;
    line-height: 1.45;
    max-width: 520px;
}

.api-config-panel-head {
    display: flex;
    align-items: flex-start;
    gap: 12px;
    width: 100%;
}

.api-config-panel-head-icon {
    flex-shrink: 0;
    width: 40px;
    height: 40px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: rgba(99, 102, 241, 0.14);
    color: #4338ca;
}

.api-config-panel-head-text {
    flex: 1;
    min-width: 0;
}

.api-config-panel-head.mcc-head .toggle-switch {
    margin-left: auto;
    flex-shrink: 0;
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
    display: flex;
    flex-direction: column;
    gap: 12px;
    transition: border-color 0.18s ease, box-shadow 0.18s ease;
}

.model-config-card:hover {
    border-color: #a5b4fc;
    box-shadow: 0 10px 28px rgba(79, 70, 229, 0.12);
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

.api-config-panel-head.mcc-head {
    align-items: center;
}

.mcc-k--ico {
    display: inline-flex;
    align-items: center;
    gap: 5px;
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
    border: 1px solid #c7d2fe;
    border-radius: var(--radius-lg);
    background: linear-gradient(165deg, #f8fafc 0%, #eef2ff 100%);
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
    min-height: 0;
}

/* 普通 Tab：让可见的 tab-content 吃掉顶栏之外的剩余高度，
   避免出现「内容明明没占满一屏却还能滚动」 */
.main-content.workspace-dashboard-body:not(.dashboard--immersive) {
    display: flex;
    flex-direction: column;
}

.main-content.workspace-dashboard-body:not(.dashboard--immersive) > .tab-content {
    flex: 1 1 auto;
    min-height: 0;
}

/* —— 全屏沉浸：仅对话助手 / 情感陪伴（根节点挂 dashboard--immersive） —— */
.main-content.workspace-dashboard-body {
    flex: 1;
    min-width: 0;
    min-height: 0;
    height: 100%;
    box-sizing: border-box;
    overflow-y: auto;
    overflow-x: hidden;
}

.main-content.workspace-dashboard-body.dashboard--immersive {
    /* 顶侧：与视口对齐；左右底保留 safe-area */
    padding: 0 max(8px, env(safe-area-inset-right, 0px)) max(6px, env(safe-area-inset-bottom, 0px))
        max(8px, env(safe-area-inset-left, 0px)) !important;
    margin-top: 0 !important;
    margin-right: 0 !important;
    margin-bottom: 0 !important;
    position: relative;
    flex: 1;
    min-height: 0;
    height: auto;
    max-height: none;
    box-sizing: border-box;
    display: flex;
    flex-direction: column;
    overflow: hidden;
    background: #ffffff;
    /*
     * #mainContent 下大量 Tab 为兄弟节点，模板中的换行/缩进会变成「仅空白」的文本节点；
     * 在 column flex 里每个文本节点都会单独成行并占用 line-height，叠成顶部大块留白。
     */
    line-height: 0;
}

/* 仅恢复直接子元素的行高，避免影响内部排版上下文 */
.main-content.workspace-dashboard-body.dashboard--immersive > * {
    line-height: 1.6;
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

.main-content.workspace-dashboard-body.dashboard--immersive .tab-chat-workbench.tab-chat-gemini-root {
    padding: 0 max(6px, env(safe-area-inset-left, 0px)) max(6px, env(safe-area-inset-bottom, 0px)) max(6px, env(safe-area-inset-right, 0px)) !important;
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
    flex: 1;
    display: flex;
    flex-direction: column;
    height: 100%;
}

.main-content.workspace-dashboard-body.dashboard--immersive .chat-messages--gemini {
    background: #ffffff;
    flex: 1;
    min-height: 0;
    align-self: stretch;
    width: 100%;
    box-sizing: border-box;
    padding: 8px 16px 120px;
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
    padding-right: 12px;
}

.main-content.workspace-dashboard-body.dashboard--immersive .chat-gemini-toolbar {
    padding: max(6px, env(safe-area-inset-top, 0px)) 12px 8px;
    row-gap: 4px;
}

.main-content.workspace-dashboard-body.dashboard--immersive .companion-chat-shell .chat-header {
    padding-top: max(8px, env(safe-area-inset-top, 0px));
    padding-bottom: 10px;
}

.main-content.workspace-dashboard-body.dashboard--immersive .companion-messages {
    flex: 1;
    min-height: 0;
    overflow-y: auto;
    padding-top: 8px;
}

/* 沉浸 Tab 的浮动待办铃铛已移除：它被 .topbar-todo-bell 的 position:relative 覆盖，
   既没固定住又在流内占了 38px，导致对话/陪伴页顶部多出一截空白 */

.topbar-actions--end {
    margin-left: auto;
    display: flex;
    align-items: center;
    justify-content: flex-end;
    flex-wrap: nowrap;
    gap: 16px;
}

.header-page-actions {
    display: flex;
    align-items: center;
    justify-content: flex-end;
    gap: 12px;
    min-height: 0;
}

.header-api-add-btn {
    flex-shrink: 0;
}

.topbar-user-chip {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    max-width: 200px;
    padding: 4px 10px 4px 4px;
    border-radius: 999px;
    border: 1px solid #e2e8f0;
    background: #fff;
    cursor: pointer;
    color: #334155;
    font-size: 0.82rem;
    font-weight: 600;
}

.topbar-user-chip:hover {
    border-color: #cbd5e1;
    background: #f8fafc;
}

.topbar-user-avatar {
    flex-shrink: 0;
    width: 32px;
    height: 32px;
    border-radius: 999px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    background: linear-gradient(135deg, #6366f1, #8b5cf6);
    color: #fff;
    font-size: 0.85rem;
    font-weight: 700;
}

.topbar-user-name {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}

.todo-card-delete-btn {
    padding: 4px 8px;
    min-width: 32px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    color: #64748b;
}

.todo-card-delete-btn:hover {
    color: #dc2626;
    background: #fef2f2;
}

.topbar-todo-bell {
    position: relative;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 38px;
    height: 38px;
    margin-right: 0;
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
    font-size: 18px;
    line-height: 1;
    display: inline-flex;
    align-items: center;
    justify-content: center;
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
    gap: 12px;
    align-items: flex-start;
    padding: 0;
    border: none;
    background: transparent;
    box-shadow: none;
}

.todo-drawer-item--done.todo-alarm-card {
    opacity: 0.55;
    filter: grayscale(0.25);
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
    min-height: 0;
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
    flex-direction: column;
    align-items: flex-start;
    gap: 3px;
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
    width: 100%;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    min-width: 0;
}

.chat-history-meta {
    max-width: calc(100% - 28px);
    display: flex;
    flex-direction: column;
    align-items: flex-start;
    gap: 2px;
    font-size: 0.68rem;
    color: #64748b;
}

.chat-history-meta-line {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    max-width: 100%;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}

.chat-history-meta-ico {
    flex-shrink: 0;
    color: #6366f1;
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
    display: inline-flex;
    align-items: center;
    justify-content: center;
    color: #94a3b8;
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

.chat-gemini-pill-ico {
    flex-shrink: 0;
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
    align-items: center;
    justify-content: flex-end;
    gap: 12px;
    flex-shrink: 0;
}

.companion-chat-header-row {
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 10px;
}

.companion-chat-header-text {
    flex: 1;
    min-width: 0;
}

.companion-header-actions {
    margin-left: auto;
    display: flex;
    align-items: center;
    justify-content: flex-end;
    gap: 12px;
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

/* 生成中：发送按钮切换为「停止」，用中性色与发送态区分 */
.send-btn--stop {
    background: #e2e8f0;
    color: #475569;
    font-size: 12px;
    line-height: 1;
}

.send-btn--stop:hover {
    background: #cbd5e1;
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
    align-items: flex-start;
    min-height: 0;
    padding: 12px 16px 20px;
}

/* el-empty 默认按内容收缩，导致提示语/拖拽区/链接框三条中线不一致 */
.paper-tab-empty :deep(.el-empty) {
    width: 100%;
    padding: 24px 0 8px;
}

.paper-tab-empty :deep(.el-empty__description) {
    margin-bottom: 8px;
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

.btn-ico-text {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
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


/* —— 个人资料库 —— */
.tab-library {
    max-width: 1280px;
    margin: 0 auto;
    padding-bottom: 48px;
}

.lib-folder-inline {
    display: inline-flex;
    align-items: center;
    gap: 4px;
}

.lib-folder-chip {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    margin-right: 10px;
}

.library-hero {
    display: flex;
    justify-content: space-between;
    gap: 20px;
    align-items: center;
    padding: 26px;
    border: 1px solid #e2e8f0;
    border-radius: 22px;
    background: linear-gradient(135deg, #f8fafc 0%, #eefdf8 48%, #fff7ed 100%);
    box-shadow: 0 18px 50px rgba(15, 23, 42, 0.08);
}

.library-kicker {
    margin: 0 0 6px;
    font-size: 0.75rem;
    font-weight: 800;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #0f766e;
}

.library-hero h2 {
    margin: 0 0 8px;
    font-size: 1.8rem;
    letter-spacing: -0.04em;
    color: #0f172a;
}

.library-hero p {
    max-width: 720px;
    margin: 0;
    color: #475569;
    line-height: 1.65;
}

.library-hero-actions,
.library-section-actions,
.library-row-actions {
    display: flex;
    gap: 8px;
    align-items: center;
    flex-wrap: wrap;
}

.library-tabs {
    display: inline-flex;
    gap: 6px;
    margin: 20px 0 14px;
    padding: 6px;
    border: 1px solid #e2e8f0;
    border-radius: 999px;
    background: #fff;
}

.library-tabs button {
    border: 0;
    border-radius: 999px;
    padding: 9px 16px;
    background: transparent;
    color: #64748b;
    font-weight: 700;
    cursor: pointer;
}

.library-tabs button.active {
    background: #0f766e;
    color: #fff;
    box-shadow: 0 8px 20px rgba(15, 118, 110, 0.18);
}

.library-section {
    border: 1px solid #e5e7eb;
    border-radius: 20px;
    padding: 20px;
    background: rgba(255, 255, 255, 0.86);
}

.library-section-head {
    display: flex;
    justify-content: space-between;
    gap: 16px;
    align-items: flex-start;
    margin-bottom: 16px;
}

.library-section-head h3 {
    margin: 0 0 6px;
    color: #0f172a;
}

.library-section-head p {
    margin: 0;
    color: #64748b;
    line-height: 1.55;
}

.library-grid-two {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 16px;
}

.library-card-panel {
    border: 1px solid #e2e8f0;
    border-radius: 16px;
    background: #fff;
    overflow: hidden;
}

.library-card-panel-head {
    display: flex;
    justify-content: space-between;
    gap: 12px;
    align-items: center;
    padding: 14px 16px;
    border-bottom: 1px solid #edf2f7;
    background: #f8fafc;
}

.library-card-panel-head h4 {
    margin: 0;
    color: #0f172a;
}

.library-card-panel-head span,
.library-empty {
    color: #64748b;
    font-size: 0.86rem;
}

.library-empty {
    padding: 22px 16px;
    text-align: center;
}

.library-empty--wide {
    border: 1px dashed #cbd5e1;
    border-radius: 14px;
    background: #f8fafc;
}

.library-list,
.library-file-list {
    list-style: none;
    margin: 0;
    padding: 0;
}

.library-row,
.library-file-row {
    display: flex;
    justify-content: space-between;
    gap: 12px;
    padding: 14px 16px;
    border-bottom: 1px solid #f1f5f9;
}

.library-row:last-child,
.library-file-row:last-child {
    border-bottom: 0;
}

.library-row-main {
    min-width: 0;
    display: flex;
    flex-direction: column;
    gap: 4px;
}

.library-row-main strong,
.library-file-title strong {
    color: #0f172a;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}

.library-row-main span,
.library-file-meta {
    color: #64748b;
    font-size: 0.78rem;
}

.library-row-main em {
    color: #0f766e;
    font-size: 0.76rem;
    font-style: normal;
}

.library-toolbar {
    display: flex;
    gap: 10px;
    margin-bottom: 12px;
}

.library-toolbar input {
    flex: 1;
    min-width: 180px;
    border: 1px solid #dbe3ef;
    border-radius: 12px;
    padding: 10px 12px;
    outline: none;
}

.library-folder-cloud {
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    margin: 0 0 14px;
}

.library-folder-cloud span {
    padding: 6px 10px;
    border-radius: 999px;
    background: #ecfdf5;
    color: #065f46;
    font-size: 0.78rem;
}

.library-folder-cloud b {
    margin-left: 4px;
}

.library-file-row {
    align-items: flex-start;
    border: 1px solid #e5e7eb;
    border-radius: 16px;
    margin-bottom: 10px;
    background: #fff;
}

.library-file-icon {
    width: 42px;
    height: 42px;
    border-radius: 14px;
    display: grid;
    place-items: center;
    background: #f1f5f9;
    flex: 0 0 auto;
}

.library-file-icon--ep {
    color: #64748b;
}

.library-file-main {
    flex: 1;
    min-width: 0;
}

.library-file-title {
    display: flex;
    gap: 8px;
    align-items: center;
    margin-bottom: 6px;
    flex-wrap: wrap;
}

.library-file-title strong {
    flex: 1;
    min-width: 0;
}

.library-rename-ico {
    border: none;
    background: transparent;
    cursor: pointer;
    padding: 0 4px;
    font-size: 0.82rem;
    line-height: 1;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    color: #94a3b8;
    opacity: 0.65;
}

.library-rename-ico:hover {
    opacity: 1;
}

.library-file-title .library-rename-input {
    flex: 1 1 100%;
    min-width: 0;
}

.library-file-title span {
    flex: 0 0 auto;
    padding: 2px 7px;
    border-radius: 999px;
    background: #eef2ff;
    color: #4f46e5;
    font-size: 0.72rem;
    font-weight: 700;
}

.library-file-excerpt {
    margin: 0 0 8px;
    color: #475569;
    font-size: 0.85rem;
    line-height: 1.55;
    display: -webkit-box;
    -webkit-line-clamp: 2;
    -webkit-box-orient: vertical;
    overflow: hidden;
}

.library-file-meta {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
}

@media (max-width: 900px) {
    .library-hero,
    .library-section-head,
    .library-row,
    .library-file-row {
        flex-direction: column;
    }
    .library-grid-two {
        grid-template-columns: 1fr;
    }
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

.companion-chat-shell .companion-chat-header-row.chat-header {
    align-items: center;
}

.companion-persona-panel {
    padding: 10px 16px 10px 14px;
    border-bottom: 1px solid #fce7f3;
    background: linear-gradient(180deg, #fffdfb, #fffafc);
}

.companion-persona-row {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    gap: 8px;
    min-width: 0;
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

/* 陪伴输入区：整条做成一个圆角盒子，麦克风/发送收进盒子右侧，
   避免输入框自己圆角、按钮却挂在框外的割裂感 */
.companion-input-toolbar {
    align-items: center;
    gap: 8px;
    padding: 7px 8px 7px 16px;
    background: #fff;
    border: 1px solid #fbcfe8;
    border-radius: 22px;
    box-sizing: border-box;
    min-height: 0;
}

.companion-input-toolbar .companion-textarea {
    flex: 1;
    min-width: 0;
    border: none !important;
    border-radius: 0 !important;
    background: transparent;
    padding: 6px 0;
    min-height: 26px;
    max-height: 160px;
    resize: none;
    box-shadow: none !important;
}

.companion-input-toolbar .chat-input-trailing {
    flex: 0 0 auto;
    flex-direction: row;
    align-items: center;
    gap: 6px;
    border: none;
    background: transparent;
    border-radius: 0;
    overflow: visible;
}

.companion-input-toolbar .chat-mic-btn,
.companion-input-toolbar .chat-input-trailing .send-btn {
    flex: 0 0 auto;
    width: 36px;
    height: 36px;
    min-height: 36px;
    border: none !important;
    border-top: none !important;
    border-radius: 12px !important;
}

.companion-input-toolbar .chat-mic-btn {
    background: #fdf2f8;
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
    min-height: 0;
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
    min-height: 0;
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

.chat-like-ico {
    flex-shrink: 0;
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

.chat-like-btn--on {
    border-color: #a855f7 !important;
    background: linear-gradient(135deg, #f5f3ff 0%, #ede9fe 100%) !important;
    color: #5b21b6 !important;
    box-shadow: 0 0 0 1px rgba(168, 85, 247, 0.2);
}

.chat-pack-skill-btn {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    padding: 4px 10px;
    font-size: 0.78rem;
    border-radius: 8px;
    border: 1px solid #fde68a;
    background: linear-gradient(135deg, #fffbeb 0%, #fef3c7 100%);
    color: #92400e;
    cursor: pointer;
    transition: background 0.15s ease, border-color 0.15s ease;
}

.chat-pack-skill-btn:hover:not(:disabled) {
    border-color: #fbbf24;
    background: #fff7ed;
}

.chat-pack-skill-btn:disabled {
    opacity: 0.45;
    cursor: not-allowed;
}

.stat-card--ratings-trend {
    flex: 1 1 280px;
    min-width: 240px;
    align-items: stretch;
    flex-direction: column;
    gap: 10px;
}

.stat-ratings-head {
    display: flex;
    align-items: flex-start;
    gap: 12px;
    width: 100%;
}

.stat-ratings-head-icon {
    width: 52px;
    min-height: 44px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}

.stat-ratings-head-text {
    flex: 1;
    min-width: 0;
}

.stat-ratings-chart-area {
    width: 100%;
    min-height: 132px;
}

.stat-ratings-chart-el {
    width: 100%;
    height: 132px;
}

.stat-ratings-empty {
    width: 100%;
    min-height: 120px;
    display: flex;
    align-items: center;
    justify-content: center;
    text-align: center;
    padding: 12px;
    font-size: 0.78rem;
    color: #64748b;
    background: #f8fafc;
    border-radius: 12px;
    border: 1px dashed #e2e8f0;
}

/* DeepSeek / R1 风格：<think> → details（正文内在 Markdown 流程中注入） */
.markdown-body :deep(.deep-think-box) {
    margin: 10px 0;
    border-radius: 12px;
    border: 1px solid #e2e8f0;
    background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
    overflow: hidden;
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.06);
}

.markdown-body :deep(.deep-think-box > summary) {
    cursor: pointer;
    list-style: none;
    padding: 10px 14px;
    font-size: 0.8rem;
    font-weight: 650;
    color: #475569;
    user-select: none;
}

.markdown-body :deep(.deep-think-box > summary::-webkit-details-marker) {
    display: none;
}

.markdown-body :deep(.think-content) {
    padding: 0 14px 12px;
}

.markdown-body :deep(.think-pre) {
    margin: 0;
    font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, 'Liberation Mono', 'Courier New', monospace;
    font-size: 0.72rem;
    line-height: 1.55;
    white-space: pre-wrap;
    word-break: break-word;
    color: #475569;
    background: rgba(255, 255, 255, 0.75);
    border-radius: 8px;
    padding: 10px 12px;
    border: 1px solid #e2e8f0;
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
