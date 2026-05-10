<script setup>
import { ref, computed, watch, onMounted, onBeforeUnmount, nextTick } from 'vue'
import { ElMessage, ElButton } from 'element-plus'
import 'element-plus/es/components/button/style/css'
import { marked } from 'marked'
import DOMPurify from 'dompurify'
import * as echarts from 'echarts/core'
import { CanvasRenderer } from 'echarts/renderers'
import { LineChart } from 'echarts/charts'
import { GridComponent, TooltipComponent } from 'echarts/components'
import { useUserStore } from '../store/user.js'

echarts.use([CanvasRenderer, LineChart, GridComponent, TooltipComponent])

marked.setOptions({ gfm: true, breaks: true })

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

const userStore = useUserStore()

const chartRef = ref(null)
/** @type {import('echarts/core').EChartsType | null} */
let chartInst = null

const emotionHistory = ref([])
const reportsLoading = ref(false)
const reports = ref([])

const generating = ref(false)
const latestHtml = ref('')

const emotionHasTrend = computed(() => Array.isArray(emotionHistory.value) && emotionHistory.value.length > 0)

function buildChartOption(hist) {
  const list = Array.isArray(hist) ? hist : []
  const isEmpty = list.length === 0
  const dates = isEmpty ? ['—'] : list.map((x) => String(x?.date || ''))
  const stresses = isEmpty
    ? [null]
    : list.map((x) => {
        const n = Number(x?.stress)
        return Number.isFinite(n) ? n : 0
      })
  const moods = isEmpty ? [] : list.map((x) => String(x?.mood || '').trim() || '—')
  return {
    animation: !isEmpty,
    grid: { left: 12, right: 12, top: 16, bottom: 28, containLabel: true },
    tooltip: {
      trigger: 'axis',
      formatter(params) {
        if (isEmpty) {
          return ''
        }
        const arr = Array.isArray(params) ? params : [params]
        const p = arr[0]
        if (!p || p.value == null || p.value === '') {
          return ''
        }
        const i = p.dataIndex
        const mood = moods[i] ?? '—'
        return `${p.axisValueLabel || p.name}<br/>压力值：${p.value}<br/>情绪：${mood}`
      },
    },
    xAxis: {
      type: 'category',
      data: dates,
      boundaryGap: false,
      axisLine: { show: false },
      axisTick: { show: false },
      splitLine: { show: false },
    },
    yAxis: {
      type: 'value',
      min: 0,
      max: 10,
      axisLine: { show: false },
      axisTick: { show: false },
      splitLine: { show: false },
    },
    series: [
      {
        type: 'line',
        smooth: true,
        symbol: isEmpty ? 'none' : 'circle',
        symbolSize: 6,
        data: stresses,
        lineStyle: {
          width: isEmpty ? 0 : 3,
          opacity: isEmpty ? 0 : 1,
          color: isEmpty
            ? 'transparent'
            : {
                type: 'linear',
                x: 0,
                y: 0,
                x2: 1,
                y2: 0,
                colorStops: [
                  { offset: 0, color: '#9fd4ff' },
                  { offset: 1, color: '#ffc0d9' },
                ],
              },
        },
        itemStyle: { color: '#b8d9ff' },
        areaStyle: isEmpty
          ? undefined
          : {
              color: {
                type: 'linear',
                x: 0,
                y: 0,
                x2: 0,
                y2: 1,
                colorStops: [
                  { offset: 0, color: 'rgba(159, 212, 255, 0.2)' },
                  { offset: 1, color: 'rgba(255, 192, 217, 0.02)' },
                ],
              },
            },
      },
    ],
  }
}

function disposeChart() {
  if (chartInst) {
    chartInst.dispose()
    chartInst = null
  }
}

function resizeChart() {
  chartInst?.resize()
}

async function loadProfileEmotion() {
  const email = userStore.user?.email
  if (!email) {
    emotionHistory.value = []
    return
  }
  try {
    const r = await fetch(`/api/user/profile?email=${encodeURIComponent(email)}`)
    const d = await r.json().catch(() => ({}))
    const h = d.success && d.profile?.emotion_history ? d.profile.emotion_history : []
    emotionHistory.value = Array.isArray(h) ? h : []
  } catch {
    emotionHistory.value = []
  }
}

async function loadReports() {
  const email = userStore.user?.email
  if (!email) {
    reports.value = []
    return
  }
  reportsLoading.value = true
  try {
    const r = await fetch(`/api/user/weekly-reports?email=${encodeURIComponent(email)}&limit=80`)
    const d = await r.json().catch(() => ({}))
    reports.value = d.success && Array.isArray(d.items) ? d.items : []
  } catch {
    reports.value = []
  } finally {
    reportsLoading.value = false
  }
}

async function generateReport() {
  const email = userStore.user?.email
  if (!email) {
    ElMessage.warning('请先登录')
    return
  }
  generating.value = true
  latestHtml.value = ''
  try {
    const r = await fetch('/api/user/weekly-review', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email }),
    })
    const data = await r.json().catch(() => ({}))
    if (!r.ok || !data.success || !data.report) {
      throw new Error(data.message || `HTTP ${r.status}`)
    }
    latestHtml.value = renderMarkdown(String(data.report))
    ElMessage.success('本周回顾已生成并归档')
    await loadReports()
    await loadProfileEmotion()
  } catch (e) {
    ElMessage.error(String(e?.message || e || '生成失败'))
  } finally {
    generating.value = false
  }
}

function onWinResize() {
  resizeChart()
}

watch(emotionHistory, () => {
  nextTick(() => {
    const el = chartRef.value
    if (!el) {
      return
    }
    if (!chartInst) {
      chartInst = echarts.init(el)
    }
    chartInst.setOption(buildChartOption(emotionHistory.value), true)
    chartInst.resize()
  })
})

onMounted(async () => {
  await loadProfileEmotion()
  await loadReports()
  nextTick(() => {
    const el = chartRef.value
    if (el) {
      disposeChart()
      chartInst = echarts.init(el)
      chartInst.setOption(buildChartOption(emotionHistory.value), true)
      chartInst.resize()
    }
  })
  if (typeof window !== 'undefined') {
    window.addEventListener('resize', onWinResize)
  }
})

onBeforeUnmount(() => {
  if (typeof window !== 'undefined') {
    window.removeEventListener('resize', onWinResize)
  }
  disposeChart()
})
</script>

<template>
  <div class="main-content growth-dashboard">
    <div class="topbar">
      <div>
        <h1 class="page-title">情绪与成长看板</h1>
        <span class="breadcrumb">首页 / 成长与情绪</span>
      </div>
      <div class="topbar-actions">
        <el-button type="primary" size="small" :loading="generating" :disabled="!userStore.user?.email" @click="generateReport">
          生成本周成长与情绪回溯
        </el-button>
      </div>
    </div>

    <section class="growth-section">
      <h2 class="growth-section-title">情绪趋势</h2>
      <div class="emotion-chart-shell">
        <div ref="chartRef" class="emotion-chart"></div>
        <div v-if="!emotionHasTrend" class="emotion-empty-overlay">
          <p>暂无情绪数据，请先在工作台对话或等待画像建模。</p>
        </div>
      </div>
    </section>

    <section v-if="latestHtml" class="growth-section growth-latest">
      <h2 class="growth-section-title">刚刚生成</h2>
      <div class="markdown-body weekly-review-body" v-html="latestHtml"></div>
    </section>

    <section class="growth-section">
      <h2 class="growth-section-title">历史周回顾</h2>
      <p v-if="reportsLoading" class="growth-muted">加载历史…</p>
      <ul v-else class="growth-report-timeline">
        <li v-for="item in reports" :key="item.id" class="growth-report-item">
          <div class="growth-report-meta">
            <span class="growth-report-week">{{ item.week_label || '—' }}</span>
            <time class="growth-report-time">{{ item.created_at }}</time>
          </div>
          <div class="markdown-body growth-report-body" v-html="renderMarkdown(item.report_body)"></div>
        </li>
      </ul>
      <p v-if="!reportsLoading && reports.length === 0" class="growth-muted">尚无归档报告，点击上方按钮生成第一条。</p>
    </section>
  </div>
</template>

<style scoped>
.growth-dashboard {
  flex: 1;
  min-width: 0;
}
.growth-section {
  margin-bottom: 28px;
}
.growth-section-title {
  font-size: 1rem;
  font-weight: 700;
  margin: 0 0 12px;
  color: var(--text-primary);
}
.emotion-chart-shell {
  position: relative;
  background: var(--surface-alt);
  border: 1px solid var(--border-light);
  border-radius: var(--radius-lg);
  padding: 12px;
}
.emotion-chart {
  width: 100%;
  height: 360px;
}
.emotion-empty-overlay {
  position: absolute;
  inset: 0;
  display: flex;
  align-items: center;
  justify-content: center;
  pointer-events: none;
  background: linear-gradient(180deg, rgba(248, 250, 252, 0.85), rgba(241, 245, 249, 0.92));
  border-radius: var(--radius-lg);
}
.emotion-empty-overlay p {
  margin: 0;
  font-size: 0.88rem;
  color: #64748b;
}
.growth-muted {
  color: var(--text-muted);
  font-size: 0.88rem;
}
.growth-report-timeline {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: 18px;
}
.growth-report-item {
  border-left: 3px solid #c7d2fe;
  padding-left: 14px;
}
.growth-report-meta {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  align-items: baseline;
  margin-bottom: 8px;
  font-size: 0.78rem;
  color: var(--text-muted);
}
.growth-report-week {
  font-weight: 700;
  color: var(--text-secondary);
}
.growth-report-body {
  font-size: 0.85rem;
}
.growth-latest {
  padding: 16px;
  border-radius: var(--radius-lg);
  border: 1px solid #e0e7ff;
  background: linear-gradient(165deg, #fafbff 0%, #fff 100%);
}
</style>
