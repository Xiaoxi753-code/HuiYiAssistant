<template>
  <div
    id="workspaceShell"
    class="dashboard-layout active"
    :class="{ 'dashboard-layout--hide-header': hideAppHeader }"
  >
    <el-header v-if="!hideAppHeader" height="0" class="workspace-global-header" />
    <div class="workspace-main-stage">
      <WorkspaceSidebar :collapsed="sidebarCollapsed" @toggle-collapse="toggleSidebarCollapse" />
      <div
        v-show="!sidebarCollapsed"
        class="workspace-sidebar-resizer"
        aria-hidden="true"
        @mousedown.prevent="startSidebarResize"
      />
      <router-view class="workspace-router-outlet" />
    </div>
  </div>
</template>

<script setup>
import { computed, onBeforeUnmount, onMounted, ref, watch } from 'vue'
import { useRoute } from 'vue-router'
import { ElHeader } from 'element-plus'
import 'element-plus/es/components/header/style/css'
import WorkspaceSidebar from '../components/WorkspaceSidebar.vue'

const route = useRoute()
const hideAppHeader = computed(() => route.meta.hideHeader === true)

const SIDEBAR_W_KEY = 'huiyi_sidebar_width_px'
const SIDEBAR_C_KEY = 'huiyi_sidebar_collapsed'

const sidebarCollapsed = ref(false)
const sidebarWidthPx = ref(260)

function applySidebarCssVar() {
  const w = sidebarCollapsed.value ? 72 : sidebarWidthPx.value
  document.documentElement.style.setProperty('--sidebar-width', `${w}px`)
}

onMounted(() => {
  try {
    sidebarCollapsed.value = localStorage.getItem(SIDEBAR_C_KEY) === '1'
    const raw = Number.parseInt(String(localStorage.getItem(SIDEBAR_W_KEY) || ''), 10)
    if (Number.isFinite(raw) && raw >= 200 && raw <= 440) {
      sidebarWidthPx.value = raw
    }
  } catch {
    /* ignore */
  }
  applySidebarCssVar()
})

watch([sidebarCollapsed, sidebarWidthPx], applySidebarCssVar)

function toggleSidebarCollapse() {
  sidebarCollapsed.value = !sidebarCollapsed.value
  try {
    localStorage.setItem(SIDEBAR_C_KEY, sidebarCollapsed.value ? '1' : '0')
  } catch {
    /* ignore */
  }
}

let sidebarResizeActive = false
let sidebarResizeStartX = 0
let sidebarResizeStartW = 260

function onSidebarResizeMouseMove(e) {
  if (!sidebarResizeActive || sidebarCollapsed.value) {
    return
  }
  const delta = e.clientX - sidebarResizeStartX
  const nw = Math.min(440, Math.max(200, sidebarResizeStartW + delta))
  sidebarWidthPx.value = nw
}

function onSidebarResizeMouseUp() {
  if (!sidebarResizeActive) {
    return
  }
  sidebarResizeActive = false
  document.removeEventListener('mousemove', onSidebarResizeMouseMove)
  document.removeEventListener('mouseup', onSidebarResizeMouseUp)
  document.body.style.cursor = ''
  document.body.style.userSelect = ''
  try {
    localStorage.setItem(SIDEBAR_W_KEY, String(sidebarWidthPx.value))
  } catch {
    /* ignore */
  }
}

function startSidebarResize(e) {
  if (sidebarCollapsed.value) {
    return
  }
  sidebarResizeActive = true
  sidebarResizeStartX = e.clientX
  sidebarResizeStartW = sidebarWidthPx.value
  document.addEventListener('mousemove', onSidebarResizeMouseMove)
  document.addEventListener('mouseup', onSidebarResizeMouseUp)
  document.body.style.cursor = 'col-resize'
  document.body.style.userSelect = 'none'
}

onBeforeUnmount(() => {
  document.removeEventListener('mousemove', onSidebarResizeMouseMove)
  document.removeEventListener('mouseup', onSidebarResizeMouseUp)
})
</script>

<style scoped>
#workspaceShell.dashboard-layout.active {
  display: flex;
  flex-direction: column;
  min-height: 100vh;
}

.workspace-global-header {
  height: 0 !important;
  min-height: 0 !important;
  margin: 0 !important;
  padding: 0 !important;
  overflow: hidden !important;
  border: none !important;
  flex-shrink: 0;
  line-height: 0;
}

.workspace-main-stage {
  display: flex;
  flex: 1;
  min-height: 0;
  width: 100%;
  align-items: stretch;
  position: relative;
}

.workspace-router-outlet {
  flex: 1;
  min-width: 0;
  min-height: 0;
}

.workspace-sidebar-resizer {
  position: fixed;
  top: 0;
  bottom: 0;
  left: calc(var(--sidebar-width, 260px) - 3px);
  width: 6px;
  z-index: 101;
  cursor: col-resize;
  background: transparent;
}

.workspace-sidebar-resizer:hover {
  background: rgba(99, 102, 241, 0.18);
}

.dashboard-layout--hide-header {
  min-height: 100vh;
}

.dashboard-layout--hide-header .workspace-main-stage {
  min-height: 100vh;
  flex: 1 1 auto;
}
</style>
