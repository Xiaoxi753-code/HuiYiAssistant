<template>
  <div
    id="workspaceShell"
    class="dashboard-layout active"
    :class="{ 'dashboard-layout--hide-header': hideAppHeader }"
  >
    <!-- 按路由 meta.hideHeader 隐藏全局顶栏占位（高度 0，不占空间） -->
    <el-header v-if="!hideAppHeader" height="0" class="workspace-global-header" />
    <div class="workspace-main-stage">
      <WorkspaceSidebar />
      <router-view />
    </div>
  </div>
</template>

<script setup>
import { computed } from 'vue'
import { useRoute } from 'vue-router'
import { ElHeader } from 'element-plus'
import 'element-plus/es/components/header/style/css'
import WorkspaceSidebar from '../components/WorkspaceSidebar.vue'

const route = useRoute()
const hideAppHeader = computed(() => route.meta.hideHeader === true)
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
}

.dashboard-layout--hide-header {
  min-height: 100vh;
}

.dashboard-layout--hide-header .workspace-main-stage {
  min-height: 100vh;
  flex: 1 1 auto;
}
</style>
