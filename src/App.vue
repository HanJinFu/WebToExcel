<script setup lang="ts">
import ExcelUploader from './components/ExcelUploader.vue'
import ExcelEditor from './components/ExcelEditor.vue'
import { ref } from 'vue'

type TabKey = 'uploader' | 'editor'

const activeTab = ref<TabKey>('uploader')
</script>

<template>
  <div class="app">
    <!-- 顶部导航 -->
    <header class="header">
      <div class="logo">
        <span class="logo-icon">✦</span>
        <span class="logo-text">Excel 工具</span>
      </div>
      <nav class="tabs">
        <button
          :class="['tab-btn', { active: activeTab === 'uploader' }]"
          @click="activeTab = 'uploader'"
        >
          样式移除
        </button>
        <button
          :class="['tab-btn', { active: activeTab === 'editor' }]"
          @click="activeTab = 'editor'"
        >
          编辑器
        </button>
      </nav>
    </header>

    <!-- 页面内容 -->
    <main class="content">
      <transition name="fade" mode="out-in">
        <ExcelUploader v-if="activeTab === 'uploader'" key="uploader" />
        <ExcelEditor v-else key="editor" />
      </transition>
    </main>
  </div>
</template>

<style scoped>
.app {
  min-height: 100vh;
  background: #f5f5f7;
  font-family: -apple-system, BlinkMacSystemFont, 'SF Pro Text', 'Helvetica Neue', sans-serif;
}

/* 顶部导航 */
.header {
  position: sticky;
  top: 0;
  z-index: 100;
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 32px;
  height: 52px;
  background: rgba(255, 255, 255, 0.72);
  backdrop-filter: saturate(180%) blur(20px);
  -webkit-backdrop-filter: saturate(180%) blur(20px);
  border-bottom: 1px solid rgba(0, 0, 0, 0.06);
}

.logo {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 15px;
  font-weight: 600;
  color: #1d1d1f;
  letter-spacing: -0.01em;
}

.logo-icon {
  font-size: 18px;
  color: #0071e3;
}

.tabs {
  display: flex;
  gap: 4px;
  background: rgba(0, 0, 0, 0.04);
  padding: 3px;
  border-radius: 10px;
}

.tab-btn {
  padding: 6px 16px;
  font-size: 13px;
  font-weight: 500;
  color: #6e6e73;
  background: transparent;
  border: none;
  border-radius: 8px;
  cursor: pointer;
  transition: all 0.2s ease;
  letter-spacing: -0.01em;
}

.tab-btn:hover {
  color: #1d1d1f;
}

.tab-btn.active {
  background: #ffffff;
  color: #1d1d1f;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08), 0 1px 2px rgba(0, 0, 0, 0.04);
}

/* 内容区域 */
.content {
  max-width: 960px;
  margin: 0 auto;
  padding: 32px 24px;
}

/* 过渡动画 */
.fade-enter-active,
.fade-leave-active {
  transition: opacity 0.25s ease, transform 0.25s ease;
}

.fade-enter-from,
.fade-leave-to {
  opacity: 0;
  transform: translateY(6px);
}
</style>
