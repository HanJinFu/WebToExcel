<script setup lang="ts">
import { ref } from 'vue'
import JSZip from 'jszip'

const fileInput = ref<HTMLInputElement | null>(null)
const processing = ref(false)
const message = ref('')
const fileName = ref('')

const triggerFileInput = () => {
  fileInput.value?.click()
}

const handleFileChange = (event: Event) => {
  const target = event.target as HTMLInputElement
  const file = target.files?.[0]
  if (!file) return

  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    message.value = '请选择 .xlsx 格式的文件'
    return
  }

  fileName.value = file.name
  processFile(file)
}

const processFile = async (file: File) => {
  processing.value = true
  message.value = '正在处理...'

  try {
    const arrayBuffer = await file.arrayBuffer()
    await processXlsxDirectly(arrayBuffer, file.name)
    message.value = '✓ 处理完成，文件已开始下载'
  } catch (error) {
    console.error('处理文件时出错:', error)
    message.value = `处理失败: ${error instanceof Error ? error.message : '未知错误'}`
  } finally {
    processing.value = false
  }
}

const processXlsxDirectly = async (arrayBuffer: ArrayBuffer, fileName: string) => {
  const zip = await JSZip.loadAsync(arrayBuffer)
  const stylesFile = zip.file('xl/styles.xml')
  if (!stylesFile) throw new Error('无法找到样式文件')

  const stylesContent = await stylesFile.async('string')
  const modifiedStyles = stylesContent.replace(/patternType="solid"/g, 'patternType="none"')
  zip.file('xl/styles.xml', modifiedStyles)

  const outputBuffer = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 9 } })
  downloadBlob(outputBuffer, fileName.replace('.xlsx', '_clean.xlsx'))
}

function downloadBlob(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = name
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}
</script>

<template>
  <div class="card">
    <div class="card-header">
      <h2 class="card-title">移除背景样式</h2>
      <p class="card-desc">上传 Excel 文件，自动清除单元格背景填充，保留全部数据</p>
    </div>

    <div class="upload-area" :class="{ dragging: false }" @click="triggerFileInput">
      <input ref="fileInput" type="file" accept=".xlsx" style="display:none" @change="handleFileChange" />
      <div class="upload-icon">↑</div>
      <p class="upload-text">{{ fileName || '点击选择或拖入 .xlsx 文件' }}</p>
      <p class="upload-hint">支持所有 Excel 2007+ 格式</p>
    </div>

    <div v-if="processing" class="status-bar loading">
      <span class="spinner"></span>
      <span>{{ message }}</span>
    </div>
    <div v-else-if="message" :class="['status-bar', message.startsWith('✓') ? 'success' : 'error']">
      {{ message }}
    </div>
  </div>
</template>

<style scoped>
.card {
  background: rgba(255, 255, 255, 0.8);
  backdrop-filter: blur(20px);
  -webkit-backdrop-filter: blur(20px);
  border-radius: 20px;
  border: 1px solid rgba(255, 255, 255, 0.6);
  box-shadow: 0 4px 24px rgba(0, 0, 0, 0.06), 0 1px 2px rgba(0, 0, 0, 0.04);
  padding: 32px;
}

.card-header {
  margin-bottom: 24px;
}

.card-title {
  font-size: 22px;
  font-weight: 600;
  color: #1d1d1f;
  letter-spacing: -0.02em;
  margin: 0 0 6px;
}

.card-desc {
  font-size: 14px;
  color: #6e6e73;
  margin: 0;
  line-height: 1.5;
}

.upload-area {
  border: 2px dashed rgba(0, 0, 0, 0.12);
  border-radius: 14px;
  padding: 48px 24px;
  text-align: center;
  cursor: pointer;
  transition: all 0.2s ease;
  background: rgba(249, 249, 249, 0.5);
}

.upload-area:hover {
  border-color: #0071e3;
  background: rgba(0, 113, 227, 0.03);
}

.upload-icon {
  font-size: 32px;
  color: #0071e3;
  margin-bottom: 12px;
}

.upload-text {
  font-size: 15px;
  color: #1d1d1f;
  font-weight: 500;
  margin: 0 0 4px;
}

.upload-hint {
  font-size: 13px;
  color: #86868b;
  margin: 0;
}

.status-bar {
  margin-top: 16px;
  padding: 12px 16px;
  border-radius: 10px;
  font-size: 14px;
  display: flex;
  align-items: center;
  gap: 8px;
}

.status-bar.loading {
  background: rgba(0, 113, 227, 0.08);
  color: #0071e3;
}

.status-bar.success {
  background: rgba(52, 199, 89, 0.1);
  color: #248a3d;
}

.status-bar.error {
  background: rgba(255, 59, 48, 0.08);
  color: #ff3b30;
}

.spinner {
  display: inline-block;
  width: 14px;
  height: 14px;
  border: 2px solid rgba(0, 113, 227, 0.3);
  border-top-color: #0071e3;
  border-radius: 50%;
  animation: spin 0.7s linear infinite;
}

@keyframes spin {
  to { transform: rotate(360deg); }
}
</style>
