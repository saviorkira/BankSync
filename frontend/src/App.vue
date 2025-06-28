<template>
  <div class="p-4 max-w-4xl mx-auto">
    <h1 class="text-2xl font-bold mb-4">银行流水回单导出</h1>
    <div class="mb-4">
      <label class="block text-sm font-medium text-gray-700">选择银行</label>
      <select v-model="bank" @change="onBankSelect" class="mt-1 block w-48 rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring focus:ring-blue-500 focus:ring-opacity-50">
        <option :value="null" disabled>请选择银行</option>
        <option value="Ningbo Bank">宁波银行</option>
      </select>
    </div>
    <div class="mb-4 flex space-x-4">
      <button @click="importExcel" class="px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600 transition">导入 Excel</button>
      <button @click="selectBasePath" class="px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600 transition">选择下载路径</button>
    </div>
    <div class="mb-4 text-sm text-gray-700">下载路径: {{ basePath }}</div>
    <div class="mb-4 flex space-x-4">
      <div>
        <label class="block text-sm font-medium text-gray-700">开始日期</label>
        <input v-model="startDate" type="text" placeholder="YYYY-MM-DD" class="mt-1 block w-48 rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring focus:ring-blue-500 focus:ring-opacity-50">
      </div>
      <div>
        <label class="block text-sm font-medium text-gray-700">结束日期</label>
        <input v-model="endDate" type="text" placeholder="YYYY-MM-DD" class="mt-1 block w-48 rounded-md border-gray-300 shadow-sm focus:border-blue-500 focus:ring focus:ring-blue-500 focus:ring-opacity-50">
      </div>
    </div>
    <button @click="runExport" :disabled="isRunning" class="px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600 transition" :class="{ 'opacity-50 cursor-not-allowed': isRunning }">
      {{ runButtonText }}
    </button>
    <div class="mt-4">
      <h2 class="text-lg font-semibold">导入数据</h2>
      <table class="min-w-full border-collapse border border-gray-300">
        <thead>
          <tr>
            <th class="border border-gray-300 px-4 py-2">项目名称</th>
            <th class="border border-gray-300 px-4 py-2">银行账号</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="(item, index) in excelData" :key="index">
            <td class="border border-gray-300 px-4 py-2">{{ item[0] }}</td>
            <td class="border border-gray-300 px-4 py-2">{{ item[1] }}</td>
          </tr>
        </tbody>
      </table>
    </div>
    <div class="mt-4">
      <h2 class="text-lg font-semibold">日志</h2>
      <textarea v-model="log" readonly class="w-full h-40 p-2 border border-gray-300 rounded-md font-mono text-sm"></textarea>
    </div>
  </div>
</template>

<script>
export default {
  data() {
    return {
      bank: null,
      excelData: [],
      basePath: 'D:\\Desktop',
      startDate: '2025-03-01',
      endDate: '2025-03-31',
      log: '',
      isRunning: false,
      runButtonText: '运行'
    }
  },
  methods: {
    async onBankSelect() {
      if (this.bank) {
        await window.pywebview.api.log(`已选择银行: ${this.bank === 'Ningbo Bank' ? '宁波银行' : '未知'}`)
      } else {
        await window.pywebview.api.log('银行选择已清空')
      }
    },
    async importExcel() {
      const result = await window.pywebview.api.import_excel()
      if (result.success) {
        this.excelData = result.data
      }
    },
    async selectBasePath() {
      const result = await window.pywebview.api.select_base_path()
      if (result.success) {
        this.basePath = result.path
      }
    },
    async runExport() {
      const result = await window.pywebview.api.run_export(this.bank, this.startDate, this.endDate)
      if (result.success) {
        this.isRunning = true
        this.runButtonText = '运行中...'
      }
    },
    updateLog(message) {
      this.log += `${message}\n`
      this.$nextTick(() => {
        const textarea = this.$refs.log
        if (textarea) {
          textarea.scrollTop = textarea.scrollHeight
        }
      })
    },
    updateRunButton(isRunning, text) {
      this.isRunning = isRunning
      this.runButtonText = text
    }
  },
  mounted() {
    window.updateLog = this.updateLog
    window.updateRunButton = this.updateRunButton
  }
}
</script>