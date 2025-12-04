// vite.config.ts
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: './', // 关键：设置为相对路径，这样生成的 index.html 引用资源时会用 ./assets
})
