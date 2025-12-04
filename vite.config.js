// vite.config.js
import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  // 关键点：设置为你的仓库名，或者 './' (相对路径)
  base: './', 
})