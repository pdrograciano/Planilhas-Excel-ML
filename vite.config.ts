import { defineConfig } from 'vite'

export default defineConfig({
  base: './',
  optimizeDeps: { include: ['exceljs', 'react/jsx-dev-runtime'] },
})
