import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

export default defineConfig({
  base: '/json-toolkit/',   // ← add this line
  plugins: [react(), tailwindcss()],
})