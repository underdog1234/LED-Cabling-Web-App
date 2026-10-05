import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'
import { resolve } from 'node:path'

export default defineConfig({
  base: './',
  plugins: [react(), tailwindcss()],
  build: {
    rollupOptions: {
      // Two pages: the LED Cabling Planner (index.html) and the standalone
      // Test Pattern Generator (generator.html).
      input: {
        main: resolve(import.meta.dirname, 'index.html'),
        generator: resolve(import.meta.dirname, 'generator.html'),
      },
    },
  },
})
