import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// One page: the LED Cabling Planner (index.html). The Test Pattern Generator
// used to be built here too, from generator.html; it is its own app in its own
// repo now (github.com/underdog1234/Test-Pattern-Generator).
export default defineConfig({
  base: './',
  plugins: [react(), tailwindcss()],
})
