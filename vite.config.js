import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  base: './',  // <-- මේ line එක අලුතෙන් add කරන්න
  build: {
    outDir: 'dist'
  }
})