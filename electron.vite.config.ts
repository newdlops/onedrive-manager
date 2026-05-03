import { resolve } from 'node:path'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
import react from '@vitejs/plugin-react'
import { loadEnv } from 'vite'

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), 'MAIN_VITE_')

  return {
    main: {
      define: {
        'import.meta.env.MAIN_VITE_MICROSOFT_CLIENT_ID': JSON.stringify(env.MAIN_VITE_MICROSOFT_CLIENT_ID ?? ''),
        'import.meta.env.MAIN_VITE_MICROSOFT_TENANT_ID': JSON.stringify(env.MAIN_VITE_MICROSOFT_TENANT_ID ?? 'consumers')
      },
      plugins: [externalizeDepsPlugin()]
    },
    preload: {
      plugins: [externalizeDepsPlugin()]
    },
    renderer: {
      resolve: {
        alias: {
          '@renderer': resolve('src/renderer/src'),
          '@shared': resolve('src/shared')
        }
      },
      plugins: [react()]
    }
  }
})
