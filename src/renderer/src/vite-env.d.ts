/// <reference types="vite/client" />

import type { AppEnvironment, RevealPathResult } from '@shared/types'

declare global {
  interface Window {
    oneDriveManager: {
      getEnvironment: () => Promise<AppEnvironment>
      revealPath: (targetPath: string) => Promise<RevealPathResult>
    }
  }
}

export {}

