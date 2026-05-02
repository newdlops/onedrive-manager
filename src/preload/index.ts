import { contextBridge, ipcRenderer } from 'electron'
import type { AppEnvironment, RevealPathResult } from '../shared/types'

const api = {
  getEnvironment: (): Promise<AppEnvironment> => ipcRenderer.invoke('app:getEnvironment'),
  revealPath: (targetPath: string): Promise<RevealPathResult> => ipcRenderer.invoke('onedrive:revealPath', targetPath)
}

contextBridge.exposeInMainWorld('oneDriveManager', api)

