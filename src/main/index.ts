import { app, BrowserWindow, ipcMain } from 'electron'
import { join } from 'node:path'
import { getAppEnvironment, revealPathInFileManager } from './platform'

function createWindow(): void {
  const mainWindow = new BrowserWindow({
    width: 1180,
    height: 760,
    minWidth: 920,
    minHeight: 620,
    title: 'OneDrive 관리자',
    backgroundColor: '#f7f5ef',
    webPreferences: {
      preload: join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true
    }
  })

  if (process.env.ELECTRON_RENDERER_URL) {
    void mainWindow.loadURL(process.env.ELECTRON_RENDERER_URL)
  } else {
    void mainWindow.loadFile(join(__dirname, '../renderer/index.html'))
  }
}

const hasSingleInstanceLock = app.requestSingleInstanceLock()

if (!hasSingleInstanceLock) {
  app.quit()
} else {
  app.on('second-instance', () => {
    const [window] = BrowserWindow.getAllWindows()

    if (!window) {
      createWindow()
      return
    }

    if (window.isMinimized()) {
      window.restore()
    }

    window.focus()
  })

  ipcMain.handle('app:getEnvironment', () => getAppEnvironment())
  ipcMain.handle('onedrive:revealPath', (_event, targetPath: unknown) => {
    return revealPathInFileManager(targetPath)
  })

  void app.whenReady().then(() => {
    createWindow()

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createWindow()
      }
    })
  })
}

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})
