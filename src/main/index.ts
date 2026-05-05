import { app, BrowserWindow, dialog, ipcMain, nativeImage, session, type NativeImage, type OpenDialogOptions, type SaveDialogOptions } from 'electron'
import { join } from 'node:path'
import { connectAccount, disconnectAccount, getAuthSession, resetAuthCache, resetAuthClient, switchActiveAccount } from './auth/microsoftAuth'
import {
  bindUnscopedDriveTransfersToActiveAccount,
  clearDriveIndexMemory,
  compareDriveFolders,
  copyDriveItems,
  deleteDriveItem,
  deleteDriveTransfer,
  downloadDriveItemsToDirectory,
  downloadDriveItemToPath,
  listDriveAccountUsage,
  listDriveTransferPage,
  listDriveTransfers,
  listDriveChildren,
  getDriveItemThumbnail,
  moveDriveItems,
  renameDriveItem,
  reconcileComparedDriveFolders,
  resetDriveIndex,
  resetDriveThumbnailCache,
  resetDriveTransfers,
  resumeDriveTransfers,
  searchDriveItems,
  setGraphActivityListener,
  startDriveTransferRetryScheduler,
  stopDriveTransfer,
  transferDriveItemsBetweenAccounts,
  uploadLocalFilesToDrive,
  warmDriveIndex,
  wakeDriveTransferQueue
} from './graph/oneDriveGraph'
import { getAppEnvironment } from './platform'
import {
  getMicrosoftAuthSettings,
  getDriveSettings,
  getTransferSettings,
  resetMicrosoftAuthSettings,
  updateDriveSettings,
  updateMicrosoftAuthSettings,
  updateTransferSettings
} from './settings'
import type {
  CopyDriveItemsRequest,
  DeleteDriveItemRequest,
  DownloadDriveItemRequest,
  DownloadDriveItemsRequest,
  DriveChildrenRequest,
  DriveFolderCompareRequest,
  DriveFolderReconcileRequest,
  DriveIndexWarmRequest,
  DriveSettingsInput,
  DriveSearchRequest,
  DriveThumbnailRequest,
  DriveTransferListRequest,
  MicrosoftAuthSettingsInput,
  MoveDriveItemsRequest,
  RenameDriveItemRequest,
  SwitchAuthAccountRequest,
  TransferDriveItemsBetweenAccountsRequest,
  TransferSettingsInput,
  UploadDroppedItemsRequest,
  UploadDriveItemsRequest
} from '../shared/types'

let runtimeAppIcon: NativeImage | null | undefined

function createWindow(): void {
  const appIcon = getRuntimeAppIcon()
  const mainWindow = new BrowserWindow({
    width: 1180,
    height: 760,
    minWidth: 920,
    minHeight: 620,
    title: 'OneDrive 관리자',
    backgroundColor: '#f3f3f3',
    ...(appIcon ? { icon: appIcon } : {}),
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
  ipcMain.handle('settings:getMicrosoftAuth', () => getMicrosoftAuthSettings())
  ipcMain.handle('settings:getTransfer', () => getTransferSettings())
  ipcMain.handle('settings:getDrive', () => getDriveSettings())
  ipcMain.handle('settings:updateMicrosoftAuth', async (_event, input: MicrosoftAuthSettingsInput) => {
    const settings = await updateMicrosoftAuthSettings(input)
    resetAuthClient()
    await resetDriveIndex()
    return settings
  })
  ipcMain.handle('settings:updateTransfer', async (_event, input: TransferSettingsInput) => {
    const settings = await updateTransferSettings(input)
    wakeDriveTransferQueue()
    return settings
  })
  ipcMain.handle('settings:updateDrive', async (_event, input: DriveSettingsInput) => {
    return updateDriveSettings(input)
  })
  ipcMain.handle('settings:resetAll', async () => {
    await resetAuthCache()
    await resetDriveIndex()
    await resetDriveThumbnailCache()
    await resetDriveTransfers()
    await resetMicrosoftAuthSettings()
    await session.defaultSession.clearStorageData()
    await session.defaultSession.clearCache()
    return getAuthSession()
  })
  ipcMain.handle('auth:getSession', () => getAuthSession())
  ipcMain.handle('auth:listAccountUsage', async () => {
    const authSession = await getAuthSession()

    if (!authSession.isAuthenticated) {
      return []
    }

    return listDriveAccountUsage(authSession.accounts)
  })
  ipcMain.handle('auth:connect', async () => {
    await bindUnscopedDriveTransfersToActiveAccount()
    const authSession = await connectAccount()
    await clearDriveIndexMemory()
    return authSession
  })
  ipcMain.handle('auth:switchAccount', async (_event, request: SwitchAuthAccountRequest) => {
    await bindUnscopedDriveTransfersToActiveAccount()
    const authSession = await switchActiveAccount(request.accountId)
    await clearDriveIndexMemory()
    return authSession
  })
  ipcMain.handle('auth:disconnect', async () => {
    await bindUnscopedDriveTransfersToActiveAccount()
    const authSession = await disconnectAccount()
    await clearDriveIndexMemory()
    return authSession
  })
  ipcMain.handle('onedrive:listChildren', (_event, request: DriveChildrenRequest) => {
    return listDriveChildren(request)
  })
  ipcMain.handle('onedrive:warmIndex', (_event, request?: DriveIndexWarmRequest) => {
    return warmDriveIndex(request?.forceRefresh ?? false)
  })
  ipcMain.handle('onedrive:searchItems', (_event, request: DriveSearchRequest) => {
    return searchDriveItems(request)
  })
  ipcMain.handle('onedrive:getThumbnail', (_event, request: DriveThumbnailRequest) => {
    return getDriveItemThumbnail(request)
  })
  ipcMain.handle('onedrive:compareFolders', (_event, request: DriveFolderCompareRequest) => {
    return compareDriveFolders(request)
  })
  ipcMain.handle('onedrive:reconcileFolders', (event, request: DriveFolderReconcileRequest) => {
    return reconcileComparedDriveFolders(request, createTransferNotifier(event.sender))
  })
  ipcMain.handle('onedrive:renameItem', (_event, request: RenameDriveItemRequest) => {
    return renameDriveItem(request)
  })
  ipcMain.handle('onedrive:deleteItem', (_event, request: DeleteDriveItemRequest) => {
    return deleteDriveItem(request)
  })
  ipcMain.handle('onedrive:copyItems', (_event, request: CopyDriveItemsRequest) => {
    return copyDriveItems(request)
  })
  ipcMain.handle('onedrive:moveItems', (_event, request: MoveDriveItemsRequest) => {
    return moveDriveItems(request)
  })
  ipcMain.handle('onedrive:transferBetweenAccounts', async (event, request: TransferDriveItemsBetweenAccountsRequest) => {
    return transferDriveItemsBetweenAccounts(request, createTransferNotifier(event.sender))
  })
  ipcMain.handle('onedrive:uploadItems', async (event, request?: UploadDriveItemsRequest) => {
    const browserWindow = BrowserWindow.fromWebContents(event.sender) ?? undefined
    const notifyTransfers = createTransferNotifier(event.sender)
    const dialogOptions: OpenDialogOptions = {
      title: '업로드할 파일 선택',
      properties: ['openFile', 'multiSelections']
    }
    const selection = browserWindow ? await dialog.showOpenDialog(browserWindow, dialogOptions) : await dialog.showOpenDialog(dialogOptions)

    if (selection.canceled) {
      return {
        cancelled: true,
        items: []
      }
    }

    return {
      cancelled: false,
      items: await uploadLocalFilesToDrive(request?.parentId ?? null, selection.filePaths, notifyTransfers)
    }
  })
  ipcMain.handle('onedrive:uploadDroppedItems', async (event, request: UploadDroppedItemsRequest) => {
    const notifyTransfers = createTransferNotifier(event.sender)

    return {
      cancelled: false,
      items: await uploadLocalFilesToDrive(request.parentId ?? null, request.paths, notifyTransfers)
    }
  })
  ipcMain.handle('onedrive:downloadItem', async (event, request: DownloadDriveItemRequest) => {
    const browserWindow = BrowserWindow.fromWebContents(event.sender) ?? undefined
    const notifyTransfers = createTransferNotifier(event.sender)
    const dialogOptions: SaveDialogOptions = {
      title: '다운로드 위치 선택',
      defaultPath: request.name
    }
    const selection = browserWindow ? await dialog.showSaveDialog(browserWindow, dialogOptions) : await dialog.showSaveDialog(dialogOptions)

    if (selection.canceled || !selection.filePath) {
      return {
        cancelled: true
      }
    }

    await downloadDriveItemToPath(request.itemId, request.name, request.size, selection.filePath, notifyTransfers)

    return {
      cancelled: false,
      localPath: selection.filePath
    }
  })
  ipcMain.handle('onedrive:downloadItems', async (event, request: DownloadDriveItemsRequest) => {
    const browserWindow = BrowserWindow.fromWebContents(event.sender) ?? undefined
    const notifyTransfers = createTransferNotifier(event.sender)
    const dialogOptions: OpenDialogOptions = {
      title: '다운로드 폴더 선택',
      properties: ['openDirectory', 'createDirectory']
    }
    const selection = browserWindow ? await dialog.showOpenDialog(browserWindow, dialogOptions) : await dialog.showOpenDialog(dialogOptions)

    if (selection.canceled || selection.filePaths.length === 0) {
      return {
        cancelled: true
      }
    }

    const [directoryPath] = selection.filePaths

    const result = await downloadDriveItemsToDirectory(request.items, directoryPath, notifyTransfers)

    return {
      cancelled: false,
      directoryPath,
      ...result
    }
  })
  ipcMain.handle('transfers:list', () => {
    return listDriveTransfers()
  })
  ipcMain.handle('transfers:listPage', (_event, request?: DriveTransferListRequest) => {
    return listDriveTransferPage(request)
  })
  ipcMain.handle('transfers:resume', (event, taskId?: string) => {
    return resumeDriveTransfers(taskId, createTransferNotifier(event.sender))
  })
  ipcMain.handle('transfers:stop', (event, taskId: string) => {
    return stopDriveTransfer(taskId, createTransferNotifier(event.sender))
  })
  ipcMain.handle('transfers:delete', (event, taskId?: string) => {
    return deleteDriveTransfer(taskId, createTransferNotifier(event.sender))
  })

  void app.whenReady().then(() => {
    setGraphActivityListener(broadcastGraphActivity)
    applyRuntimeAppIcon()
    createWindow()
    startDriveTransferRetryScheduler(broadcastTransferUpdates)

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

function createTransferNotifier(sender: Electron.WebContents) {
  return (tasks: unknown): void => {
    if (!sender.isDestroyed()) {
      sender.send('transfers:updated', tasks)
    }
  }
}

function broadcastTransferUpdates(tasks: unknown): void {
  for (const window of BrowserWindow.getAllWindows()) {
    if (!window.webContents.isDestroyed()) {
      window.webContents.send('transfers:updated', tasks)
    }
  }
}

function broadcastGraphActivity(event: unknown): void {
  for (const window of BrowserWindow.getAllWindows()) {
    if (!window.webContents.isDestroyed()) {
      window.webContents.send('graph:activity', event)
    }
  }
}

function applyRuntimeAppIcon(): void {
  const appIcon = getRuntimeAppIcon()

  if (appIcon && process.platform === 'darwin' && app.dock) {
    app.dock.setIcon(appIcon)
  }
}

function getRuntimeAppIcon(): NativeImage | undefined {
  if (runtimeAppIcon !== undefined) {
    return runtimeAppIcon ?? undefined
  }

  const appIcon = nativeImage.createFromPath(join(__dirname, '../../build/icon.png'))
  runtimeAppIcon = appIcon.isEmpty() ? null : appIcon

  return runtimeAppIcon ?? undefined
}
