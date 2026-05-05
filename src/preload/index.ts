import { contextBridge, ipcRenderer, webUtils } from 'electron'
import type { IpcRendererEvent } from 'electron'
import type {
  AppEnvironment,
  AuthSession,
  CloudDriveItem,
  CopyDriveItemsRequest,
  CopyDriveItemsResult,
  DeleteDriveItemRequest,
  DownloadDriveItemRequest,
  DownloadDriveItemResult,
  DownloadDriveItemsRequest,
  DownloadDriveItemsResult,
  DriveAccountUsage,
  DriveTransferListRequest,
  DriveTransferListResult,
  DriveTransferTask,
  DriveChildrenRequest,
  DriveChildrenResult,
  DriveFolderCompareRequest,
  DriveFolderCompareResult,
  DriveFolderReconcileRequest,
  DriveFolderReconcileResult,
  DriveIndexStatus,
  DriveIndexWarmRequest,
  DriveSettings,
  DriveSettingsInput,
  DriveSearchRequest,
  DriveSearchResult,
  DriveThumbnailRequest,
  DriveThumbnailResult,
  GraphActivityEvent,
  MicrosoftAuthSettings,
  MicrosoftAuthSettingsInput,
  MoveDriveItemsRequest,
  RenameDriveItemRequest,
  SwitchAuthAccountRequest,
  TransferDriveItemsBetweenAccountsRequest,
  TransferDriveItemsBetweenAccountsResult,
  TransferSettings,
  TransferSettingsInput,
  UploadDroppedItemsRequest,
  UploadDriveItemsRequest,
  UploadDriveItemsResult
} from '../shared/types'

const api = {
  getEnvironment: (): Promise<AppEnvironment> => ipcRenderer.invoke('app:getEnvironment'),
  getMicrosoftAuthSettings: (): Promise<MicrosoftAuthSettings> => ipcRenderer.invoke('settings:getMicrosoftAuth'),
  updateMicrosoftAuthSettings: (input: MicrosoftAuthSettingsInput): Promise<MicrosoftAuthSettings> =>
    ipcRenderer.invoke('settings:updateMicrosoftAuth', input),
  getTransferSettings: (): Promise<TransferSettings> => ipcRenderer.invoke('settings:getTransfer'),
  updateTransferSettings: (input: TransferSettingsInput): Promise<TransferSettings> => ipcRenderer.invoke('settings:updateTransfer', input),
  getDriveSettings: (): Promise<DriveSettings> => ipcRenderer.invoke('settings:getDrive'),
  updateDriveSettings: (input: DriveSettingsInput): Promise<DriveSettings> => ipcRenderer.invoke('settings:updateDrive', input),
  getAuthSession: (): Promise<AuthSession> => ipcRenderer.invoke('auth:getSession'),
  listAccountUsage: (): Promise<DriveAccountUsage[]> => ipcRenderer.invoke('auth:listAccountUsage'),
  connectAccount: (): Promise<AuthSession> => ipcRenderer.invoke('auth:connect'),
  switchAccount: (request: SwitchAuthAccountRequest): Promise<AuthSession> => ipcRenderer.invoke('auth:switchAccount', request),
  disconnectAccount: (): Promise<AuthSession> => ipcRenderer.invoke('auth:disconnect'),
  resetAllSettings: (): Promise<AuthSession> => ipcRenderer.invoke('settings:resetAll'),
  listDriveChildren: (request: DriveChildrenRequest): Promise<DriveChildrenResult> =>
    ipcRenderer.invoke('onedrive:listChildren', request),
  warmDriveIndex: (request?: DriveIndexWarmRequest): Promise<DriveIndexStatus> => ipcRenderer.invoke('onedrive:warmIndex', request),
  searchDriveItems: (request: DriveSearchRequest): Promise<DriveSearchResult> => ipcRenderer.invoke('onedrive:searchItems', request),
  getDriveThumbnail: (request: DriveThumbnailRequest): Promise<DriveThumbnailResult> =>
    ipcRenderer.invoke('onedrive:getThumbnail', request),
  compareDriveFolders: (request: DriveFolderCompareRequest): Promise<DriveFolderCompareResult> =>
    ipcRenderer.invoke('onedrive:compareFolders', request),
  reconcileDriveFolders: (request: DriveFolderReconcileRequest): Promise<DriveFolderReconcileResult> =>
    ipcRenderer.invoke('onedrive:reconcileFolders', request),
  renameDriveItem: (request: RenameDriveItemRequest): Promise<CloudDriveItem> => ipcRenderer.invoke('onedrive:renameItem', request),
  deleteDriveItem: (request: DeleteDriveItemRequest): Promise<void> => ipcRenderer.invoke('onedrive:deleteItem', request),
  copyDriveItems: (request: CopyDriveItemsRequest): Promise<CopyDriveItemsResult> => ipcRenderer.invoke('onedrive:copyItems', request),
  moveDriveItems: (request: MoveDriveItemsRequest): Promise<CloudDriveItem[]> => ipcRenderer.invoke('onedrive:moveItems', request),
  transferDriveItemsBetweenAccounts: (
    request: TransferDriveItemsBetweenAccountsRequest
  ): Promise<TransferDriveItemsBetweenAccountsResult> => ipcRenderer.invoke('onedrive:transferBetweenAccounts', request),
  uploadDriveItems: (request?: UploadDriveItemsRequest): Promise<UploadDriveItemsResult> =>
    ipcRenderer.invoke('onedrive:uploadItems', request),
  uploadDroppedItems: (request: UploadDroppedItemsRequest): Promise<UploadDriveItemsResult> =>
    ipcRenderer.invoke('onedrive:uploadDroppedItems', request),
  getDroppedFilePaths: (files: File[]): string[] => files.map((file) => webUtils.getPathForFile(file)).filter(Boolean),
  downloadDriveItem: (request: DownloadDriveItemRequest): Promise<DownloadDriveItemResult> =>
    ipcRenderer.invoke('onedrive:downloadItem', request),
  downloadDriveItems: (request: DownloadDriveItemsRequest): Promise<DownloadDriveItemsResult> =>
    ipcRenderer.invoke('onedrive:downloadItems', request),
  listTransfers: (): Promise<DriveTransferTask[]> => ipcRenderer.invoke('transfers:list'),
  listTransferPage: (request?: DriveTransferListRequest): Promise<DriveTransferListResult> => ipcRenderer.invoke('transfers:listPage', request),
  resumeTransfers: (taskId?: string): Promise<DriveTransferTask[]> => ipcRenderer.invoke('transfers:resume', taskId),
  stopTransfer: (taskId: string): Promise<DriveTransferTask[]> => ipcRenderer.invoke('transfers:stop', taskId),
  deleteTransfer: (taskId?: string): Promise<DriveTransferTask[]> => ipcRenderer.invoke('transfers:delete', taskId),
  onTransfersUpdated: (callback: (tasks: DriveTransferTask[]) => void): (() => void) => {
    const listener = (_event: IpcRendererEvent, tasks: DriveTransferTask[]): void => callback(tasks)

    ipcRenderer.on('transfers:updated', listener)
    return () => ipcRenderer.removeListener('transfers:updated', listener)
  },
  onGraphActivity: (callback: (event: GraphActivityEvent) => void): (() => void) => {
    const listener = (_event: IpcRendererEvent, graphEvent: GraphActivityEvent): void => callback(graphEvent)

    ipcRenderer.on('graph:activity', listener)
    return () => ipcRenderer.removeListener('graph:activity', listener)
  }
}

contextBridge.exposeInMainWorld('oneDriveManager', api)
