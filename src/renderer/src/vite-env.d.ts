/// <reference types="vite/client" />

import type { AppEnvironment } from '@shared/types'
import type {
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
  DriveFileIconRequest,
  DriveFileIconResult,
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
} from '@shared/types'

declare global {
  interface Window {
    oneDriveManager: {
      getEnvironment: () => Promise<AppEnvironment>
      getMicrosoftAuthSettings: () => Promise<MicrosoftAuthSettings>
      updateMicrosoftAuthSettings: (input: MicrosoftAuthSettingsInput) => Promise<MicrosoftAuthSettings>
      getTransferSettings: () => Promise<TransferSettings>
      updateTransferSettings: (input: TransferSettingsInput) => Promise<TransferSettings>
      getDriveSettings: () => Promise<DriveSettings>
      updateDriveSettings: (input: DriveSettingsInput) => Promise<DriveSettings>
      getAuthSession: () => Promise<AuthSession>
      listAccountUsage: () => Promise<DriveAccountUsage[]>
      connectAccount: () => Promise<AuthSession>
      switchAccount: (request: SwitchAuthAccountRequest) => Promise<AuthSession>
      disconnectAccount: () => Promise<AuthSession>
      resetAllSettings: () => Promise<AuthSession>
      listDriveChildren: (request: DriveChildrenRequest) => Promise<DriveChildrenResult>
      warmDriveIndex: (request?: DriveIndexWarmRequest) => Promise<DriveIndexStatus>
      searchDriveItems: (request: DriveSearchRequest) => Promise<DriveSearchResult>
      getDriveThumbnail: (request: DriveThumbnailRequest) => Promise<DriveThumbnailResult>
      getDriveFileIcon: (request: DriveFileIconRequest) => Promise<DriveFileIconResult>
      compareDriveFolders: (request: DriveFolderCompareRequest) => Promise<DriveFolderCompareResult>
      reconcileDriveFolders: (request: DriveFolderReconcileRequest) => Promise<DriveFolderReconcileResult>
      renameDriveItem: (request: RenameDriveItemRequest) => Promise<CloudDriveItem>
      deleteDriveItem: (request: DeleteDriveItemRequest) => Promise<void>
      copyDriveItems: (request: CopyDriveItemsRequest) => Promise<CopyDriveItemsResult>
      moveDriveItems: (request: MoveDriveItemsRequest) => Promise<CloudDriveItem[]>
      transferDriveItemsBetweenAccounts: (
        request: TransferDriveItemsBetweenAccountsRequest
      ) => Promise<TransferDriveItemsBetweenAccountsResult>
      uploadDriveItems: (request?: UploadDriveItemsRequest) => Promise<UploadDriveItemsResult>
      uploadDroppedItems: (request: UploadDroppedItemsRequest) => Promise<UploadDriveItemsResult>
      getDroppedFilePaths: (files: File[]) => string[]
      downloadDriveItem: (request: DownloadDriveItemRequest) => Promise<DownloadDriveItemResult>
      downloadDriveItems: (request: DownloadDriveItemsRequest) => Promise<DownloadDriveItemsResult>
      listTransfers: () => Promise<DriveTransferTask[]>
      listTransferPage: (request?: DriveTransferListRequest) => Promise<DriveTransferListResult>
      resumeTransfers: (taskId?: string) => Promise<DriveTransferTask[]>
      stopTransfer: (taskId: string) => Promise<DriveTransferTask[]>
      deleteTransfer: (taskId?: string) => Promise<DriveTransferTask[]>
      onTransfersUpdated: (callback: (tasks: DriveTransferTask[]) => void) => () => void
      onGraphActivity: (callback: (event: GraphActivityEvent) => void) => () => void
    }
  }
}

export {}
