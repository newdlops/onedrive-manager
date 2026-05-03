export type PlatformName = 'aix' | 'darwin' | 'freebsd' | 'linux' | 'openbsd' | 'sunos' | 'win32'

export type PlatformInfo = {
  name: PlatformName
  arch: string
  release: string
  homeDirectory: string
  appDataPath: string
}

export type AppEnvironment = {
  platform: PlatformInfo
}

export type AuthAccount = {
  homeAccountId: string
  username: string
  name: string
}

export type MicrosoftAuthSettingsSource = 'app' | 'environment' | 'missing'

export type MicrosoftAuthSettings = {
  clientId: string
  tenantId: string
  isConfigured: boolean
  source: MicrosoftAuthSettingsSource
}

export type MicrosoftAuthSettingsInput = {
  clientId: string
  tenantId?: string
}

export type TransferSettings = {
  maxConcurrentTransfers: number
  minConcurrentTransfers: number
  maxAllowedConcurrentTransfers: number
}

export type TransferSettingsInput = {
  maxConcurrentTransfers: number
}

export type AuthSession = {
  isConfigured: boolean
  isAuthenticated: boolean
  account?: AuthAccount
  accounts: AuthAccount[]
  activeAccountId?: string
  settings: MicrosoftAuthSettings
  scopes: string[]
  message?: string
}

export type DriveAccountUsage = {
  accountId: string
  used: number
  total?: number
  remaining?: number
  state?: string
  isUnavailable?: boolean
  error?: string
}

export type SwitchAuthAccountRequest = {
  accountId: string
}

export type CloudDriveItemType = 'file' | 'folder' | 'package'

export type CloudDriveItem = {
  id: string
  name: string
  type: CloudDriveItemType
  size: number
  lastModifiedDateTime?: string
  webUrl?: string
  parentId?: string
  childCount?: number
  mimeType?: string
}

export type DriveFolderRef = {
  id: string | null
  name: string
}

export type DriveChildrenRequest = {
  folderId?: string | null
  nextLink?: string | null
  forceRefresh?: boolean
}

export type DriveChildrenResult = {
  folderId: string | null
  items: CloudDriveItem[]
  nextLink?: string
}

export type DriveIndexWarmRequest = {
  forceRefresh?: boolean
}

export type DriveIndexStatus = {
  isReady: boolean
  isFresh: boolean
  isSyncing: boolean
  itemCount: number
  folderCount: number
  fileCount: number
  packageCount: number
  syncedAt?: string
}

export type DriveFolderCompareEndpoint = {
  accountId: string
  folderId: string | null
  folderName: string
}

export type DriveFolderCompareRequest = {
  source: DriveFolderCompareEndpoint
  target: DriveFolderCompareEndpoint
}

export type DriveFolderCompareDifferenceKind = 'only-in-source' | 'only-in-target' | 'different'

export type DriveFolderCompareDifferenceReason = 'type' | 'content'

export type DriveFolderCompareItem = {
  id: string
  path: string
  name: string
  type: CloudDriveItemType
  size: number
  lastModifiedDateTime?: string
}

export type DriveFolderCompareDifference = {
  kind: DriveFolderCompareDifferenceKind
  path: string
  source?: DriveFolderCompareItem
  target?: DriveFolderCompareItem
  reasons: DriveFolderCompareDifferenceReason[]
}

export type DriveFolderCompareResult = {
  source: DriveFolderCompareEndpoint
  target: DriveFolderCompareEndpoint
  comparedAt: string
  sourceItemCount: number
  targetItemCount: number
  differenceCount: number
  onlyInSourceCount: number
  onlyInTargetCount: number
  changedCount: number
  differences: DriveFolderCompareDifference[]
}

export type DriveFolderReconcilePriority = 'source' | 'target'

export type DriveFolderReconcileRequest = {
  compare: DriveFolderCompareResult
  priority: DriveFolderReconcilePriority
}

export type DriveFolderReconcileResult = {
  queuedCount: number
  sourceToTargetCount: number
  targetToSourceCount: number
  createdFolderCount: number
  skippedCount: number
}

export type RenameDriveItemRequest = {
  itemId: string
  name: string
}

export type DeleteDriveItemRequest = {
  itemId: string
}

export type CopyDriveItemsRequest = {
  itemIds: string[]
  parentId?: string | null
}

export type CopyDriveItemsResult = {
  items: CloudDriveItem[]
  hasPendingOperations: boolean
}

export type MoveDriveItemsRequest = {
  itemIds: string[]
  parentId?: string | null
}

export type DownloadDriveItemRequest = {
  itemId: string
  name: string
  type?: CloudDriveItemType
  size?: number
}

export type DownloadDriveItemResult = {
  cancelled: boolean
  localPath?: string
}

export type DownloadDriveItemsRequest = {
  items: DownloadDriveItemRequest[]
}

export type DownloadDriveItemsResult = {
  cancelled: boolean
  directoryPath?: string
  queuedCount?: number
  createdFolderCount?: number
  skippedCount?: number
}

export type UploadDriveItemsRequest = {
  parentId?: string | null
}

export type UploadDroppedItemsRequest = {
  parentId?: string | null
  paths: string[]
}

export type UploadDriveItemsResult = {
  cancelled: boolean
  items: CloudDriveItem[]
}

export type DriveTransferKind = 'upload' | 'download' | 'account-transfer'

export type DriveTransferStatus = 'queued' | 'running' | 'paused' | 'retrying' | 'completed' | 'failed'

export type DriveTransferTask = {
  id: string
  kind: DriveTransferKind
  status: DriveTransferStatus
  name: string
  transferredBytes: number
  totalBytes: number
  bytesPerSecond?: number
  attemptCount?: number
  nextRetryAt?: string
  lastError?: string
  failureStage?: string
  createdAt: string
  updatedAt: string
  localPath?: string
  message?: string
}

export type DriveTransferSummary = {
  totalCount: number
  activeCount: number
  completedCount: number
  queuedCount: number
  runningCount: number
  pausedCount: number
  retryingCount: number
  failedCount: number
  totalBytes: number
  transferredBytes: number
  bytesPerSecond: number
}

export type DriveTransferListRequest = {
  offset?: number
  limit?: number
}

export type DriveTransferListResult = {
  tasks: DriveTransferTask[]
  summary: DriveTransferSummary
  offset: number
  limit: number
  totalTaskCount: number
}

export type TransferDriveItemRef = {
  itemId: string
  name: string
  type: CloudDriveItemType
  size?: number
}

export type TransferDriveItemsBetweenAccountsRequest = {
  sourceAccountId: string
  targetAccountId: string
  targetParentId?: string | null
  items: TransferDriveItemRef[]
  deleteSourceOnComplete?: boolean
}

export type TransferDriveItemsBetweenAccountsResult = {
  queuedCount: number
}
