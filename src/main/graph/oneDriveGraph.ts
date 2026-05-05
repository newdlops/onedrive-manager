import { app } from 'electron'
import { Buffer } from 'node:buffer'
import { createWriteStream } from 'node:fs'
import { mkdir, open, readFile, readdir, rename, rm, stat, unlink, writeFile } from 'node:fs/promises'
import { once } from 'node:events'
import { createHash, randomUUID } from 'node:crypto'
import { basename, dirname, extname, join, win32 } from 'node:path'
import { Readable } from 'node:stream'
import { getActiveAccountId, getGraphAccessToken } from '../auth/microsoftAuth'
import { getDriveSettings, getTransferSettings } from '../settings'
import {
  appendDriveDeltaStaging,
  applyDriveIndexDeltaToStore,
  clearDriveDeltaStaging,
  closeDriveIndexStores,
  createEmptyDriveIndex,
  findDriveIndexChildByName,
  getDriveIndexStoreDirectory,
  getLegacyAccountDriveIndexPath,
  listDriveIndexChildren,
  readDriveIndexFromStore,
  readLastOccurrenceDriveDeltaStaging,
  resetDriveDeltaStaging,
  searchDriveIndexItems,
  writeDriveIndexToStore,
  type DriveIndex,
  type DriveDeltaStagePayload,
  type DriveIndexDeltaChange
} from './driveIndexStore'
import type {
  AuthAccount,
  CloudDriveItem,
  CloudDriveItemType,
  CopyDriveItemsRequest,
  CopyDriveItemsResult,
  DeleteDriveItemRequest,
  DriveAccountUsage,
  DriveFolderCompareDifference,
  DriveFolderCompareDifferenceReason,
  DriveFolderCompareEndpoint,
  DriveFolderCompareItem,
  DriveFolderCompareRequest,
  DriveFolderCompareResult,
  DriveFolderReconcilePriority,
  DriveFolderReconcileRequest,
  DriveFolderReconcileResult,
  DownloadDriveItemRequest,
  DownloadDriveItemsResult,
  DriveTransferListRequest,
  DriveTransferListResult,
  DriveTransferSummary,
  DriveTransferTask,
  DriveChildrenRequest,
  DriveChildrenResult,
  DriveIndexStatus,
  DriveSearchRequest,
  DriveSearchResult,
  DriveThumbnailRequest,
  DriveThumbnailResult,
  DriveThumbnailSize,
  GraphActivityEvent,
  MoveDriveItemsRequest,
  RenameDriveItemRequest,
  TransferDriveItemRef,
  TransferDriveItemsBetweenAccountsRequest,
  TransferDriveItemsBetweenAccountsResult
} from '../../shared/types'

const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0'
const DRIVE_INDEX_LEGACY_FILE_NAME = 'drive-index.json'
const DRIVE_TRANSFERS_LEGACY_FILE_NAME = 'drive-transfers.json'
const DRIVE_TRANSFERS_DIR_NAME = 'drive-transfers'
const DRIVE_TRANSFERS_INDEX_FILE_NAME = 'index.json'
const DRIVE_TRANSFERS_SUMMARY_FILE_NAME = 'summary.json'
const DRIVE_TRANSFERS_TASKS_DIR_NAME = 'tasks'
const DRIVE_TRANSFERS_TEMP_DIR_NAME = 'temp'
const DRIVE_THUMBNAIL_CACHE_DIR_NAME = 'drive-thumbnails'
const INDEX_FRESH_MS = 60_000
const DRIVE_ITEM_SELECT = 'id,name,size,lastModifiedDateTime,webUrl,parentReference,folder,file,package,deleted,cTag,eTag'
const UPLOAD_CHUNK_SIZE_BYTES = 10 * 1024 * 1024
const TRANSFER_PROGRESS_SAVE_INTERVAL_MS = 3_000
const TRANSFER_VISIBLE_LIMIT = 100
const TRANSFER_LIST_MAX_LIMIT = 500
const TRANSFER_RETRY_BATCH_LIMIT = 64
const TRANSFER_DISPATCH_DELAY_MS = 100
const TRANSFER_RETRY_IDLE_DELAY_MS = 30_000
const TRANSFER_RETRY_READY_WAIT_MS = 5_000
const TRANSFER_RETRY_STEP_DELAY_MS = 5_000
const TRANSFER_RETRY_BASE_DELAY_MS = 5_000
const TRANSFER_RETRY_MAX_DELAY_MS = 30 * 60_000
const DEFAULT_THROTTLE_RETRY_DELAY_MS = 60_000
const TRANSFER_SLOT_MIN = 1
const UPLOAD_QUEUE_BATCH_SIZE = 500
const COPY_OPERATION_POLL_INTERVAL_MS = 1000
const COPY_OPERATION_MAX_WAIT_MS = 120_000
const DRIVE_THUMBNAIL_DEFAULT_SIZE: DriveThumbnailSize = 'c160x120_crop'
const DRIVE_THUMBNAIL_BATCH_LIMIT_MIN = 4
const DRIVE_THUMBNAIL_BATCH_LIMIT_INITIAL = 10
const DRIVE_THUMBNAIL_BATCH_LIMIT_MAX = 20
const DRIVE_THUMBNAIL_BATCH_DELAY_MS = 24
const DRIVE_THUMBNAIL_BATCH_STEP_DELAY_INITIAL_MS = 120
const DRIVE_THUMBNAIL_BATCH_STEP_DELAY_MIN_MS = 24
const DRIVE_THUMBNAIL_BATCH_STEP_DELAY_MAX_MS = 2_000
const DRIVE_THUMBNAIL_ADAPTIVE_SUCCESS_TARGET = 60
const DRIVE_THUMBNAIL_URL_CACHE_MS = 20 * 60_000
const DRIVE_THUMBNAIL_MISSING_CACHE_MS = 5 * 60_000
const DRIVE_THUMBNAIL_MAX_RETRIES = 3
const DRIVE_THUMBNAIL_RETRY_BASE_DELAY_MS = 5_000
const DRIVE_THUMBNAIL_RETRY_MAX_DELAY_MS = 5 * 60_000
const DRIVE_THUMBNAIL_DISK_CACHE_MS = 30 * 24 * 60 * 60_000
const DRIVE_THUMBNAIL_MAX_CACHE_BYTES = 5 * 1024 * 1024
const DRIVE_THUMBNAIL_DEFAULT_CONTENT_TYPE = 'image/jpeg'
const DRIVE_THUMBNAIL_WARM_BATCH_SIZE = 500
const DRIVE_THUMBNAIL_WARM_BATCH_DELAY_MS = 100
const DRIVE_THUMBNAIL_IMAGE_EXTENSIONS = new Set(['avif', 'bmp', 'gif', 'heic', 'heif', 'jpeg', 'jpg', 'png', 'tif', 'tiff', 'webp'])
const DRIVE_THUMBNAIL_VIDEO_EXTENSIONS = new Set(['avi', 'm2ts', 'm4v', 'mkv', 'mov', 'mp4', 'mpeg', 'mpg', 'webm', 'wmv'])
const DRIVE_ITEM_TYPE_ORDER: Record<CloudDriveItemType, number> = {
  folder: 0,
  package: 1,
  file: 2
}
const DRIVE_ITEM_NAME_COLLATOR = new Intl.Collator('ko-KR', {
  numeric: true,
  sensitivity: 'base'
})

type GraphConflictBehavior = 'rename' | 'replace'

const driveIndexSyncPromises = new Map<string, Promise<DriveIndex>>()
let activeDriveIndexSnapshot: DriveIndex | null = null
let activeDriveIndexSnapshotAccountId: string | null = null
let driveIndexGeneration = 0
let transferIndexCache: DriveTransferIndex | null = null
let transferRetryTimer: ReturnType<typeof setTimeout> | null = null
let transferRetryTimerDueAt = 0
let transferRetryWorkerPromise: Promise<void> | null = null
let transferRetryProgressListener: DriveTransferProgressListener | undefined
let transferScanCursor = 0
let transferThrottleUntil = 0
let nextRetryingTransferDispatchAt = 0
let hasCheckedLegacyDriveTransfers = false
let transferMetadataMutationPromise: Promise<void> = Promise.resolve()
const runningTransferIds = new Set<string>()
const transferWorkerSlotIds = new Set<string>()
const transferAbortControllers = new Map<string, AbortController>()
const transferPauseRequests = new Set<string>()
const transferDeleteRequests = new Set<string>()
const driveThumbnailCache = new Map<string, { result: DriveThumbnailResult; expiresAtMs: number }>()
const driveThumbnailPendingRequests = new Map<string, Promise<DriveThumbnailResult>>()
const driveThumbnailWarmPromises = new Map<string, Promise<void>>()
const driveThumbnailWarmSignatures = new Map<string, string>()
const driveThumbnailQueue: DriveThumbnailQueueEntry[] = []
let driveThumbnailQueueTimer: ReturnType<typeof setTimeout> | null = null
let driveThumbnailQueueFlushPromise: Promise<void> | null = null
let driveThumbnailThrottleUntil = 0
let driveThumbnailBatchLimit = DRIVE_THUMBNAIL_BATCH_LIMIT_INITIAL
let driveThumbnailBatchStepDelayMs = DRIVE_THUMBNAIL_BATCH_STEP_DELAY_INITIAL_MS
let driveThumbnailSuccessfulItemCount = 0
let graphActivityListener: ((event: GraphActivityEvent) => void) | null = null

type GraphDriveRoot = {
  id?: string
}

type GraphDrive = {
  id?: string
  quota?: {
    deleted?: number
    remaining?: number
    state?: string
    total?: number
    used?: number
  }
}

type GraphDeltaResponse = {
  value?: GraphDriveItem[]
  '@odata.nextLink'?: string
  '@odata.deltaLink'?: string
}

type GraphChildrenResponse = {
  value?: GraphDriveItem[]
  '@odata.nextLink'?: string
}

type GraphCopyOperationResponse = {
  error?: {
    code?: string
    message?: string
  }
  id?: string
  percentageComplete?: number
  resourceId?: string
  resourceLocation?: string
  status?: string
  statusDescription?: string
}

type GraphUploadSessionResponse = {
  uploadUrl?: string
  expirationDateTime?: string
  nextExpectedRanges?: string[]
}

type GraphThumbnail = {
  height?: number
  width?: number
  url?: string
}

type GraphThumbnailSet = {
  id?: string
  [key: string]: unknown
}

type GraphThumbnailResponse = {
  value?: GraphThumbnailSet[]
}

type GraphBatchThumbnailResponse = {
  responses?: GraphBatchThumbnailItemResponse[]
}

type GraphBatchThumbnailItemResponse = {
  id?: string
  status?: number
  headers?: Record<string, string>
  body?: unknown
}

type DriveThumbnailDiskCacheEntry = {
  version: 1
  itemId: string
  contentType: string
  width?: number
  height?: number
  createdAtMs: number
  expiresAtMs: number
}

type GraphDriveItem = {
  id?: string
  name?: string
  size?: number
  cTag?: string
  eTag?: string
  lastModifiedDateTime?: string
  webUrl?: string
  parentReference?: {
    id?: string
  }
  folder?: {
    childCount?: number
  }
  file?: {
    mimeType?: string
    hashes?: {
      quickXorHash?: string
      sha1Hash?: string
      sha256Hash?: string
    }
  }
  package?: {
    type?: string
  }
  deleted?: Record<string, unknown>
}

type GraphErrorResponse = {
  error?: {
    code?: string
    message?: string
  }
}

type GraphRequestInit = Omit<RequestInit, 'headers'> & {
  headers?: Record<string, string>
}

type DriveTransferState = {
  version: 1
  tasks: InternalDriveTransferTask[]
}

type DriveTransferIndex = {
  version: 1
  taskIds: string[]
}

type DriveTransferCompletedSummary = {
  completedCount: number
  completedBytes: number
}

type DriveTransferSummaryFile = {
  version: 1
  accounts: Record<string, DriveTransferCompletedSummary>
}

type InternalDriveTransferTask = DriveTransferTask & {
  accountId?: string
  parentId?: string | null
  itemId?: string
  sourceAccountId?: string
  targetAccountId?: string
  sourceItemId?: string
  targetParentId?: string | null
  sourceSize?: number
  downloadedBytes?: number
  uploadedBytes?: number
  transferGroupId?: string
  expectedGroupTaskIds?: string[]
  completedGroupTaskIds?: string[]
  cleanupOnly?: boolean
  deleteSourceOnComplete?: boolean
  sourceDeleted?: boolean
  conflictBehavior?: GraphConflictBehavior
  uploadUrl?: string
  expirationDateTime?: string
  tempPath?: string
  resultItem?: CloudDriveItem
  speedSampleBytes?: number
  speedSampleAt?: number
}

type DriveItemContentFingerprint = {
  quickXorHash?: string
  cTag?: string
  eTag?: string
}

type InternalFolderCompareItem = DriveFolderCompareItem & {
  contentFingerprint?: DriveItemContentFingerprint
}

type DriveTransferProgressListener = (tasks: DriveTransferTask[]) => void

type TransferFailureStage =
  | 'auth'
  | 'metadata'
  | 'local-file'
  | 'upload-session'
  | 'upload-chunk'
  | 'download-link'
  | 'download-stream'
  | 'finalize'
  | 'unknown'

type TransferRetryCandidateMode = 'any' | 'non-retrying' | 'retrying'

type TransferRetryCandidateOptions = {
  mode?: TransferRetryCandidateMode
  minimumRetryingWaitMs?: number
}

type LocalUploadFile = {
  parentId: string
  localPath: string
}

type AccountTransferQueueContext = {
  sourceAccountId: string
  targetAccountId: string
  sourceAccessToken: string
  targetAccessToken: string
  targetRootId: string
  targetParentId: string
  item: TransferDriveItemRef
  deleteSourceOnComplete: boolean
  conflictBehavior?: GraphConflictBehavior
  onSkippedItem?: () => void
  transferGroupId: string
  expectedGroupTaskIds: string[]
  tasks: InternalDriveTransferTask[]
  flushTasks: () => Promise<void>
}

class GraphResponseError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly location: string | null,
    readonly retryAfterMs?: number
  ) {
    super(message)
  }
}

class TransferPausedError extends Error {
  constructor() {
    super('전송이 중지되었습니다.')
  }
}

type InternalDriveThumbnailRequest = {
  accountId: string | null
  itemId: string
  cacheKey: string
  priority: 'normal' | 'high'
  size: DriveThumbnailSize
  cacheOnly: boolean
}

type DriveThumbnailQueueEntry = {
  request: InternalDriveThumbnailRequest
  cacheKey: string
  attemptCount: number
  resolve: (result: DriveThumbnailResult) => void
  reject: (error: unknown) => void
}

type DriveThumbnailBatchOutcome =
  | {
      status: 'fulfilled'
      result: DriveThumbnailResult
    }
  | {
      status: 'rejected'
      reason: unknown
    }

type GraphActivityInput = Omit<GraphActivityEvent, 'id' | 'at'>

export function setGraphActivityListener(listener: ((event: GraphActivityEvent) => void) | null): void {
  graphActivityListener = listener
}

function emitGraphActivity(input: GraphActivityInput): void {
  graphActivityListener?.({
    id: randomUUID(),
    at: new Date().toISOString(),
    ...input
  })
}

export async function listDriveChildren(request: DriveChildrenRequest): Promise<DriveChildrenResult> {
  const accountId = await getActiveAccountId()
  const index = await ensureDriveIndexForListing(request.forceRefresh ?? false)
  const parentId = request.folderId ?? index.rootItemId

  if (!parentId) {
    throw new Error('OneDrive 루트 폴더 정보를 확인하지 못했습니다.')
  }

  const currentIndex =
    request.forceRefresh || !isFolderLocallyNavigable(index, parentId) ? await fetchAndMergeFolderChildren(index, parentId) : index
  const storedChildren = await listDriveIndexChildren(accountId, parentId)

  return {
    folderId: request.folderId ?? null,
    items: sortDriveItems(storedChildren.length > 0 ? storedChildren : Object.values(currentIndex.items).filter((item) => item.parentId === parentId))
  }
}

export async function warmDriveIndex(forceSync = false): Promise<DriveIndexStatus> {
  const accountId = await getActiveAccountId()
  const currentIndex = await getCurrentDriveIndex(accountId)
  const driveSettings = await getDriveSettings()

  if ((forceSync || (driveSettings.indexMode === 'automatic' && !isIndexReadyAndFresh(currentIndex))) && !isDriveIndexSyncing(accountId)) {
    emitGraphActivity({
      level: 'info',
      scope: 'index',
      title: '탐색 인덱스 동기화 시작',
      message: forceSync ? '수동 요청으로 OneDrive 전체 인덱스를 갱신합니다.' : 'OneDrive 탐색 인덱스를 백그라운드로 갱신합니다.'
    })
    startDriveIndexSync(currentIndex, accountId)
  }

  const statusIndex = activeDriveIndexSnapshotAccountId === accountId ? activeDriveIndexSnapshot ?? currentIndex : currentIndex

  if (isIndexUsable(statusIndex) && !isDriveIndexSyncing(accountId)) {
    startDriveThumbnailCacheWarmup(statusIndex, accountId)
  }

  return createDriveIndexStatus(statusIndex, accountId)
}

export async function searchDriveItems(request: DriveSearchRequest): Promise<DriveSearchResult> {
  const accountId = await getActiveAccountId()
  await ensureDriveIndexForListing(false)

  return {
    items: await searchDriveIndexItems(accountId, request?.query ?? '', request?.limit ?? 100)
  }
}

export async function getDriveItemThumbnail(request: DriveThumbnailRequest): Promise<DriveThumbnailResult> {
  const normalizedRequest = await normalizeDriveThumbnailRequest(request)
  const cacheKey = createDriveThumbnailCacheKey(normalizedRequest)

  if (normalizedRequest.cacheOnly) {
    if (await hasFreshDriveThumbnailDiskCache(cacheKey, normalizedRequest.itemId)) {
      return {
        itemId: normalizedRequest.itemId,
        status: 'ready'
      }
    }
  } else {
    const cached = driveThumbnailCache.get(cacheKey)

    if (cached && cached.expiresAtMs > Date.now()) {
      return cached.result
    }

    const diskCached = await readCachedDriveThumbnailResult(cacheKey, normalizedRequest.itemId)

    if (diskCached) {
      cacheDriveThumbnailResult(cacheKey, diskCached)
      return diskCached
    }
  }

  const pendingKey = createDriveThumbnailPendingKey(normalizedRequest, cacheKey)
  const pendingRequest = driveThumbnailPendingRequests.get(pendingKey)

  if (pendingRequest) {
    return pendingRequest
  }

  const requestPromise = enqueueDriveThumbnailRequest(normalizedRequest, cacheKey)
  driveThumbnailPendingRequests.set(pendingKey, requestPromise)

  try {
    return await requestPromise
  } finally {
    driveThumbnailPendingRequests.delete(pendingKey)
  }
}

export async function compareDriveFolders(request: DriveFolderCompareRequest): Promise<DriveFolderCompareResult> {
  const source = normalizeFolderCompareEndpoint(request.source, '기준 폴더 정보를 확인하지 못했습니다.')
  const target = normalizeFolderCompareEndpoint(request.target, '비교 대상 폴더 정보를 확인하지 못했습니다.')
  const [sourceItems, targetItems] = await Promise.all([collectFolderCompareItems(source), collectFolderCompareItems(target)])
  const differences: DriveFolderCompareDifference[] = []
  const allKeys = [...new Set([...sourceItems.keys(), ...targetItems.keys()])].sort((left, right) => DRIVE_ITEM_NAME_COLLATOR.compare(left, right))
  let onlyInSourceCount = 0
  let onlyInTargetCount = 0
  let changedCount = 0

  for (const key of allKeys) {
    const sourceItem = sourceItems.get(key)
    const targetItem = targetItems.get(key)

    if (sourceItem && !targetItem) {
      onlyInSourceCount += 1
      differences.push({
        kind: 'only-in-source',
        path: sourceItem.path,
        source: toPublicFolderCompareItem(sourceItem),
        reasons: []
      })
      continue
    }

    if (!sourceItem && targetItem) {
      onlyInTargetCount += 1
      differences.push({
        kind: 'only-in-target',
        path: targetItem.path,
        target: toPublicFolderCompareItem(targetItem),
        reasons: []
      })
      continue
    }

    if (!sourceItem || !targetItem) {
      continue
    }

    const reasons = getFolderCompareDifferenceReasons(sourceItem, targetItem)

    if (reasons.length > 0) {
      changedCount += 1
      differences.push({
        kind: 'different',
        path: sourceItem.path,
        source: toPublicFolderCompareItem(sourceItem),
        target: toPublicFolderCompareItem(targetItem),
        reasons
      })
    }
  }

  return {
    source,
    target,
    comparedAt: new Date().toISOString(),
    sourceItemCount: sourceItems.size,
    targetItemCount: targetItems.size,
    differenceCount: differences.length,
    onlyInSourceCount,
    onlyInTargetCount,
    changedCount,
    differences
  }
}

function normalizeFolderCompareEndpoint(endpoint: DriveFolderCompareEndpoint, errorMessage: string): DriveFolderCompareEndpoint {
  const accountId = validateAccountId(endpoint.accountId, errorMessage)
  const folderId = endpoint.folderId == null ? null : validateDriveItemId(endpoint.folderId)
  const folderName = endpoint.folderName.trim() || 'OneDrive'

  return {
    accountId,
    folderId,
    folderName
  }
}

async function collectFolderCompareItems(endpoint: DriveFolderCompareEndpoint): Promise<Map<string, InternalFolderCompareItem>> {
  const accessToken = await getGraphAccessToken(endpoint.accountId)
  const rootItemId = await getRootItemId(accessToken)
  const rootFolderId = endpoint.folderId ?? rootItemId
  const items = new Map<string, InternalFolderCompareItem>()
  const foldersToVisit: Array<{ folderId: string; relativePath: string }> = [{ folderId: rootFolderId, relativePath: '' }]
  const visitedFolderIds = new Set<string>()

  while (foldersToVisit.length > 0) {
    const currentFolder = foldersToVisit.shift()

    if (!currentFolder || visitedFolderIds.has(currentFolder.folderId)) {
      continue
    }

    visitedFolderIds.add(currentFolder.folderId)
    let url = createChildrenUrlByRootId(rootItemId, currentFolder.folderId)

    while (true) {
      const response = await graphGet<GraphChildrenResponse>(url, accessToken)

      for (const child of response.value ?? []) {
        if (!child.id) {
          continue
        }

        const item = mapDriveItem(child)
        const relativePath = joinComparePath(currentFolder.relativePath, item.name)
        const compareItem = createFolderCompareItem(item, child, relativePath)

        items.set(normalizeComparePath(relativePath), compareItem)

        if (item.type === 'folder') {
          foldersToVisit.push({
            folderId: child.id,
            relativePath
          })
        }
      }

      if (!response['@odata.nextLink']) {
        break
      }

      url = parseGraphUrl(response['@odata.nextLink'], 'OneDrive 폴더 다음 페이지 주소가 올바르지 않습니다.')
    }
  }

  return items
}

function createFolderCompareItem(item: CloudDriveItem, graphItem: GraphDriveItem, path: string): InternalFolderCompareItem {
  return {
    id: item.id,
    path,
    name: item.name,
    type: item.type,
    size: item.size,
    lastModifiedDateTime: item.lastModifiedDateTime,
    contentFingerprint: getGraphItemContentFingerprint(graphItem)
  }
}

function getGraphItemContentFingerprint(item: GraphDriveItem): DriveItemContentFingerprint | undefined {
  const quickXorHash = normalizeContentHash(item.file?.hashes?.quickXorHash)
  const cTag = normalizeContentHash(item.cTag)
  const eTag = normalizeContentHash(item.eTag)

  if (!quickXorHash && !cTag && !eTag) {
    return undefined
  }

  return {
    quickXorHash,
    cTag,
    eTag
  }
}

function normalizeContentHash(hash: string | undefined): string | undefined {
  const normalizedHash = hash?.trim().toLocaleLowerCase('en-US')

  return normalizedHash || undefined
}

function toPublicFolderCompareItem(item: InternalFolderCompareItem): DriveFolderCompareItem {
  return {
    id: item.id,
    path: item.path,
    name: item.name,
    type: item.type,
    size: item.size,
    lastModifiedDateTime: item.lastModifiedDateTime
  }
}

function getFolderCompareDifferenceReasons(
  sourceItem: InternalFolderCompareItem,
  targetItem: InternalFolderCompareItem
): DriveFolderCompareDifferenceReason[] {
  const reasons: DriveFolderCompareDifferenceReason[] = []

  if (sourceItem.type !== targetItem.type) {
    reasons.push('type')
    return reasons
  }

  if (sourceItem.type === 'folder') {
    return reasons
  }

  if (sourceItem.size !== targetItem.size) {
    reasons.push('content')
    return reasons
  }

  if (hasDifferentCommonContentFingerprint(sourceItem.contentFingerprint, targetItem.contentFingerprint)) {
    reasons.push('content')
  }

  return reasons
}

function hasDifferentCommonContentFingerprint(
  sourceFingerprint: DriveItemContentFingerprint | undefined,
  targetFingerprint: DriveItemContentFingerprint | undefined
): boolean {
  if (!sourceFingerprint || !targetFingerprint) {
    return false
  }

  for (const key of ['quickXorHash', 'cTag', 'eTag'] as const) {
    if (sourceFingerprint[key] && targetFingerprint[key]) {
      return sourceFingerprint[key] !== targetFingerprint[key]
    }
  }

  return false
}

export async function reconcileComparedDriveFolders(
  request: DriveFolderReconcileRequest,
  onProgress?: DriveTransferProgressListener
): Promise<DriveFolderReconcileResult> {
  const priority = validateFolderReconcilePriority(request.priority)
  const source = normalizeFolderCompareEndpoint(request.compare.source, '기준 폴더 정보를 확인하지 못했습니다.')
  const target = normalizeFolderCompareEndpoint(request.compare.target, '비교 대상 폴더 정보를 확인하지 못했습니다.')

  if (request.compare.differences.length === 0) {
    return {
      queuedCount: 0,
      sourceToTargetCount: 0,
      targetToSourceCount: 0,
      createdFolderCount: 0,
      skippedCount: 0
    }
  }

  const [sourceAccessToken, targetAccessToken] = await Promise.all([
    getGraphAccessToken(source.accountId),
    getGraphAccessToken(target.accountId)
  ])
  const [sourceRootId, targetRootId] = await Promise.all([getRootItemId(sourceAccessToken), getRootItemId(targetAccessToken)])
  const sourceFolderId = source.folderId ?? sourceRootId
  const targetFolderId = target.folderId ?? targetRootId
  const sourceFolderCache = new Map<string, string>([['', sourceFolderId]])
  const targetFolderCache = new Map<string, string>([['', targetFolderId]])
  const queuedSourceFolders = new Set<string>()
  const queuedTargetFolders = new Set<string>()
  const blockedTypePaths = new Set(
    request.compare.differences
      .filter((difference) => difference.kind === 'different' && difference.reasons.includes('type'))
      .map((difference) => normalizeComparePath(difference.path))
  )
  const tasks: InternalDriveTransferTask[] = []
  let queuedCount = 0
  let sourceToTargetCount = 0
  let targetToSourceCount = 0
  let createdFolderCount = 0
  let skippedCount = 0

  async function flushTasks(): Promise<void> {
    if (tasks.length === 0) {
      return
    }

    const nextTasks = tasks.splice(0, tasks.length)
    queuedCount += nextTasks.length
    await registerDriveTransferTasks(nextTasks)
    await writeDriveTransferState({ version: 1, tasks: nextTasks }, onProgress)
  }

  async function queueDirectionItem({
    item,
    sourceAccountId,
    targetAccountId,
    sourceToken,
    targetToken,
    targetRootItemId,
    targetBaseFolderId,
    targetCache,
    queuedFolderPaths
  }: {
    item: DriveFolderCompareItem
    sourceAccountId: string
    targetAccountId: string
    sourceToken: string
    targetToken: string
    targetRootItemId: string
    targetBaseFolderId: string
    targetCache: Map<string, string>
    queuedFolderPaths: Set<string>
  }): Promise<'queued' | 'created-folder' | 'covered' | 'skipped'> {
    if (hasAncestorComparePath(item.path, queuedFolderPaths)) {
      return 'covered'
    }

    if (item.type === 'package' || hasAncestorComparePath(item.path, blockedTypePaths)) {
      return 'skipped'
    }

    const targetParentId = await ensureRemoteDriveFolderPathRaw(
      targetRootItemId,
      targetBaseFolderId,
      getCompareParentPath(item.path),
      targetToken,
      targetCache
    )
    const transferGroupId = randomUUID()
    const expectedGroupTaskIds: string[] = []

    await queueAccountTransferItem({
      sourceAccountId,
      targetAccountId,
      sourceAccessToken: sourceToken,
      targetAccessToken: targetToken,
      targetRootId: targetRootItemId,
      targetParentId,
      item: {
        itemId: item.id,
        name: item.name,
        type: item.type,
        size: item.size
      },
      deleteSourceOnComplete: false,
      conflictBehavior: 'replace',
      onSkippedItem: () => {
        skippedCount += 1
      },
      transferGroupId,
      expectedGroupTaskIds,
      tasks,
      flushTasks
    })

    if (item.type === 'folder') {
      queuedFolderPaths.add(normalizeComparePath(item.path))
      return 'created-folder'
    }

    return 'queued'
  }

  for (const difference of [...request.compare.differences].sort(compareDifferencesForReconcile)) {
    if (difference.kind === 'only-in-source' && difference.source) {
      const status = await queueDirectionItem({
        item: difference.source,
        sourceAccountId: source.accountId,
        targetAccountId: target.accountId,
        sourceToken: sourceAccessToken,
        targetToken: targetAccessToken,
        targetRootItemId: targetRootId,
        targetBaseFolderId: targetFolderId,
        targetCache: targetFolderCache,
        queuedFolderPaths: queuedSourceFolders
      })

      if (status === 'skipped') {
        skippedCount += 1
      } else if (status !== 'covered') {
        sourceToTargetCount += 1
        createdFolderCount += status === 'created-folder' ? 1 : 0
      }
      continue
    }

    if (difference.kind === 'only-in-target' && difference.target) {
      const status = await queueDirectionItem({
        item: difference.target,
        sourceAccountId: target.accountId,
        targetAccountId: source.accountId,
        sourceToken: targetAccessToken,
        targetToken: sourceAccessToken,
        targetRootItemId: sourceRootId,
        targetBaseFolderId: sourceFolderId,
        targetCache: sourceFolderCache,
        queuedFolderPaths: queuedTargetFolders
      })

      if (status === 'skipped') {
        skippedCount += 1
      } else if (status !== 'covered') {
        targetToSourceCount += 1
        createdFolderCount += status === 'created-folder' ? 1 : 0
      }
      continue
    }

    if (difference.kind !== 'different') {
      skippedCount += 1
      continue
    }

    if (difference.reasons.includes('type') || !difference.source || !difference.target) {
      skippedCount += 1
      continue
    }

    const preferredItem = priority === 'source' ? difference.source : difference.target
    const status = await queueDirectionItem(
      priority === 'source'
        ? {
            item: preferredItem,
            sourceAccountId: source.accountId,
            targetAccountId: target.accountId,
            sourceToken: sourceAccessToken,
            targetToken: targetAccessToken,
            targetRootItemId: targetRootId,
            targetBaseFolderId: targetFolderId,
            targetCache: targetFolderCache,
            queuedFolderPaths: queuedSourceFolders
          }
        : {
            item: preferredItem,
            sourceAccountId: target.accountId,
            targetAccountId: source.accountId,
            sourceToken: targetAccessToken,
            targetToken: sourceAccessToken,
            targetRootItemId: sourceRootId,
            targetBaseFolderId: sourceFolderId,
            targetCache: sourceFolderCache,
            queuedFolderPaths: queuedTargetFolders
          }
    )

    if (status === 'skipped') {
      skippedCount += 1
    } else if (status === 'covered') {
      continue
    } else if (priority === 'source') {
      sourceToTargetCount += 1
    } else {
      targetToSourceCount += 1
    }
  }

  await flushTasks()
  scheduleDriveTransferRetryWorker(TRANSFER_DISPATCH_DELAY_MS)

  return {
    queuedCount,
    sourceToTargetCount,
    targetToSourceCount,
    createdFolderCount,
    skippedCount
  }
}

function validateFolderReconcilePriority(priority: DriveFolderReconcilePriority): DriveFolderReconcilePriority {
  if (priority === 'source' || priority === 'target') {
    return priority
  }

  throw new Error('폴더 맞추기 기준을 선택하세요.')
}

function compareDifferencesForReconcile(left: DriveFolderCompareDifference, right: DriveFolderCompareDifference): number {
  const leftDepth = getComparePathDepth(left.path)
  const rightDepth = getComparePathDepth(right.path)

  if (leftDepth !== rightDepth) {
    return leftDepth - rightDepth
  }

  return DRIVE_ITEM_NAME_COLLATOR.compare(left.path, right.path)
}

function getComparePathDepth(path: string): number {
  return splitComparePath(path).length
}

function getCompareParentPath(path: string): string {
  const segments = splitComparePath(path)

  return segments.slice(0, -1).join('/')
}

function splitComparePath(path: string): string[] {
  return path
    .split('/')
    .map((segment) => segment.trim())
    .filter(Boolean)
}

function hasAncestorComparePath(path: string, ancestors: Set<string>): boolean {
  const segments = splitComparePath(path)
  let currentPath = ''

  for (let index = 0; index < segments.length - 1; index += 1) {
    currentPath = joinComparePath(currentPath, segments[index])

    if (ancestors.has(normalizeComparePath(currentPath))) {
      return true
    }
  }

  return false
}

async function ensureRemoteDriveFolderPathRaw(
  rootItemId: string,
  baseFolderId: string,
  relativePath: string,
  accessToken: string,
  folderCache: Map<string, string>
): Promise<string> {
  const segments = splitComparePath(relativePath)
  let parentId = baseFolderId
  let currentPath = ''

  for (const segment of segments) {
    currentPath = joinComparePath(currentPath, segment)
    const cacheKey = normalizeComparePath(currentPath)
    const cachedFolderId = folderCache.get(cacheKey)

    if (cachedFolderId) {
      parentId = cachedFolderId
      continue
    }

    const folder = await ensureRemoteDriveFolderRaw(rootItemId, parentId, segment, accessToken)
    folderCache.set(cacheKey, folder.id)
    parentId = folder.id
  }

  return parentId
}

function joinComparePath(parentPath: string, name: string): string {
  return parentPath ? `${parentPath}/${name}` : name
}

function normalizeComparePath(path: string): string {
  return path.normalize('NFC').trim().toLocaleLowerCase('ko-KR')
}

export async function listDriveAccountUsage(accounts: AuthAccount[]): Promise<DriveAccountUsage[]> {
  return Promise.all(
    accounts.map(async (account) => {
      try {
        const accessToken = await getGraphAccessToken(account.homeAccountId)
        const drive = await graphGet<GraphDrive>(createDriveUsageUrl(), accessToken)
        const quota = drive.quota ?? {}

        return {
          accountId: account.homeAccountId,
          used: Math.max(quota.used ?? 0, 0),
          total: quota.total === undefined ? undefined : Math.max(quota.total, 0),
          remaining: quota.remaining === undefined ? undefined : Math.max(quota.remaining, 0),
          state: quota.state
        }
      } catch (error) {
        return {
          accountId: account.homeAccountId,
          used: 0,
          isUnavailable: true,
          error: error instanceof Error ? error.message : 'OneDrive 사용량을 확인하지 못했습니다.'
        }
      }
    })
  )
}

export async function clearDriveIndexMemory(): Promise<void> {
  driveIndexGeneration += 1
  driveIndexSyncPromises.clear()
  activeDriveIndexSnapshot = null
  activeDriveIndexSnapshotAccountId = null
}

export async function resetDriveIndex(): Promise<void> {
  await clearDriveIndexMemory()
  closeDriveIndexStores()

  try {
    await rm(getDriveIndexStoreDirectory(), { recursive: true, force: true })
    await unlink(getLegacyDriveIndexPath())
  } catch {
    // Index file may not exist on a fresh install.
  }
}

export async function resetDriveTransfers(): Promise<void> {
  stopDriveTransferRetryScheduler()
  runningTransferIds.clear()
  transferWorkerSlotIds.clear()
  transferAbortControllers.forEach((controller) => controller.abort())
  transferAbortControllers.clear()
  transferPauseRequests.clear()
  transferDeleteRequests.clear()
  transferIndexCache = null
  transferScanCursor = 0
  transferThrottleUntil = 0
  nextRetryingTransferDispatchAt = 0
  hasCheckedLegacyDriveTransfers = false

  try {
    await rm(getDriveTransfersDirectory(), { recursive: true, force: true })
    await unlink(getLegacyDriveTransfersPath())
  } catch {
    // Transfer state file may not exist on a fresh install.
  }
}

export async function resetDriveThumbnailCache(): Promise<void> {
  const resetError = new Error('썸네일 캐시가 초기화되었습니다.')

  for (const entry of driveThumbnailQueue) {
    entry.reject(resetError)
  }

  driveThumbnailCache.clear()
  driveThumbnailPendingRequests.clear()
  driveThumbnailWarmPromises.clear()
  driveThumbnailWarmSignatures.clear()
  driveThumbnailQueue.splice(0, driveThumbnailQueue.length)
  driveThumbnailThrottleUntil = 0
  resetDriveThumbnailAdaptiveRate()

  if (driveThumbnailQueueTimer) {
    clearTimeout(driveThumbnailQueueTimer)
    driveThumbnailQueueTimer = null
  }

  try {
    await rm(getDriveThumbnailCacheDirectory(), { recursive: true, force: true })
  } catch {
    // Thumbnail cache may not exist yet.
  }
}

export async function bindUnscopedDriveTransfersToActiveAccount(): Promise<void> {
  const activeAccountId = await getActiveAccountId()

  if (!activeAccountId) {
    return
  }

  const index = await readDriveTransferIndex()

  for (const taskId of index.taskIds) {
    const task = await readDriveTransferTask(taskId)

    if (task && !task.accountId) {
      task.accountId = activeAccountId
      await writeDriveTransferTask(task)
    }
  }
}

export async function renameDriveItem(request: RenameDriveItemRequest): Promise<CloudDriveItem> {
  const itemId = validateDriveItemId(request.itemId)
  const name = validateDriveItemName(request.name)
  const accessToken = await getGraphAccessToken()
  const updatedItem = await graphSend<GraphDriveItem>(createItemUrl(itemId, true), accessToken, {
    method: 'PATCH',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ name })
  })

  return mergeRemoteDriveItem(updatedItem)
}

export async function deleteDriveItem(request: DeleteDriveItemRequest): Promise<void> {
  const itemId = validateDriveItemId(request.itemId)
  const accessToken = await getGraphAccessToken()

  await graphSend<void>(createItemUrl(itemId), accessToken, {
    method: 'DELETE'
  })

  const index = await getCurrentDriveIndex()
  removeIndexedItem(index, itemId)
  activeDriveIndexSnapshot = index
  await writeDriveIndex(index)
}

export async function copyDriveItems(request: CopyDriveItemsRequest): Promise<CopyDriveItemsResult> {
  const itemIds = [...new Set(request.itemIds.map((itemId) => validateDriveItemId(itemId)))]

  if (itemIds.length === 0) {
    return {
      items: [],
      hasPendingOperations: false
    }
  }

  const index = await ensureDriveIndexForListing(false)
  const targetParentId = request.parentId ?? index.rootItemId

  if (!targetParentId) {
    throw new Error('복사할 대상 폴더 정보를 확인하지 못했습니다.')
  }

  const accessToken = await getGraphAccessToken()
  const driveId = await getDriveId(accessToken)
  const currentIndex = await fetchAndMergeFolderChildren(index, targetParentId)
  const occupiedTargetNames = new Set(
    Object.values(currentIndex.items)
      .filter((item) => item.parentId === targetParentId)
      .map((item) => normalizeDriveItemNameForConflict(item.name))
  )
  let hasPendingOperations = false

  for (const itemId of itemIds) {
    const item = currentIndex.items[itemId]

    if (itemId === targetParentId || isIndexedDescendant(currentIndex, targetParentId, itemId)) {
      throw new Error('폴더를 자기 자신이나 하위 폴더로 복사할 수 없습니다.')
    }

    const nextName = item ? getAvailableCopyName(item.name, item.type, occupiedTargetNames) : undefined

    if (item) {
      occupiedTargetNames.add(normalizeDriveItemNameForConflict(nextName ?? item.name))
    }

    const didComplete = await copyDriveItemToFolder(itemId, targetParentId, driveId, accessToken, nextName)
    hasPendingOperations = hasPendingOperations || !didComplete
  }

  const updatedIndex = await fetchAndMergeFolderChildren(await getCurrentDriveIndex(), targetParentId)

  return {
    items: sortDriveItems(Object.values(updatedIndex.items).filter((item) => item.parentId === targetParentId)),
    hasPendingOperations
  }
}

export async function moveDriveItems(request: MoveDriveItemsRequest): Promise<CloudDriveItem[]> {
  const itemIds = [...new Set(request.itemIds.map((itemId) => validateDriveItemId(itemId)))]

  if (itemIds.length === 0) {
    return []
  }

  const index = await ensureDriveIndexForListing(false)
  const targetParentId = request.parentId ?? index.rootItemId

  if (!targetParentId) {
    throw new Error('이동할 대상 폴더 정보를 확인하지 못했습니다.')
  }

  const accessToken = await getGraphAccessToken()
  const movedItems: CloudDriveItem[] = []

  for (const itemId of itemIds) {
    const item = index.items[itemId]

    if (itemId === targetParentId || item?.parentId === targetParentId) {
      continue
    }

    if (isIndexedDescendant(index, targetParentId, itemId)) {
      throw new Error('폴더를 자기 하위 폴더로 이동할 수 없습니다.')
    }

    const updatedItem = await graphSend<GraphDriveItem>(createItemUrl(itemId, true), accessToken, {
      method: 'PATCH',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        parentReference: {
          id: targetParentId
        }
      })
    })

    movedItems.push(await mergeRemoteDriveItem(updatedItem, targetParentId))
  }

  return movedItems
}

export async function transferDriveItemsBetweenAccounts(
  request: TransferDriveItemsBetweenAccountsRequest,
  onProgress?: DriveTransferProgressListener
): Promise<TransferDriveItemsBetweenAccountsResult> {
  const sourceAccountId = validateAccountId(request.sourceAccountId, '원본 계정을 확인하지 못했습니다.')
  const targetAccountId = validateAccountId(request.targetAccountId, '대상 계정을 확인하지 못했습니다.')

  if (sourceAccountId === targetAccountId) {
    throw new Error('계정 간 전송은 서로 다른 계정 사이에서만 사용할 수 있습니다.')
  }

  if (request.items.length === 0) {
    return { queuedCount: 0 }
  }

  const sourceAccessToken = await getGraphAccessToken(sourceAccountId)
  const targetAccessToken = await getGraphAccessToken(targetAccountId)
  const targetRootId = await getRootItemId(targetAccessToken)
  const targetParentId = request.targetParentId ?? targetRootId
  const tasks: InternalDriveTransferTask[] = []
  let queuedCount = 0

  async function flushTasks(): Promise<void> {
    if (tasks.length === 0) {
      return
    }

    const nextTasks = tasks.splice(0, tasks.length)
    queuedCount += nextTasks.length
    await registerDriveTransferTasks(nextTasks)
    await writeDriveTransferState({ version: 1, tasks: nextTasks }, onProgress)
  }

  for (const item of request.items) {
    const transferGroupId = randomUUID()
    const expectedGroupTaskIds: string[] = []

    await queueAccountTransferItem({
      sourceAccountId,
      targetAccountId,
      sourceAccessToken,
      targetAccessToken,
      targetRootId,
      targetParentId,
      item,
      deleteSourceOnComplete: item.type === 'file' && (request.deleteSourceOnComplete ?? false),
      transferGroupId,
      expectedGroupTaskIds,
      tasks,
      flushTasks
    })

    if (item.type === 'folder' && request.deleteSourceOnComplete) {
      tasks.push(createAccountTransferCleanupTask(sourceAccountId, targetAccountId, item, transferGroupId, expectedGroupTaskIds))
    }
  }

  await flushTasks()
  scheduleDriveTransferRetryWorker(TRANSFER_DISPATCH_DELAY_MS)

  return { queuedCount }
}

export async function listDriveTransfers(): Promise<DriveTransferTask[]> {
  return publicTransferTasks((await readDriveTransferState()).tasks)
}

export async function listDriveTransferPage(request: DriveTransferListRequest = {}): Promise<DriveTransferListResult> {
  await migrateLegacyDriveTransfers()

  const activeAccountId = await getActiveAccountId()
  const offset = Math.max(0, Math.floor(request.offset ?? 0))
  const limit = Math.min(TRANSFER_LIST_MAX_LIMIT, Math.max(1, Math.floor(request.limit ?? TRANSFER_VISIBLE_LIMIT)))
  const index = await readDriveTransferIndex()
  const completedSummary = (await hasIndexedTransferTaskForAccount(index, activeAccountId))
    ? await readCompletedTransferSummary(activeAccountId)
    : await resetCompletedTransferSummary(activeAccountId)
  const pageTasks: InternalDriveTransferTask[] = []
  const summary: DriveTransferSummary = {
    totalCount: completedSummary.completedCount,
    activeCount: 0,
    completedCount: completedSummary.completedCount,
    queuedCount: 0,
    runningCount: 0,
    pausedCount: 0,
    retryingCount: 0,
    failedCount: 0,
    totalBytes: completedSummary.completedBytes,
    transferredBytes: completedSummary.completedBytes,
    bytesPerSecond: 0
  }
  let visibleTaskIndex = 0

  for (const taskId of index.taskIds) {
    const task = await readDriveTransferTask(taskId)

    if (!task || !isDriveTransferTaskForAccount(task, activeAccountId)) {
      continue
    }

    if (!task.cleanupOnly) {
      applyTransferTaskToSummary(summary, task)
    }

    if (visibleTaskIndex >= offset && pageTasks.length < limit) {
      pageTasks.push(normalizeLoadedTransferTask(task))
    }

    visibleTaskIndex += 1
  }

  return {
    tasks: publicTransferTasks(sortTransferTasks(pageTasks)),
    summary,
    offset,
    limit,
    totalTaskCount: visibleTaskIndex
  }
}

export function startDriveTransferRetryScheduler(onProgress?: DriveTransferProgressListener): void {
  transferRetryProgressListener = onProgress
  scheduleDriveTransferRetryWorker(TRANSFER_DISPATCH_DELAY_MS)
}

export function wakeDriveTransferQueue(): void {
  scheduleDriveTransferRetryWorker(TRANSFER_DISPATCH_DELAY_MS)
}

export async function resumeDriveTransfers(taskId?: string, onProgress?: DriveTransferProgressListener): Promise<DriveTransferTask[]> {
  const selectedTask = taskId ? await readDriveTransferTask(taskId.trim()) : null
  const resumableTasks = taskId ? (selectedTask && isTransferReadyToRun(selectedTask, true) ? [selectedTask] : []) : await findRetryableTransferTasks(TRANSFER_RETRY_BATCH_LIMIT, true)

  for (const task of resumableTasks) {
    if (!task || transferWorkerSlotIds.has(task.id) || runningTransferIds.has(task.id)) {
      continue
    }

    task.status = 'queued'
    task.bytesPerSecond = 0
    task.nextRetryAt = undefined
    task.message = `${getTransferKindLabelForMessage(task.kind)} 대기 중`
    task.updatedAt = new Date().toISOString()
    await writeDriveTransferTask(task)
  }

  scheduleDriveTransferRetryWorker(TRANSFER_DISPATCH_DELAY_MS)
  await emitTransferSnapshot(onProgress)
  return listDriveTransfers()
}

export async function stopDriveTransfer(taskId: string, onProgress?: DriveTransferProgressListener): Promise<DriveTransferTask[]> {
  const normalizedTaskId = taskId.trim()

  if (!normalizedTaskId) {
    throw new Error('중지할 전송 작업을 선택하세요.')
  }

  transferPauseRequests.add(normalizedTaskId)
  transferAbortControllers.get(normalizedTaskId)?.abort()

  const task = await readDriveTransferTask(normalizedTaskId)

  if (task && task.status !== 'completed') {
    task.status = 'paused'
    task.bytesPerSecond = 0
    task.nextRetryAt = undefined
    task.message = '전송이 중지되었습니다.'
    task.updatedAt = new Date().toISOString()
    await writeDriveTransferTask(task)
    await emitTransferSnapshot(onProgress)
  }

  return listDriveTransfers()
}

export async function deleteDriveTransfer(taskId?: string, onProgress?: DriveTransferProgressListener): Promise<DriveTransferTask[]> {
  if (!taskId) {
    const index = await readDriveTransferIndex()
    const activeAccountId = await getActiveAccountId()
    const taskIdsToDelete: string[] = []

    for (const taskId of index.taskIds) {
      const task = await readDriveTransferTask(taskId)

      if (task && isDriveTransferTaskForAccount(task, activeAccountId)) {
        taskIdsToDelete.push(task.id)
        transferDeleteRequests.add(task.id)
        transferPauseRequests.add(task.id)
        transferAbortControllers.get(task.id)?.abort()
        await cleanupTransferTask(task)
      }
    }

    for (const taskId of taskIdsToDelete) {
      await deleteDriveTransferTaskFile(taskId)
    }

    await removeDriveTransferIndexTaskIds(taskIdsToDelete)
    await resetCompletedTransferSummary(activeAccountId)
    for (const taskId of taskIdsToDelete) {
      if (!transferWorkerSlotIds.has(taskId) && !runningTransferIds.has(taskId)) {
        transferDeleteRequests.delete(taskId)
        transferPauseRequests.delete(taskId)
      }
    }

    transferScanCursor = 0
    await emitTransferSnapshot(onProgress)
    return []
  }

  const normalizedTaskId = taskId.trim()
  const task = await readDriveTransferTask(normalizedTaskId)
  const tasksToDelete = task ? [task] : []

  for (const task of tasksToDelete) {
    transferDeleteRequests.add(task.id)
    transferPauseRequests.add(task.id)
    transferAbortControllers.get(task.id)?.abort()
    await cleanupTransferTask(task)
  }

  await deleteDriveTransferTaskFile(normalizedTaskId)
  await removeDriveTransferIndexTaskIds([normalizedTaskId])
  await emitTransferSnapshot(onProgress)

  for (const task of tasksToDelete) {
    if (!transferWorkerSlotIds.has(task.id) && !runningTransferIds.has(task.id)) {
      transferDeleteRequests.delete(task.id)
      transferPauseRequests.delete(task.id)
    }
  }

  return listDriveTransfers()
}

export async function uploadLocalFilesToDrive(
  parentId: string | null | undefined,
  localPaths: string[],
  onProgress?: DriveTransferProgressListener
): Promise<CloudDriveItem[]> {
  if (localPaths.length === 0) {
    return []
  }

  const index = await ensureDriveIndexForListing(false)
  const targetParentId = parentId ?? index.rootItemId

  if (!targetParentId) {
    throw new Error('업로드할 OneDrive 폴더 정보를 확인하지 못했습니다.')
  }

  await queueLocalUploadPaths(targetParentId, localPaths, onProgress)
  scheduleDriveTransferRetryWorker(TRANSFER_DISPATCH_DELAY_MS)

  return []
}

async function queueLocalUploadPaths(
  targetParentId: string,
  localPaths: string[],
  onProgress?: DriveTransferProgressListener
): Promise<void> {
  const index = await ensureDriveIndexForListing(false)
  const accessToken = await getGraphAccessToken()
  let pendingFiles: LocalUploadFile[] = []

  async function flushPendingFiles(): Promise<void> {
    if (pendingFiles.length === 0) {
      return
    }

    const batch = pendingFiles
    pendingFiles = []
    await queueUploadFileBatch(batch, onProgress)
  }

  for (const localPath of localPaths) {
    const localStat = await stat(localPath)

    if (localStat.isFile()) {
      pendingFiles.push({ parentId: targetParentId, localPath })

      if (pendingFiles.length >= UPLOAD_QUEUE_BATCH_SIZE) {
        await flushPendingFiles()
      }

      continue
    }

    if (localStat.isDirectory()) {
      const remoteFolder = await ensureRemoteDriveFolder(index, targetParentId, basename(localPath), accessToken)
      await queueLocalDirectoryUpload(index, remoteFolder.id, localPath, accessToken, async (file) => {
        pendingFiles.push(file)

        if (pendingFiles.length >= UPLOAD_QUEUE_BATCH_SIZE) {
          await flushPendingFiles()
        }
      })
    }
  }

  await flushPendingFiles()
}

async function queueLocalDirectoryUpload(
  index: DriveIndex,
  targetParentId: string,
  localDirectoryPath: string,
  accessToken: string,
  onFile: (file: LocalUploadFile) => Promise<void>
): Promise<void> {
  const entries = await readdir(localDirectoryPath, { withFileTypes: true })

  for (const entry of entries) {
    if (entry.isSymbolicLink()) {
      continue
    }

    const entryPath = join(localDirectoryPath, entry.name)

    if (entry.isFile()) {
      await onFile({ parentId: targetParentId, localPath: entryPath })
      continue
    }

    if (entry.isDirectory()) {
      const remoteFolder = await ensureRemoteDriveFolder(index, targetParentId, entry.name, accessToken)
      await queueLocalDirectoryUpload(index, remoteFolder.id, entryPath, accessToken, onFile)
    }
  }
}

async function queueAccountTransferItem(context: AccountTransferQueueContext): Promise<void> {
  const sourceItemId = validateDriveItemId(context.item.itemId)
  const itemName = validateDriveItemName(context.item.name)

  if (context.item.type === 'package') {
    if (context.onSkippedItem) {
      context.onSkippedItem()
      return
    }

    throw new Error(`패키지 항목은 계정 간 전송을 아직 지원하지 않습니다: ${itemName}`)
  }

  if (context.item.type === 'file') {
    context.tasks.push(createAccountTransferTask(context, sourceItemId, itemName, context.item.size ?? 0))

    if (context.tasks.length >= UPLOAD_QUEUE_BATCH_SIZE) {
      await context.flushTasks()
    }

    return
  }

  const targetFolder = await ensureRemoteDriveFolderRaw(
    context.targetRootId,
    context.targetParentId,
    itemName,
    context.targetAccessToken
  )

  await queueAccountTransferFolderChildren({
    ...context,
    targetParentId: targetFolder.id,
    deleteSourceOnComplete: false,
    item: {
      ...context.item,
      itemId: sourceItemId,
      name: itemName
    }
  })
}

async function queueAccountTransferFolderChildren(context: AccountTransferQueueContext): Promise<void> {
  let url = createChildrenUrlForItem(context.item.itemId)

  while (true) {
    const response = await graphGet<GraphChildrenResponse>(url, context.sourceAccessToken)

    for (const child of response.value ?? []) {
      if (!child.id) {
        continue
      }

      const item = mapDriveItem(child)

      await queueAccountTransferItem({
        ...context,
        item: {
          itemId: child.id,
          name: item.name,
          type: item.type,
          size: item.size
        }
      })
    }

    if (!response['@odata.nextLink']) {
      break
    }

    url = parseGraphUrl(response['@odata.nextLink'], 'OneDrive 폴더 다음 페이지 주소가 올바르지 않습니다.')
  }
}

function createAccountTransferTask(
  context: AccountTransferQueueContext,
  sourceItemId: string,
  name: string,
  sourceSize: number
): InternalDriveTransferTask {
  const now = new Date().toISOString()
  const taskId = randomUUID()
  const normalizedSourceSize = Math.max(sourceSize, 0)

  context.expectedGroupTaskIds.push(taskId)

  return {
    id: taskId,
    accountId: context.targetAccountId,
    sourceAccountId: context.sourceAccountId,
    targetAccountId: context.targetAccountId,
    sourceItemId,
    targetParentId: context.targetParentId,
    sourceSize: normalizedSourceSize,
    downloadedBytes: 0,
    uploadedBytes: 0,
    transferGroupId: context.transferGroupId,
    deleteSourceOnComplete: context.deleteSourceOnComplete,
    conflictBehavior: context.conflictBehavior,
    kind: 'account-transfer',
    status: 'queued',
    name,
    tempPath: getAccountTransferTempPath(taskId),
    transferredBytes: 0,
    totalBytes: normalizedSourceSize * 2,
    createdAt: now,
    updatedAt: now,
    message: context.deleteSourceOnComplete ? '계정 간 이동 대기 중' : '계정 간 복사 대기 중'
  }
}

function createAccountTransferCleanupTask(
  sourceAccountId: string,
  targetAccountId: string,
  item: TransferDriveItemRef,
  transferGroupId: string,
  expectedGroupTaskIds: string[]
): InternalDriveTransferTask {
  const now = new Date().toISOString()

  return {
    id: randomUUID(),
    accountId: targetAccountId,
    sourceAccountId,
    targetAccountId,
    sourceItemId: validateDriveItemId(item.itemId),
    sourceSize: 0,
    downloadedBytes: 0,
    uploadedBytes: 0,
    transferGroupId,
    expectedGroupTaskIds: [...new Set(expectedGroupTaskIds)],
    completedGroupTaskIds: [],
    cleanupOnly: true,
    deleteSourceOnComplete: true,
    kind: 'account-transfer',
    status: 'queued',
    name: validateDriveItemName(item.name),
    transferredBytes: 0,
    totalBytes: 0,
    createdAt: now,
    updatedAt: now,
    message: '원본 폴더 삭제 대기 중'
  }
}

async function queueUploadFileBatch(files: LocalUploadFile[], onProgress?: DriveTransferProgressListener): Promise<void> {
  const tasks: InternalDriveTransferTask[] = []
  const accountId = await getActiveAccountId()

  for (const file of files) {
    const fileStat = await stat(file.localPath)

    if (!fileStat.isFile()) {
      continue
    }

    const fileName = validateDriveItemName(basename(file.localPath))
    const now = new Date().toISOString()
    const task: InternalDriveTransferTask = {
      id: randomUUID(),
      accountId: accountId ?? undefined,
      kind: 'upload',
      status: 'queued',
      name: fileName,
      localPath: file.localPath,
      parentId: file.parentId,
      transferredBytes: 0,
      totalBytes: fileStat.size,
      createdAt: now,
      updatedAt: now,
      message: '업로드 대기 중'
    }

    tasks.push(task)
  }

  if (tasks.length === 0) {
    return
  }

  await registerDriveTransferTasks(tasks)
  await writeDriveTransferState({ version: 1, tasks }, onProgress)
}

export async function downloadDriveItemToPath(
  itemId: string,
  name: string,
  totalBytes: number | undefined,
  localPath: string,
  onProgress?: DriveTransferProgressListener
): Promise<void> {
  const state = await readDriveTransferState()
  const accountId = await getActiveAccountId()
  const task = queueDownloadTask(state, itemId, name, totalBytes, localPath, accountId)

  await registerDriveTransferTasks([task])
  await writeDriveTransferState({ version: 1, tasks: [task] }, onProgress)
  scheduleDriveTransferRetryWorker(TRANSFER_DISPATCH_DELAY_MS)
}

export async function downloadDriveItemsToDirectory(
  items: DownloadDriveItemRequest[],
  directoryPath: string,
  onProgress?: DriveTransferProgressListener
): Promise<Required<Pick<DownloadDriveItemsResult, 'queuedCount' | 'createdFolderCount' | 'skippedCount'>>> {
  const normalizedDirectoryPath = normalizeDownloadDirectoryPath(directoryPath)

  if (!normalizedDirectoryPath.trim()) {
    throw new Error('다운로드할 폴더를 선택하세요.')
  }

  if (items.length === 0) {
    throw new Error('다운로드할 항목을 선택하세요.')
  }

  await ensureLocalDirectory(normalizedDirectoryPath, '다운로드 대상 폴더를 준비하지 못했습니다.')

  const reservedPaths = new Set<string>()
  const state = await readDriveTransferState()
  const tasks: InternalDriveTransferTask[] = []
  const accountId = await getActiveAccountId()
  const accessToken = items.some((item) => item.type === 'folder') ? await getGraphAccessToken(accountId ?? undefined) : null
  const stats = {
    queuedCount: 0,
    createdFolderCount: 0,
    skippedCount: 0
  }

  async function flushTasks(): Promise<void> {
    if (tasks.length === 0) {
      return
    }

    const nextTasks = tasks.splice(0, tasks.length)
    await registerDriveTransferTasks(nextTasks)
    await writeDriveTransferState({ version: 1, tasks: nextTasks }, onProgress)
    scheduleDriveTransferRetryWorker(TRANSFER_DISPATCH_DELAY_MS)
  }

  for (const item of items) {
    const name = validateDriveItemName(item.name)
    const localPath = await createUniqueDownloadPath(normalizedDirectoryPath, name, reservedPaths)

    reservedPaths.add(normalizeLocalPathKey(localPath))
    await queueDriveItemDownloadToPath({
      state,
      tasks,
      item: {
        ...item,
        name
      },
      localPath,
      accountId,
      accessToken,
      reservedPaths,
      stats,
      flushTasks
    })
  }

  await flushTasks()
  scheduleDriveTransferRetryWorker(TRANSFER_DISPATCH_DELAY_MS)
  return stats
}

async function queueDriveItemDownloadToPath({
  state,
  tasks,
  item,
  localPath,
  accountId,
  accessToken,
  reservedPaths,
  stats,
  flushTasks
}: {
  state: DriveTransferState
  tasks: InternalDriveTransferTask[]
  item: DownloadDriveItemRequest
  localPath: string
  accountId: string | null
  accessToken: string | null
  reservedPaths: Set<string>
  stats: Required<Pick<DownloadDriveItemsResult, 'queuedCount' | 'createdFolderCount' | 'skippedCount'>>
  flushTasks: () => Promise<void>
}): Promise<void> {
  const itemType = item.type ?? 'file'

  if (itemType === 'package') {
    stats.skippedCount += 1
    return
  }

  if (itemType === 'folder') {
    if (!accessToken) {
      throw new Error('폴더 다운로드를 위한 OneDrive 인증 정보를 찾지 못했습니다.')
    }

    stats.createdFolderCount += 1
    await ensureLocalDirectory(localPath, '다운로드 폴더를 만들지 못했습니다.')
    await queueDriveFolderChildrenDownload({
      state,
      tasks,
      folderId: item.itemId,
      localDirectoryPath: localPath,
      accountId,
      accessToken,
      reservedPaths,
      stats,
      flushTasks
    })
    return
  }

  tasks.push(queueDownloadTask(state, item.itemId, item.name, item.size, localPath, accountId))
  stats.queuedCount += 1

  if (tasks.length >= UPLOAD_QUEUE_BATCH_SIZE) {
    await flushTasks()
  }
}

async function queueDriveFolderChildrenDownload({
  state,
  tasks,
  folderId,
  localDirectoryPath,
  accountId,
  accessToken,
  reservedPaths,
  stats,
  flushTasks
}: {
  state: DriveTransferState
  tasks: InternalDriveTransferTask[]
  folderId: string
  localDirectoryPath: string
  accountId: string | null
  accessToken: string
  reservedPaths: Set<string>
  stats: Required<Pick<DownloadDriveItemsResult, 'queuedCount' | 'createdFolderCount' | 'skippedCount'>>
  flushTasks: () => Promise<void>
}): Promise<void> {
  let url = createChildrenUrlForItem(validateDriveItemId(folderId))

  while (true) {
    const response = await graphGet<GraphChildrenResponse>(url, accessToken)

    for (const child of response.value ?? []) {
      if (!child.id) {
        continue
      }

      const childItem = mapDriveItem(child)
      const childName = validateDriveItemName(childItem.name)
      const localPath = await createUniqueDownloadPath(localDirectoryPath, childName, reservedPaths)

      reservedPaths.add(normalizeLocalPathKey(localPath))
      await queueDriveItemDownloadToPath({
        state,
        tasks,
        item: {
          itemId: childItem.id,
          name: childName,
          type: childItem.type,
          size: childItem.size
        },
        localPath,
        accountId,
        accessToken,
        reservedPaths,
        stats,
        flushTasks
      })
    }

    if (!response['@odata.nextLink']) {
      break
    }

    url = parseGraphUrl(response['@odata.nextLink'], 'OneDrive 폴더 다음 페이지 주소가 올바르지 않습니다.')
  }
}

function queueDownloadTask(
  state: DriveTransferState,
  itemId: string,
  name: string,
  totalBytes: number | undefined,
  localPath: string,
  accountId: string | null
): InternalDriveTransferTask {
  const normalizedItemId = validateDriveItemId(itemId)
  const existingTask = state.tasks.find(
    (task) => task.kind === 'download' && task.status !== 'completed' && task.itemId === normalizedItemId && task.localPath === localPath
  )

  if (existingTask) {
    existingTask.status = 'queued'
    existingTask.bytesPerSecond = 0
    existingTask.message = '다운로드 대기 중'
    existingTask.updatedAt = new Date().toISOString()
    return existingTask
  }

  const now = new Date().toISOString()
  const task: InternalDriveTransferTask = {
    id: randomUUID(),
    accountId: accountId ?? undefined,
    kind: 'download',
    status: 'queued',
    name: validateDriveItemName(name),
    itemId: normalizedItemId,
    localPath,
    tempPath: `${localPath}.onedrive-download`,
    transferredBytes: 0,
    totalBytes: Math.max(totalBytes ?? 0, 0),
    createdAt: now,
    updatedAt: now,
    message: '다운로드 대기 중'
  }

  state.tasks.unshift(task)
  return task
}

async function ensureDriveIndexForListing(forceSync: boolean): Promise<DriveIndex> {
  const accountId = await getActiveAccountId()
  let currentIndex = await getCurrentDriveIndex(accountId)
  const driveSettings = await getDriveSettings()

  if (!currentIndex.rootItemId) {
    currentIndex = await ensureRootItemId(currentIndex, accountId)
  }

  if (driveSettings.indexMode === 'automatic' && (forceSync || !isIndexReadyAndFresh(currentIndex)) && !isDriveIndexSyncing(accountId)) {
    emitGraphActivity({
      level: 'info',
      scope: 'index',
      title: '탐색 인덱스 갱신 예약',
      message: '폴더 목록을 표시하면서 OneDrive 전체 인덱스를 갱신합니다.'
    })
    startDriveIndexSync(currentIndex, accountId)
  }

  return currentIndex
}

function startDriveIndexSync(existingIndex: DriveIndex, accountId: string | null): Promise<DriveIndex> {
  const generation = driveIndexGeneration
  const syncKey = getDriveIndexSyncKey(accountId)

  activeDriveIndexSnapshot = existingIndex
  activeDriveIndexSnapshotAccountId = accountId
  const syncPromise = syncDriveIndex(existingIndex, generation, accountId)
    .then((index) => {
      if (generation === driveIndexGeneration) {
        activeDriveIndexSnapshot = index
        activeDriveIndexSnapshotAccountId = accountId
        startDriveThumbnailCacheWarmup(index, accountId, generation)
      }

      return index
    })
    .catch((error) => {
      if (generation === driveIndexGeneration) {
        console.error('OneDrive index sync failed:', error)
      }

      return activeDriveIndexSnapshotAccountId === accountId ? activeDriveIndexSnapshot ?? existingIndex : existingIndex
    })
    .finally(() => {
      if (generation === driveIndexGeneration && driveIndexSyncPromises.get(syncKey) === syncPromise) {
        driveIndexSyncPromises.delete(syncKey)
      }
    })

  driveIndexSyncPromises.set(syncKey, syncPromise)
  return syncPromise
}

async function syncDriveIndex(existingIndex: DriveIndex, generation: number, accountId: string | null): Promise<DriveIndex> {
  const accessToken = await getGraphAccessToken(accountId)
  const root = await graphGet<GraphDriveRoot>(createRootUrl(), accessToken)
  const rootItemId = root.id

  if (!rootItemId) {
    throw new Error('OneDrive 루트 폴더 정보를 확인하지 못했습니다.')
  }

  if (generation !== driveIndexGeneration) {
    return existingIndex
  }

  const index: DriveIndex =
    existingIndex.rootItemId && existingIndex.rootItemId !== rootItemId
      ? createEmptyDriveIndex(rootItemId)
      : {
          ...existingIndex,
          expandedFolderIds: existingIndex.expandedFolderIds ?? {},
          rootItemId
        }
  let shouldRewriteDriveIndexStore = !parseStoredDriveDeltaUrl(index.deltaLink)

  activeDriveIndexSnapshot = index
  activeDriveIndexSnapshotAccountId = accountId

  const storedDeltaUrl = parseStoredDriveDeltaUrl(index.deltaLink)

  if (index.deltaLink && !storedDeltaUrl) {
    Object.assign(index, createEmptyDriveIndex(rootItemId))
    activeDriveIndexSnapshot = index
    activeDriveIndexSnapshotAccountId = accountId
    shouldRewriteDriveIndexStore = true
  }

  let url = storedDeltaUrl ?? createDeltaUrl()
  let didRestartFromFullEnumeration = false
  let deltaSequence = 0

  emitGraphActivity({
    level: 'info',
    scope: 'index',
    title: storedDeltaUrl ? '인덱스 변경분 확인 중' : '전체 인덱스 수집 중',
    message: storedDeltaUrl ? 'OneDrive 변경분(delta)을 가져오는 중입니다.' : 'OneDrive 전체 항목을 페이지 단위로 수집하는 중입니다.',
    progress: {
      current: 0,
      indeterminate: true
    }
  })

  await resetDriveDeltaStaging(accountId)

  while (true) {
    if (generation !== driveIndexGeneration) {
      return index
    }

    try {
      const response = await graphGet<GraphDeltaResponse>(url, accessToken, {
        deltaExcludeParent: 'true'
      })

      if (generation !== driveIndexGeneration) {
        return index
      }

      const stagePayloads = createDriveDeltaStagePayloads(response.value ?? [], deltaSequence)

      if (stagePayloads.length > 0) {
        deltaSequence += stagePayloads.length
        await appendDriveDeltaStaging(accountId, stagePayloads)
        emitGraphActivity({
          level: 'info',
          scope: 'index',
          title: '인덱스 항목 수집 중',
          message: `${deltaSequence.toLocaleString('ko-KR')}개 OneDrive 항목 정보를 수집했습니다.`,
          progress: {
            current: deltaSequence,
            indeterminate: true
          }
        })
      }

      activeDriveIndexSnapshot = index
      activeDriveIndexSnapshotAccountId = accountId

      if (response['@odata.nextLink']) {
        url = parseDriveDeltaUrl(response['@odata.nextLink'], 'OneDrive 인덱스 다음 페이지 주소가 올바르지 않습니다.')
        continue
      }

      emitGraphActivity({
        level: 'info',
        scope: 'index',
        title: '인덱스 적용 중',
        message: `${deltaSequence.toLocaleString('ko-KR')}개 변경 항목을 로컬 인덱스에 적용합니다.`,
        progress: {
          current: deltaSequence,
          total: Math.max(deltaSequence, 1)
        }
      })
      const indexChanges = await applyStagedDriveDelta(index, accountId)
      index.deltaLink = normalizeDriveDeltaLink(response['@odata.deltaLink'])
      index.syncedAt = new Date().toISOString()

      if (shouldRewriteDriveIndexStore) {
        await writeDriveIndex(index, accountId)
      } else {
        activeDriveIndexSnapshot = index
        activeDriveIndexSnapshotAccountId = accountId
        await applyDriveIndexDeltaToStore(index, accountId, indexChanges)
      }

      await clearDriveDeltaStaging(accountId)
      emitGraphActivity({
        level: 'success',
        scope: 'index',
        title: '탐색 인덱스 갱신 완료',
        message: `${Object.keys(index.items).length.toLocaleString('ko-KR')}개 항목을 인덱싱했습니다.`,
        progress: {
          current: 1,
          total: 1
        }
      })
      return index
    } catch (error) {
      if (error instanceof GraphResponseError && error.status === 410 && !didRestartFromFullEnumeration) {
        didRestartFromFullEnumeration = true
        emitGraphActivity({
          level: 'warning',
          scope: 'index',
          title: '인덱스 전체 재수집',
          message: 'OneDrive delta 토큰이 만료되어 전체 인덱스를 다시 수집합니다.',
          detail: error.message,
          status: error.status,
          retryAfterMs: error.retryAfterMs
        })
        Object.assign(index, createEmptyDriveIndex(rootItemId))
        activeDriveIndexSnapshot = index
        activeDriveIndexSnapshotAccountId = accountId
        deltaSequence = 0
        shouldRewriteDriveIndexStore = true
        await resetDriveDeltaStaging(accountId)
        url = parseStoredDriveDeltaUrl(error.location) ?? createDeltaUrl()
        continue
      }

      throw error
    }
  }
}

async function ensureRootItemId(index: DriveIndex, accountId: string | null): Promise<DriveIndex> {
  const accessToken = await getGraphAccessToken(accountId)
  const root = await graphGet<GraphDriveRoot>(createRootUrl(), accessToken)

  if (!root.id) {
    throw new Error('OneDrive 루트 폴더 정보를 확인하지 못했습니다.')
  }

  index.rootItemId = root.id
  activeDriveIndexSnapshot = index
  activeDriveIndexSnapshotAccountId = accountId
  await writeDriveIndex(index, accountId)
  return index
}

async function fetchAndMergeFolderChildren(index: DriveIndex, parentId: string): Promise<DriveIndex> {
  const accessToken = await getGraphAccessToken()
  const children: GraphDriveItem[] = []
  let url = createChildrenUrl(index, parentId)

  while (true) {
    const response = await graphGet<GraphChildrenResponse>(url, accessToken)
    children.push(...(response.value ?? []))

    if (!response['@odata.nextLink']) {
      break
    }

    url = parseGraphUrl(response['@odata.nextLink'], 'OneDrive 폴더 다음 페이지 주소가 올바르지 않습니다.')
  }

  const fetchedIds = new Set(children.flatMap((item) => (item.id ? [item.id] : [])))

  for (const child of Object.values(index.items)) {
    if (child.parentId === parentId && !fetchedIds.has(child.id)) {
      removeIndexedItem(index, child.id)
    }
  }

  for (const child of children) {
    if (!child.id) {
      continue
    }

    const previousItem = index.items[child.id]
    index.items[child.id] = {
      ...mapDriveItem(child, previousItem),
      parentId
    }
  }

  index.expandedFolderIds[parentId] = true
  activeDriveIndexSnapshot = index
  await writeDriveIndex(index)
  return index
}

async function ensureRemoteDriveFolder(
  index: DriveIndex,
  parentId: string,
  folderName: string,
  accessToken: string
): Promise<CloudDriveItem> {
  const normalizedFolderName = validateDriveItemName(folderName)
  const currentIndex = await fetchAndMergeFolderChildren(index, parentId)
  const existingFolder =
    (await findDriveIndexChildByName(await getActiveAccountId(), parentId, normalizedFolderName)) ??
    Object.values(currentIndex.items).find(
      (item) =>
        item.parentId === parentId &&
        item.type === 'folder' &&
        item.name.localeCompare(normalizedFolderName, 'ko-KR', { sensitivity: 'base' }) === 0
    )

  if (existingFolder?.type === 'folder') {
    return existingFolder
  }

  const graphItem = await graphSend<GraphDriveItem>(createChildrenCollectionUrl(currentIndex, parentId), accessToken, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      name: normalizedFolderName,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'rename'
    })
  })

  return mergeRemoteDriveItem(graphItem, parentId)
}

async function ensureRemoteDriveFolderRaw(
  rootItemId: string,
  parentId: string,
  folderName: string,
  accessToken: string
): Promise<CloudDriveItem> {
  const normalizedFolderName = validateDriveItemName(folderName)
  const existingChildren = await listRawDriveChildren(rootItemId, parentId, accessToken)
  const existingFolder = existingChildren.find(
    (item) => item.type === 'folder' && item.name.localeCompare(normalizedFolderName, 'ko-KR', { sensitivity: 'base' }) === 0
  )

  if (existingFolder) {
    return existingFolder
  }

  const graphItem = await graphSend<GraphDriveItem>(createChildrenCollectionUrlByRootId(rootItemId, parentId), accessToken, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      name: normalizedFolderName,
      folder: {},
      '@microsoft.graph.conflictBehavior': 'rename'
    })
  })

  return mapDriveItem(graphItem)
}

async function listRawDriveChildren(rootItemId: string, parentId: string, accessToken: string): Promise<CloudDriveItem[]> {
  const children: CloudDriveItem[] = []
  let url = createChildrenUrlByRootId(rootItemId, parentId)

  while (true) {
    const response = await graphGet<GraphChildrenResponse>(url, accessToken)

    children.push(...(response.value ?? []).flatMap((item) => (item.id ? [mapDriveItem(item)] : [])))

    if (!response['@odata.nextLink']) {
      break
    }

    url = parseGraphUrl(response['@odata.nextLink'], 'OneDrive 폴더 다음 페이지 주소가 올바르지 않습니다.')
  }

  return children
}

async function copyDriveItemToFolder(
  itemId: string,
  parentId: string,
  driveId: string,
  accessToken: string,
  name?: string
): Promise<boolean> {
  const response = await graphFetch(createCopyUrl(itemId), accessToken, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      parentReference: {
        driveId,
        id: parentId
      },
      ...(name ? { name } : {})
    })
  })

  if (response.status !== 202) {
    return true
  }

  const monitorUrl = response.headers.get('Location')

  if (!monitorUrl) {
    return false
  }

  return waitForGraphCopyOperation(parseTrustedCopyMonitorUrl(monitorUrl), accessToken)
}

async function waitForGraphCopyOperation(monitorUrl: URL, accessToken: string): Promise<boolean> {
  const deadline = Date.now() + COPY_OPERATION_MAX_WAIT_MS

  while (Date.now() < deadline) {
    await delay(COPY_OPERATION_POLL_INTERVAL_MS)

    const operation = await graphGet<GraphCopyOperationResponse>(monitorUrl, accessToken)
    const status = operation.status?.toLowerCase()

    if (status === 'completed') {
      return true
    }

    if (status === 'failed' || status === 'deletefailed') {
      throw new Error(operation.error?.message ?? operation.statusDescription ?? 'OneDrive 항목 복사에 실패했습니다.')
    }
  }

  return false
}

async function processAccountTransferTask(task: InternalDriveTransferTask, state: DriveTransferState, onProgress?: DriveTransferProgressListener): Promise<void> {
  if (runningTransferIds.has(task.id)) {
    return
  }

  if (!task.sourceAccountId || !task.targetAccountId || !task.sourceItemId || (!task.cleanupOnly && (!task.targetParentId || !task.tempPath))) {
    await failTransferTask(task, state, '계정 간 전송 정보를 찾지 못했습니다.', onProgress)
    return
  }

  runningTransferIds.add(task.id)
  transferPauseRequests.delete(task.id)
  const abortController = new AbortController()
  transferAbortControllers.set(task.id, abortController)
  const sourceSize = Math.max(task.sourceSize ?? 0, 0)
  let failureStage: TransferFailureStage = 'metadata'

  try {
    throwIfTransferPaused(task.id)

    if (task.cleanupOnly) {
      await processAccountTransferCleanupTask(task, state, onProgress, abortController.signal)
      return
    }

    task.status = 'running'
    task.message = '계정 간 다운로드 중'
    task.bytesPerSecond = 0
    task.nextRetryAt = undefined
    task.speedSampleAt = Date.now()
    task.speedSampleBytes = task.transferredBytes
    await writeDriveTransferState(state, onProgress)

    failureStage = 'auth'
    const sourceAccessToken = await getGraphAccessToken(task.sourceAccountId)
    const targetAccessToken = await getGraphAccessToken(task.targetAccountId)
    failureStage = 'download-stream'
    await downloadAccountTransferSource(task, sourceAccessToken, sourceSize, state, onProgress, abortController.signal)

    failureStage = 'upload-session'
    task.message = '계정 간 업로드 중'
    await writeDriveTransferState(state, onProgress)
    await uploadAccountTransferTarget(task, targetAccessToken, sourceSize, state, onProgress, abortController.signal)

    if (task.deleteSourceOnComplete && !task.sourceDeleted) {
      failureStage = 'finalize'
      await waitForDriveTransferThrottle(abortController.signal)
      await graphSend<void>(createItemUrl(task.sourceItemId), sourceAccessToken, {
        method: 'DELETE',
        signal: abortController.signal
      })
      task.sourceDeleted = true
    }

    if (task.tempPath) {
      await unlinkIfExists(task.tempPath)
    }
    task.status = 'completed'
    task.transferredBytes = task.totalBytes
    task.bytesPerSecond = 0
    task.uploadUrl = undefined
    task.expirationDateTime = undefined
    task.nextRetryAt = undefined
    task.lastError = undefined
    task.failureStage = undefined
    task.message = task.deleteSourceOnComplete ? '계정 간 이동 완료' : '계정 간 복사 완료'
    task.updatedAt = new Date().toISOString()
    await removeCompletedDriveTransferTask(task, onProgress)
  } catch (error) {
    if (transferDeleteRequests.has(task.id)) {
      return
    }

    if (isTransferPauseError(error, task.id)) {
      await pauseTransferTask(task, state, onProgress)
      throw new TransferPausedError()
    }

    await scheduleRetryTransferTask(
      task,
      state,
      getErrorMessage(error, '계정 간 전송을 완료하지 못했습니다.'),
      failureStage,
      onProgress,
      getRetryAfterDelayMs(error),
      isResponseThrottleError(error)
    )
  } finally {
    runningTransferIds.delete(task.id)
    transferAbortControllers.delete(task.id)
    transferPauseRequests.delete(task.id)
    transferDeleteRequests.delete(task.id)
  }
}

async function downloadAccountTransferSource(
  task: InternalDriveTransferTask,
  sourceAccessToken: string,
  sourceSize: number,
  state: DriveTransferState,
  onProgress: DriveTransferProgressListener | undefined,
  signal: AbortSignal
): Promise<void> {
  if (!task.sourceItemId || !task.tempPath) {
    throw new Error('계정 간 다운로드할 원본 정보를 찾지 못했습니다.')
  }

  await mkdir(dirname(task.tempPath), { recursive: true })
  let offset = await getExistingFileSize(task.tempPath)

  if (sourceSize > 0 && offset >= sourceSize) {
    task.downloadedBytes = sourceSize
    task.transferredBytes = (task.uploadedBytes ?? 0) + sourceSize
    await writeDriveTransferState(state, onProgress)
    return
  }

  let response = await createDownloadResponse(task.sourceItemId, offset, signal, sourceAccessToken)

  if (offset > 0 && response.status === 200) {
    offset = 0
  }

  if (!response.ok) {
    throw new Error('원본 계정에서 파일을 다운로드하지 못했습니다.')
  }

  if (!response.body) {
    throw new Error('원본 파일 내용이 비어 있습니다.')
  }

  const inferredSourceSize = inferDownloadTotalBytes(response, offset, sourceSize)
  task.sourceSize = Math.max(inferredSourceSize, sourceSize)
  task.totalBytes = task.sourceSize * 2
  task.downloadedBytes = offset
  task.transferredBytes = offset + (task.uploadedBytes ?? 0)
  await writeDriveTransferState(state, onProgress)

  const writeStream = createWriteStream(task.tempPath, {
    flags: offset > 0 && response.status === 206 ? 'a' : 'w'
  })
  let lastProgressSaveAt = 0

  try {
    for await (const chunk of Readable.fromWeb(response.body as unknown as Parameters<typeof Readable.fromWeb>[0])) {
      throwIfTransferPaused(task.id)

      if (!writeStream.write(chunk)) {
        await once(writeStream, 'drain')
      }

      task.downloadedBytes = (task.downloadedBytes ?? 0) + Buffer.byteLength(chunk)
      task.transferredBytes = (task.downloadedBytes ?? 0) + (task.uploadedBytes ?? 0)
      lastProgressSaveAt = await saveTransferProgress(task, state, lastProgressSaveAt, onProgress)
    }
  } finally {
    writeStream.end()
    await once(writeStream, 'finish')
  }

  task.downloadedBytes = Math.max(task.downloadedBytes ?? 0, task.sourceSize ?? sourceSize)
  task.transferredBytes = (task.downloadedBytes ?? 0) + (task.uploadedBytes ?? 0)
  await writeDriveTransferState(state, onProgress)
}

async function processAccountTransferCleanupTask(
  task: InternalDriveTransferTask,
  state: DriveTransferState,
  onProgress?: DriveTransferProgressListener,
  signal?: AbortSignal
): Promise<void> {
  if (!task.sourceAccountId || !task.sourceItemId || !task.transferGroupId) {
    throw new Error('원본 폴더 삭제 정보를 찾지 못했습니다.')
  }

  if (!(await isAccountTransferGroupCompleted(task.transferGroupId, task.expectedGroupTaskIds, task.completedGroupTaskIds))) {
    const retryDelayMs = TRANSFER_RETRY_BASE_DELAY_MS

    task.status = 'retrying'
    task.bytesPerSecond = 0
    task.nextRetryAt = new Date(Date.now() + retryDelayMs).toISOString()
    task.message = '하위 파일 전송 완료 대기 중'
    task.updatedAt = new Date().toISOString()
    await writeDriveTransferState(state, onProgress)
    scheduleDriveTransferRetryWorker(retryDelayMs)
    return
  }

  task.status = 'running'
  task.message = '원본 폴더 삭제 중'
  await writeDriveTransferState(state, onProgress)

  const sourceAccessToken = await getGraphAccessToken(task.sourceAccountId)
  await waitForDriveTransferThrottle(signal)
  await graphSend<void>(createItemUrl(task.sourceItemId), sourceAccessToken, {
    method: 'DELETE',
    signal
  })

  task.status = 'completed'
  task.message = '원본 폴더 삭제 완료'
  task.nextRetryAt = undefined
  task.lastError = undefined
  task.failureStage = undefined
  task.updatedAt = new Date().toISOString()
  await removeCompletedDriveTransferTask(task, onProgress)
}

async function isAccountTransferGroupCompleted(
  transferGroupId: string,
  expectedTaskIds: string[] = [],
  completedTaskIds: string[] = []
): Promise<boolean> {
  if (expectedTaskIds.length > 0) {
    const completedTaskIdSet = new Set(completedTaskIds)

    for (const taskId of expectedTaskIds) {
      if (completedTaskIdSet.has(taskId)) {
        continue
      }

      const task = await readDriveTransferTask(taskId)

      if (!task || task.status !== 'completed') {
        return false
      }
    }

    return true
  }

  const index = await readDriveTransferIndex()

  for (const taskId of index.taskIds) {
    const task = await readDriveTransferTask(taskId)

    if (task?.transferGroupId === transferGroupId && !task.cleanupOnly && task.status !== 'completed') {
      return false
    }
  }

  return true
}

async function uploadAccountTransferTarget(
  task: InternalDriveTransferTask,
  targetAccessToken: string,
  sourceSize: number,
  state: DriveTransferState,
  onProgress: DriveTransferProgressListener | undefined,
  signal: AbortSignal
): Promise<void> {
  if (!task.targetParentId || !task.tempPath || !task.targetAccountId) {
    throw new Error('계정 간 업로드할 대상 정보를 찾지 못했습니다.')
  }

  const fileSize = Math.max(task.sourceSize ?? sourceSize, 0)
  await waitForDriveTransferThrottle(signal)
  const targetRootId = await getRootItemId(targetAccessToken)
  const conflictBehavior = task.conflictBehavior ?? 'rename'

  if (fileSize === 0) {
    await waitForDriveTransferThrottle(signal)
    const graphItem = await graphSend<GraphDriveItem>(
      createUploadContentUrlByRootId(targetRootId, task.targetParentId, task.name, conflictBehavior),
      targetAccessToken,
      {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/octet-stream'
        },
        body: Buffer.alloc(0),
        signal
      }
    )

    task.resultItem = mapDriveItem(graphItem)
    task.uploadedBytes = 0
    task.transferredBytes = 0
    await writeDriveTransferState(state, onProgress)
    return
  }

  const fileStat = await stat(task.tempPath)

  if (!fileStat.isFile() || fileStat.size < fileSize) {
    throw new Error('계정 간 전송 임시 파일이 아직 완성되지 않았습니다.')
  }

  let offset = await ensureAccountTransferUploadSessionOffset(task, targetRootId, targetAccessToken, state, onProgress, signal)
  const fileHandle = await open(task.tempPath, 'r')
  let lastProgressSaveAt = 0

  try {
    while (offset < fileSize) {
      throwIfTransferPaused(task.id)

      if (!task.uploadUrl) {
        throw new Error('OneDrive 업로드 세션을 찾지 못했습니다.')
      }

      const chunkLength = Math.min(UPLOAD_CHUNK_SIZE_BYTES, fileSize - offset)
      const buffer = Buffer.allocUnsafe(chunkLength)
      const { bytesRead } = await fileHandle.read(buffer, 0, chunkLength, offset)

      if (bytesRead <= 0) {
        throw new Error('계정 간 전송 임시 파일을 읽지 못했습니다.')
      }

      await waitForDriveTransferThrottle(signal)
      const response = await fetch(task.uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Length': String(bytesRead),
          'Content-Range': `bytes ${offset}-${offset + bytesRead - 1}/${fileSize}`
        },
        body: bytesRead === buffer.length ? buffer : buffer.subarray(0, bytesRead),
        signal
      })

      if (response.status === 202) {
        const session = (await response.json()) as GraphUploadSessionResponse
        offset = parseUploadRangeStart(session.nextExpectedRanges?.[0]) ?? offset + bytesRead
        task.expirationDateTime = session.expirationDateTime ?? task.expirationDateTime
        task.uploadedBytes = Math.min(offset, fileSize)
        task.transferredBytes = (task.downloadedBytes ?? fileSize) + task.uploadedBytes
        lastProgressSaveAt = await saveTransferProgress(task, state, lastProgressSaveAt, onProgress)
        continue
      }

      if (response.ok) {
        const graphItem = (await response.json()) as GraphDriveItem
        task.resultItem = mapDriveItem(graphItem)
        task.uploadedBytes = fileSize
        task.transferredBytes = (task.downloadedBytes ?? fileSize) + fileSize
        task.uploadUrl = undefined
        task.expirationDateTime = undefined
        await writeDriveTransferState(state, onProgress)
        return
      }

      if (response.status === 404) {
        task.uploadUrl = undefined
        offset = await ensureAccountTransferUploadSessionOffset(task, targetRootId, targetAccessToken, state, onProgress, signal)
        continue
      }

      throw await createResponseError(response, 'OneDrive 파일 업로드에 실패했습니다.')
    }
  } finally {
    await fileHandle.close()
  }

  throw new Error('대상 계정 업로드가 완료되지 않았습니다.')
}

async function ensureAccountTransferUploadSessionOffset(
  task: InternalDriveTransferTask,
  targetRootId: string,
  targetAccessToken: string,
  state: DriveTransferState,
  onProgress?: DriveTransferProgressListener,
  signal?: AbortSignal
): Promise<number> {
  const fileSize = Math.max(task.sourceSize ?? 0, 0)

  if (task.uploadUrl && !isExpired(task.expirationDateTime)) {
    await waitForDriveTransferThrottle(signal)
    const response = await fetch(task.uploadUrl, { signal })

    if (response.ok) {
      const session = (await response.json()) as GraphUploadSessionResponse
      task.expirationDateTime = session.expirationDateTime ?? task.expirationDateTime
      task.uploadedBytes = Math.min(parseUploadRangeStart(session.nextExpectedRanges?.[0]) ?? task.uploadedBytes ?? 0, fileSize)
      task.transferredBytes = (task.downloadedBytes ?? fileSize) + task.uploadedBytes
      await writeDriveTransferState(state, onProgress)
      return task.uploadedBytes
    }

    if (response.status !== 404) {
      throw await createResponseError(response, '대상 계정 업로드에 실패했습니다.')
    }
  }

  if (!task.targetParentId) {
    throw new Error('대상 폴더 정보를 찾지 못했습니다.')
  }

  await waitForDriveTransferThrottle(signal)
  const uploadSession = await graphSend<GraphUploadSessionResponse>(
    createUploadSessionUrlByRootId(targetRootId, task.targetParentId, task.name),
    targetAccessToken,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        item: {
          '@microsoft.graph.conflictBehavior': task.conflictBehavior ?? 'rename',
          name: task.name
        }
      }),
      signal
    }
  )

  if (!uploadSession.uploadUrl) {
    throw new Error('대상 계정 업로드 세션을 만들지 못했습니다.')
  }

  task.uploadUrl = uploadSession.uploadUrl
  task.expirationDateTime = uploadSession.expirationDateTime
  task.uploadedBytes = 0
  task.transferredBytes = task.downloadedBytes ?? fileSize
  await writeDriveTransferState(state, onProgress)
  return 0
}

async function processUploadTask(
  task: InternalDriveTransferTask,
  state: DriveTransferState,
  onProgress?: DriveTransferProgressListener
): Promise<CloudDriveItem | null> {
  if (runningTransferIds.has(task.id)) {
    return task.resultItem ?? null
  }

  if (!task.localPath || !task.parentId) {
    await failTransferTask(task, state, '업로드할 파일 정보를 찾지 못했습니다.', onProgress)
    return null
  }

  runningTransferIds.add(task.id)
  transferPauseRequests.delete(task.id)
  const abortController = new AbortController()
  transferAbortControllers.set(task.id, abortController)
  let failureStage: TransferFailureStage = 'metadata'

  try {
    throwIfTransferPaused(task.id)
    failureStage = 'local-file'
    const fileStat = await stat(task.localPath)

    if (!fileStat.isFile() || fileStat.size !== task.totalBytes) {
      throw new Error('로컬 파일이 이동되었거나 크기가 변경되었습니다.')
    }

    task.status = 'running'
    task.message = '업로드 중'
    task.bytesPerSecond = 0
    task.nextRetryAt = undefined
    task.speedSampleAt = Date.now()
    task.speedSampleBytes = task.transferredBytes
    await writeDriveTransferState(state, onProgress)

    failureStage = 'metadata'
    await waitForDriveTransferThrottle(abortController.signal)
    const index = await ensureDriveIndexForListing(false)
    failureStage = 'auth'
    const accessToken = await getGraphAccessToken()

    if (task.totalBytes === 0) {
      failureStage = 'upload-chunk'
      await waitForDriveTransferThrottle(abortController.signal)
      const graphItem = await graphSend<GraphDriveItem>(createUploadContentUrl(index, task.parentId, task.name), accessToken, {
        method: 'PUT',
        headers: {
          'Content-Type': 'application/octet-stream'
        },
        body: Buffer.alloc(0)
      })
      const uploadedItem = await mergeRemoteDriveItem(graphItem, task.parentId)

      task.resultItem = uploadedItem
      task.status = 'completed'
      task.message = '업로드 완료'
      task.nextRetryAt = undefined
      task.lastError = undefined
      task.failureStage = undefined
      await removeCompletedDriveTransferTask(task, onProgress)
      return uploadedItem
    }

    failureStage = 'upload-session'
    let offset = await ensureUploadSessionOffset(task, index, accessToken, state, onProgress, abortController.signal)
    const fileHandle = await open(task.localPath, 'r')
    let lastProgressSaveAt = 0

    try {
      while (offset < task.totalBytes) {
        throwIfTransferPaused(task.id)

        if (!task.uploadUrl) {
          throw new Error('OneDrive 업로드 세션을 찾지 못했습니다.')
        }

        failureStage = 'local-file'
        const chunkLength = Math.min(UPLOAD_CHUNK_SIZE_BYTES, task.totalBytes - offset)
        const buffer = Buffer.allocUnsafe(chunkLength)
        const { bytesRead } = await fileHandle.read(buffer, 0, chunkLength, offset)

        if (bytesRead <= 0) {
          throw new Error('로컬 파일을 읽지 못했습니다.')
        }

        failureStage = 'upload-chunk'
        await waitForDriveTransferThrottle(abortController.signal)
        const response = await fetch(task.uploadUrl, {
          method: 'PUT',
          headers: {
            'Content-Length': String(bytesRead),
            'Content-Range': `bytes ${offset}-${offset + bytesRead - 1}/${task.totalBytes}`
          },
          body: bytesRead === buffer.length ? buffer : buffer.subarray(0, bytesRead),
          signal: abortController.signal
        })

        if (response.status === 202) {
          const session = (await response.json()) as GraphUploadSessionResponse
          offset = parseUploadRangeStart(session.nextExpectedRanges?.[0]) ?? offset + bytesRead
          task.expirationDateTime = session.expirationDateTime ?? task.expirationDateTime
          task.transferredBytes = Math.min(offset, task.totalBytes)
          lastProgressSaveAt = await saveTransferProgress(task, state, lastProgressSaveAt, onProgress)
          continue
        }

        if (response.ok) {
          const graphItem = (await response.json()) as GraphDriveItem
          const uploadedItem = await mergeRemoteDriveItem(graphItem, task.parentId)

          task.resultItem = uploadedItem
          task.status = 'completed'
          task.transferredBytes = task.totalBytes
          task.bytesPerSecond = 0
          task.message = '업로드 완료'
          task.uploadUrl = undefined
          task.expirationDateTime = undefined
          task.nextRetryAt = undefined
          task.lastError = undefined
          task.failureStage = undefined
          await removeCompletedDriveTransferTask(task, onProgress)
          return uploadedItem
        }

        if (response.status === 404) {
          task.uploadUrl = undefined
          failureStage = 'upload-session'
          offset = await ensureUploadSessionOffset(task, index, accessToken, state, onProgress, abortController.signal)
          continue
        }

        throw await createResponseError(response, 'OneDrive 파일 업로드에 실패했습니다.')
      }
    } finally {
      await fileHandle.close()
    }

    throw new Error('OneDrive 파일 업로드가 완료되지 않았습니다.')
  } catch (error) {
    if (transferDeleteRequests.has(task.id)) {
      return null
    }

    if (isTransferPauseError(error, task.id)) {
      await pauseTransferTask(task, state, onProgress)
      throw new TransferPausedError()
    }

    await scheduleRetryTransferTask(
      task,
      state,
      getErrorMessage(error, '업로드를 완료하지 못했습니다.'),
      failureStage,
      onProgress,
      getRetryAfterDelayMs(error),
      isResponseThrottleError(error)
    )
    return null
  } finally {
    runningTransferIds.delete(task.id)
    transferAbortControllers.delete(task.id)
    transferPauseRequests.delete(task.id)
    transferDeleteRequests.delete(task.id)
  }
}

async function ensureUploadSessionOffset(
  task: InternalDriveTransferTask,
  index: DriveIndex,
  accessToken: string,
  state: DriveTransferState,
  onProgress?: DriveTransferProgressListener,
  signal?: AbortSignal
): Promise<number> {
  if (task.uploadUrl && !isExpired(task.expirationDateTime)) {
    await waitForDriveTransferThrottle(signal)
    const response = await fetch(task.uploadUrl, { signal })

    if (response.ok) {
      const session = (await response.json()) as GraphUploadSessionResponse
      task.expirationDateTime = session.expirationDateTime ?? task.expirationDateTime
      task.transferredBytes = Math.min(parseUploadRangeStart(session.nextExpectedRanges?.[0]) ?? task.transferredBytes, task.totalBytes)
      await writeDriveTransferState(state, onProgress)
      return task.transferredBytes
    }

    if (response.status !== 404) {
      throw await createResponseError(response, 'OneDrive 파일 업로드에 실패했습니다.')
    }
  }

  await waitForDriveTransferThrottle(signal)
  const uploadSession = await graphSend<GraphUploadSessionResponse>(createUploadSessionUrl(index, task.parentId ?? '', task.name), accessToken, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      item: {
        '@microsoft.graph.conflictBehavior': 'rename',
        name: task.name
      }
    }),
    signal
  })

  if (!uploadSession.uploadUrl) {
    throw new Error('OneDrive 업로드 세션을 만들지 못했습니다.')
  }

  task.uploadUrl = uploadSession.uploadUrl
  task.expirationDateTime = uploadSession.expirationDateTime
  task.transferredBytes = 0
  await writeDriveTransferState(state, onProgress)
  return 0
}

async function processDownloadTask(
  task: InternalDriveTransferTask,
  state: DriveTransferState,
  onProgress?: DriveTransferProgressListener
): Promise<void> {
  if (runningTransferIds.has(task.id)) {
    return
  }

  if (!task.itemId || !task.localPath) {
    await failTransferTask(task, state, '다운로드할 파일 정보를 찾지 못했습니다.', onProgress)
    return
  }

  runningTransferIds.add(task.id)
  transferPauseRequests.delete(task.id)
  const abortController = new AbortController()
  transferAbortControllers.set(task.id, abortController)
  let failureStage: TransferFailureStage = 'metadata'

  try {
    throwIfTransferPaused(task.id)
    task.tempPath = task.tempPath ?? `${task.localPath}.onedrive-download`
    task.status = 'running'
    task.message = '다운로드 중'
    task.bytesPerSecond = 0
    task.nextRetryAt = undefined
    task.speedSampleAt = Date.now()
    task.speedSampleBytes = task.transferredBytes
    await writeDriveTransferState(state, onProgress)
    failureStage = 'local-file'
    await mkdir(dirname(task.localPath), { recursive: true })

    let offset = await getExistingFileSize(task.tempPath)

    if (offset > 0 && task.totalBytes > 0 && offset >= task.totalBytes) {
      failureStage = 'finalize'
      await rename(task.tempPath, task.localPath)
      task.status = 'completed'
      task.transferredBytes = task.totalBytes
      task.message = '다운로드 완료'
      task.nextRetryAt = undefined
      task.lastError = undefined
      task.failureStage = undefined
      await removeCompletedDriveTransferTask(task, onProgress)
      return
    }

    failureStage = 'download-link'
    let response = await createDownloadResponse(task.itemId, offset, abortController.signal)

    if (offset > 0 && response.status === 200) {
      offset = 0
    }

    if (!response.ok) {
      throw new Error('OneDrive 파일을 다운로드하지 못했습니다.')
    }

    if (!response.body) {
      throw new Error('다운로드할 파일 내용이 비어 있습니다.')
    }

    task.totalBytes = inferDownloadTotalBytes(response, offset, task.totalBytes)
    task.transferredBytes = offset
    await writeDriveTransferState(state, onProgress)

    const writeStream = createWriteStream(task.tempPath, {
      flags: offset > 0 && response.status === 206 ? 'a' : 'w'
    })
    let lastProgressSaveAt = 0

    try {
      failureStage = 'download-stream'
      for await (const chunk of Readable.fromWeb(response.body as unknown as Parameters<typeof Readable.fromWeb>[0])) {
        throwIfTransferPaused(task.id)

        if (!writeStream.write(chunk)) {
          await once(writeStream, 'drain')
        }

        task.transferredBytes += Buffer.byteLength(chunk)
        lastProgressSaveAt = await saveTransferProgress(task, state, lastProgressSaveAt, onProgress)
      }
    } finally {
      writeStream.end()
      await once(writeStream, 'finish')
    }

    failureStage = 'finalize'
    await rename(task.tempPath, task.localPath)
    task.status = 'completed'
    task.transferredBytes = task.totalBytes || task.transferredBytes
    task.bytesPerSecond = 0
    task.message = '다운로드 완료'
    task.nextRetryAt = undefined
    task.lastError = undefined
    task.failureStage = undefined
    await removeCompletedDriveTransferTask(task, onProgress)
  } catch (error) {
    if (transferDeleteRequests.has(task.id)) {
      return
    }

    if (isTransferPauseError(error, task.id)) {
      await pauseTransferTask(task, state, onProgress)
      throw new TransferPausedError()
    }

    await scheduleRetryTransferTask(
      task,
      state,
      getErrorMessage(error, '다운로드를 완료하지 못했습니다.'),
      failureStage,
      onProgress,
      getRetryAfterDelayMs(error),
      isResponseThrottleError(error)
    )
  } finally {
    runningTransferIds.delete(task.id)
    transferAbortControllers.delete(task.id)
    transferPauseRequests.delete(task.id)
    transferDeleteRequests.delete(task.id)
  }
}

async function mergeRemoteDriveItem(item: GraphDriveItem, fallbackParentId?: string): Promise<CloudDriveItem> {
  if (!item.id) {
    throw new Error('OneDrive 항목 정보를 확인하지 못했습니다.')
  }

  const index = await getCurrentDriveIndex()
  const previousItem = index.items[item.id]
  const nextItem = {
    ...mapDriveItem(item, previousItem),
    parentId: item.parentReference?.id ?? previousItem?.parentId ?? fallbackParentId
  }

  if (!nextItem.parentId) {
    throw new Error('OneDrive 상위 폴더 정보를 확인하지 못했습니다.')
  }

  index.items[nextItem.id] = nextItem
  index.expandedFolderIds[nextItem.parentId] = true
  activeDriveIndexSnapshot = index
  await writeDriveIndex(index)
  return nextItem
}

function createDriveDeltaStagePayloads(items: GraphDriveItem[], startSequence: number): DriveDeltaStagePayload[] {
  const payloads: DriveDeltaStagePayload[] = []
  let sequence = startSequence

  for (const item of items) {
    if (!item.id) {
      continue
    }

    payloads.push({
      itemId: item.id,
      sequence,
      payload: JSON.stringify(item),
      isTombstone: Boolean(item.deleted)
    })
    sequence += 1
  }

  return payloads
}

async function applyStagedDriveDelta(index: DriveIndex, accountId: string | null): Promise<DriveIndexDeltaChange[]> {
  const stagePayloads = await readLastOccurrenceDriveDeltaStaging(accountId)
  const changes: DriveIndexDeltaChange[] = []

  for (const stagePayload of stagePayloads) {
    const change = applyDeltaItem(index, JSON.parse(stagePayload.payload) as GraphDriveItem)

    if (change) {
      changes.push(change)
    }
  }

  return changes
}

function applyDeltaItem(index: DriveIndex, item: GraphDriveItem): DriveIndexDeltaChange | null {
  if (!item.id || item.id === index.rootItemId) {
    return null
  }

  if (item.deleted) {
    removeIndexedItem(index, item.id)
    return {
      itemId: item.id,
      deleted: true
    }
  }

  const previousItem = index.items[item.id]
  const nextItem = mapDriveItem(item, previousItem)

  if (!nextItem.parentId) {
    return null
  }

  index.items[nextItem.id] = nextItem

  return {
    itemId: nextItem.id,
    item: nextItem
  }
}

function removeIndexedItem(index: DriveIndex, itemId: string): void {
  delete index.items[itemId]

  for (const child of Object.values(index.items)) {
    if (child.parentId === itemId) {
      removeIndexedItem(index, child.id)
    }
  }
}

function isIndexedDescendant(index: DriveIndex, candidateChildId: string, candidateAncestorId: string): boolean {
  let currentItem = index.items[candidateChildId]
  const visitedItemIds = new Set<string>()

  while (currentItem?.parentId) {
    if (currentItem.parentId === candidateAncestorId) {
      return true
    }

    if (visitedItemIds.has(currentItem.parentId)) {
      return false
    }

    visitedItemIds.add(currentItem.parentId)
    currentItem = index.items[currentItem.parentId]
  }

  return false
}

function mapDriveItem(item: GraphDriveItem, previousItem?: CloudDriveItem): CloudDriveItem {
  return {
    id: item.id ?? previousItem?.id ?? '',
    name: item.name ?? previousItem?.name ?? '이름 없음',
    type: getItemType(item, previousItem),
    size: item.size ?? previousItem?.size ?? 0,
    lastModifiedDateTime: item.lastModifiedDateTime ?? previousItem?.lastModifiedDateTime,
    webUrl: item.webUrl ?? previousItem?.webUrl,
    parentId: item.parentReference?.id ?? previousItem?.parentId,
    childCount: item.folder?.childCount ?? previousItem?.childCount,
    mimeType: item.file?.mimeType ?? previousItem?.mimeType,
    quickXorHash: normalizeContentHash(item.file?.hashes?.quickXorHash) ?? previousItem?.quickXorHash,
    cTag: normalizeContentHash(item.cTag) ?? previousItem?.cTag,
    eTag: normalizeContentHash(item.eTag) ?? previousItem?.eTag
  }
}

function getItemType(item: GraphDriveItem, previousItem?: CloudDriveItem): CloudDriveItemType {
  if (item.folder) {
    return 'folder'
  }

  if (item.package) {
    return 'package'
  }

  if (item.file) {
    return 'file'
  }

  return previousItem?.type ?? 'file'
}

function sortDriveItems(items: CloudDriveItem[]): CloudDriveItem[] {
  return [...items].sort((left, right) => {
    const typeOrder = DRIVE_ITEM_TYPE_ORDER[left.type] - DRIVE_ITEM_TYPE_ORDER[right.type]

    if (typeOrder !== 0) {
      return typeOrder
    }

    return DRIVE_ITEM_NAME_COLLATOR.compare(left.name, right.name)
  })
}

function isIndexReadyAndFresh(index: DriveIndex): boolean {
  if (!isIndexUsable(index) || !index.syncedAt) {
    return false
  }

  return Date.now() - new Date(index.syncedAt).getTime() < INDEX_FRESH_MS
}

function isIndexUsable(index: DriveIndex): boolean {
  return Boolean(index.rootItemId && index.deltaLink)
}

function isFolderLocallyNavigable(index: DriveIndex, parentId: string): boolean {
  return isIndexUsable(index) || Boolean(index.expandedFolderIds[parentId])
}

function createDriveIndexStatus(index: DriveIndex, accountId: string | null): DriveIndexStatus {
  const items = Object.values(index.items)

  return {
    isReady: isIndexUsable(index),
    isFresh: isIndexReadyAndFresh(index),
    isSyncing: isDriveIndexSyncing(accountId),
    itemCount: items.length,
    folderCount: items.filter((item) => item.type === 'folder').length,
    fileCount: items.filter((item) => item.type === 'file').length,
    packageCount: items.filter((item) => item.type === 'package').length,
    syncedAt: index.syncedAt
  }
}

function isDriveIndexSyncing(accountId: string | null): boolean {
  return driveIndexSyncPromises.has(getDriveIndexSyncKey(accountId))
}

function getDriveIndexSyncKey(accountId: string | null): string {
  return accountId ?? '__active__'
}

async function normalizeDriveThumbnailRequest(request: DriveThumbnailRequest): Promise<InternalDriveThumbnailRequest> {
  const itemId = request.itemId?.trim()

  if (!itemId) {
    throw new Error('썸네일을 가져올 OneDrive 항목을 확인하지 못했습니다.')
  }

  return {
    accountId: request.accountId ?? (await getActiveAccountId()),
    itemId,
    cacheKey: request.cacheKey?.trim() || 'current',
    priority: request.priority === 'high' ? 'high' : 'normal',
    size: normalizeDriveThumbnailSize(request.size),
    cacheOnly: request.cacheOnly === true
  }
}

function normalizeDriveThumbnailSize(size: DriveThumbnailSize | undefined): DriveThumbnailSize {
  if (!size) {
    return DRIVE_THUMBNAIL_DEFAULT_SIZE
  }

  if (size === 'small' || size === 'medium' || size === 'large' || /^c\d+x\d+(?:_crop)?$/.test(size)) {
    return size
  }

  return DRIVE_THUMBNAIL_DEFAULT_SIZE
}

function createDriveThumbnailCacheKey(request: InternalDriveThumbnailRequest): string {
  return [request.accountId ?? 'active', request.itemId, request.size, request.cacheKey].join('\u001f')
}

function createDriveThumbnailPendingKey(request: InternalDriveThumbnailRequest, cacheKey: string): string {
  return `${cacheKey}\u001f${request.cacheOnly ? 'cache' : 'result'}`
}

function enqueueDriveThumbnailRequest(request: InternalDriveThumbnailRequest, cacheKey: string): Promise<DriveThumbnailResult> {
  const requestPromise = new Promise<DriveThumbnailResult>((resolve, reject) => {
    const entry: DriveThumbnailQueueEntry = {
      request,
      cacheKey,
      attemptCount: 0,
      resolve,
      reject
    }

    if (request.priority === 'high') {
      driveThumbnailQueue.unshift(entry)
    } else {
      driveThumbnailQueue.push(entry)
    }
  })

  scheduleDriveThumbnailQueueFlush()
  return requestPromise
}

function scheduleDriveThumbnailQueueFlush(delayMs = DRIVE_THUMBNAIL_BATCH_DELAY_MS): void {
  if (driveThumbnailQueueTimer) {
    return
  }

  const normalizedDelayMs = Math.max(DRIVE_THUMBNAIL_BATCH_DELAY_MS, delayMs, getDriveThumbnailThrottleDelayMs())

  driveThumbnailQueueTimer = setTimeout(() => {
    driveThumbnailQueueTimer = null
    void flushDriveThumbnailQueue()
  }, normalizedDelayMs)
}

async function flushDriveThumbnailQueue(): Promise<void> {
  if (driveThumbnailQueueFlushPromise) {
    return driveThumbnailQueueFlushPromise
  }

  driveThumbnailQueueFlushPromise = flushDriveThumbnailQueueOnce().finally(() => {
    driveThumbnailQueueFlushPromise = null

    if (driveThumbnailQueue.length > 0) {
      scheduleDriveThumbnailQueueFlush()
    }
  })

  return driveThumbnailQueueFlushPromise
}

async function flushDriveThumbnailQueueOnce(): Promise<void> {
  while (driveThumbnailQueue.length > 0) {
    const throttleDelayMs = getDriveThumbnailThrottleDelayMs()

    if (throttleDelayMs > 0) {
      scheduleDriveThumbnailQueueFlush(throttleDelayMs)
      return
    }

    await fetchDriveThumbnailQueueBatch(takeDriveThumbnailQueueBatch())

    const nextThrottleDelayMs = getDriveThumbnailThrottleDelayMs()

    if (nextThrottleDelayMs > 0) {
      scheduleDriveThumbnailQueueFlush(nextThrottleDelayMs)
      return
    }

    if (driveThumbnailQueue.length > 0) {
      await delay(driveThumbnailBatchStepDelayMs)
    }
  }
}

function takeDriveThumbnailQueueBatch(): DriveThumbnailQueueEntry[] {
  const firstEntry = driveThumbnailQueue.shift()

  if (!firstEntry) {
    return []
  }

  const groupKey = getDriveThumbnailQueueGroupKey(firstEntry)
  const batchEntries = [firstEntry]

  for (let index = 0; index < driveThumbnailQueue.length && batchEntries.length < driveThumbnailBatchLimit; ) {
    const entry = driveThumbnailQueue[index]

    if (canJoinDriveThumbnailQueueBatch(firstEntry, entry, groupKey)) {
      batchEntries.push(entry)
      driveThumbnailQueue.splice(index, 1)
      continue
    }

    index += 1
  }

  return batchEntries
}

function getDriveThumbnailQueueGroupKey(entry: DriveThumbnailQueueEntry): string {
  return `${entry.request.accountId ?? 'active'}\u001f${entry.request.size}`
}

function canJoinDriveThumbnailQueueBatch(
  firstEntry: DriveThumbnailQueueEntry,
  entry: DriveThumbnailQueueEntry,
  groupKey = getDriveThumbnailQueueGroupKey(firstEntry)
): boolean {
  if (getDriveThumbnailQueueGroupKey(entry) !== groupKey) {
    return false
  }

  if (firstEntry.request.priority === 'high' && entry.request.priority !== 'high') {
    return false
  }

  return true
}

async function fetchDriveThumbnailQueueBatch(entries: DriveThumbnailQueueEntry[]): Promise<void> {
  if (entries.length === 0) {
    return
  }

  try {
    const outcomes = await fetchDriveThumbnailBatch(entries.map((entry) => entry.request))
    const successfulEntries = entries
      .map((entry, index) => ({
        entry,
        outcome: outcomes[index]
      }))
      .filter(
        (value): value is { entry: DriveThumbnailQueueEntry; outcome: Extract<DriveThumbnailBatchOutcome, { status: 'fulfilled' }> } =>
          value.outcome?.status === 'fulfilled'
      )
    const failedEntries = entries
      .map((entry, index) => ({
        entry,
        outcome: outcomes[index]
      }))
      .filter(
        (value): value is { entry: DriveThumbnailQueueEntry; outcome: Extract<DriveThumbnailBatchOutcome, { status: 'rejected' }> } =>
          value.outcome?.status === 'rejected'
      )

    const materializedResults = await Promise.all(
      successfulEntries.map(async ({ entry, outcome }) => ({
        entry,
        result: await materializeDriveThumbnailResult(entry.cacheKey, outcome.result, {
          cacheOnly: entry.request.cacheOnly
        })
      }))
    )

    for (const { entry, result } of materializedResults) {
      if (!entry.request.cacheOnly) {
        cacheDriveThumbnailResult(entry.cacheKey, result)
      }

      entry.resolve(result)
    }

    if (successfulEntries.length > 0) {
      recordDriveThumbnailBatchSuccess(successfulEntries.length)
    }

    if (failedEntries.length > 0) {
      handleDriveThumbnailQueueBatchError(
        failedEntries.map(({ entry }) => entry),
        selectDriveThumbnailBatchError(failedEntries.map(({ outcome }) => outcome.reason))
      )
    }
  } catch (error) {
    handleDriveThumbnailQueueBatchError(entries, error)
  }
}

function handleDriveThumbnailQueueBatchError(entries: DriveThumbnailQueueEntry[], error: unknown): void {
  const retryableEntries = entries.filter((entry) => entry.attemptCount < DRIVE_THUMBNAIL_MAX_RETRIES)
  const retryDelayMs = getDriveThumbnailRetryDelayMs(retryableEntries, error)

  if (retryDelayMs !== undefined) {
    recordDriveThumbnailBatchThrottle()
    emitGraphActivity({
      level: 'warning',
      scope: 'thumbnail',
      title: '썸네일 요청 제한',
      message: `${entries.length.toLocaleString('ko-KR')}개 썸네일 요청이 제한되어 ${formatRetryDelay(retryDelayMs)} 후 재시도합니다.`,
      detail: getErrorMessage(error, 'OneDrive 썸네일 요청이 제한되었습니다.'),
      status: error instanceof GraphResponseError ? error.status : undefined,
      retryAfterMs: error instanceof GraphResponseError ? error.retryAfterMs : undefined
    })
    retryDriveThumbnailQueueEntries(retryableEntries, retryDelayMs)

    for (const entry of entries) {
      if (!retryableEntries.includes(entry)) {
        entry.reject(error)
      }
    }

    return
  }

  if (isDriveThumbnailRetryableError(error)) {
    recordDriveThumbnailBatchThrottle()
    throttleDriveThumbnails(getRetryAfterDelayMs(error) ?? DRIVE_THUMBNAIL_RETRY_BASE_DELAY_MS)
    emitGraphActivity({
      level: 'warning',
      scope: 'thumbnail',
      title: '썸네일 요청 보류',
      message: '썸네일 요청 제한이 반복되어 일부 항목은 기본 아이콘으로 표시합니다.',
      detail: getErrorMessage(error, 'OneDrive 썸네일 요청이 제한되었습니다.'),
      status: error instanceof GraphResponseError ? error.status : undefined,
      retryAfterMs: error instanceof GraphResponseError ? error.retryAfterMs : undefined
    })

    for (const entry of entries) {
      const result = createMissingDriveThumbnailResult(entry.request.itemId)

      if (!entry.request.cacheOnly) {
        cacheDriveThumbnailResult(entry.cacheKey, result)
      }

      entry.resolve(result)
    }

    return
  }

  for (const entry of entries) {
    entry.reject(error)
  }
}

function selectDriveThumbnailBatchError(errors: unknown[]): unknown {
  return (
    errors
      .filter((error): error is GraphResponseError => error instanceof GraphResponseError && error.retryAfterMs !== undefined)
      .sort((left, right) => (right.retryAfterMs ?? 0) - (left.retryAfterMs ?? 0))[0] ??
    errors.find((error) => error instanceof GraphResponseError) ??
    errors[0] ??
    new Error('OneDrive 썸네일 요청에 실패했습니다.')
  )
}

async function fetchDriveThumbnailBatch(requests: InternalDriveThumbnailRequest[]): Promise<DriveThumbnailBatchOutcome[]> {
  if (requests.length === 1) {
    const [request] = requests

    try {
      return [
        {
          status: 'fulfilled',
          result: await fetchDriveThumbnail(request)
        }
      ]
    } catch (error) {
      return [
        {
          status: 'rejected',
          reason: error
        }
      ]
    }
  }

  const [firstRequest] = requests
  const accessToken = await getGraphAccessToken(firstRequest.accountId)
  const response = await graphSend<GraphBatchThumbnailResponse>(createGraphBatchUrl(), accessToken, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      requests: requests.map((request, index) => ({
        id: String(index),
        method: 'GET',
        url: createThumbnailBatchRequestUrl(request.itemId, request.size)
      }))
    })
  })
  const itemResponses = new Map(response.responses?.map((itemResponse) => [itemResponse.id, itemResponse]) ?? [])

  return requests.map((request, index) => {
    const itemResponse = itemResponses.get(String(index))

    try {
      return {
        status: 'fulfilled',
        result: readDriveThumbnailBatchItemResult(request, itemResponse)
      }
    } catch (error) {
      return {
        status: 'rejected',
        reason: error
      }
    }
  })
}

async function fetchDriveThumbnail(request: InternalDriveThumbnailRequest): Promise<DriveThumbnailResult> {
  const accessToken = await getGraphAccessToken(request.accountId)

  try {
    const response = await graphGet<GraphThumbnailResponse>(createThumbnailUrl(request.itemId, request.size), accessToken)

    return readDriveThumbnailResult(request, response)
  } catch (error) {
    if (error instanceof GraphResponseError && (error.status === 400 || error.status === 404)) {
      return createMissingDriveThumbnailResult(request.itemId)
    }

    throw error
  }
}

function readDriveThumbnailBatchItemResult(
  request: InternalDriveThumbnailRequest,
  itemResponse: GraphBatchThumbnailItemResponse | undefined
): DriveThumbnailResult {
  if (!itemResponse || typeof itemResponse.status !== 'number') {
    return createMissingDriveThumbnailResult(request.itemId)
  }

  if (itemResponse.status === 400 || itemResponse.status === 404) {
    return createMissingDriveThumbnailResult(request.itemId)
  }

  if (itemResponse.status < 200 || itemResponse.status >= 300) {
    throw new GraphResponseError(
      formatDriveThumbnailBatchError(itemResponse),
      itemResponse.status,
      null,
      parseRetryAfterMsFromRecord(itemResponse.headers)
    )
  }

  return readDriveThumbnailResult(request, itemResponse.body as GraphThumbnailResponse)
}

function readDriveThumbnailResult(request: InternalDriveThumbnailRequest, response: GraphThumbnailResponse | undefined): DriveThumbnailResult {
  const thumbnailSet = response?.value?.[0]
  const thumbnail =
    getGraphThumbnail(thumbnailSet, request.size) ??
    getGraphThumbnail(thumbnailSet, 'medium') ??
    getGraphThumbnail(thumbnailSet, 'small') ??
    getGraphThumbnail(thumbnailSet, 'large')

  if (!thumbnail?.url) {
    return createMissingDriveThumbnailResult(request.itemId)
  }

  return {
    itemId: request.itemId,
    status: 'ready',
    url: thumbnail.url,
    width: thumbnail.width,
    height: thumbnail.height
  }
}

function getGraphThumbnail(thumbnailSet: GraphThumbnailSet | undefined, size: DriveThumbnailSize | 'medium' | 'small' | 'large'): GraphThumbnail | null {
  const value = thumbnailSet?.[size]

  if (!value || typeof value !== 'object') {
    return null
  }

  const thumbnail = value as GraphThumbnail

  return typeof thumbnail.url === 'string' && thumbnail.url.length > 0 ? thumbnail : null
}

function createMissingDriveThumbnailResult(itemId: string): DriveThumbnailResult {
  return {
    itemId,
    status: 'missing'
  }
}

function cacheDriveThumbnailResult(cacheKey: string, result: DriveThumbnailResult): void {
  driveThumbnailCache.set(cacheKey, {
    result,
    expiresAtMs: Date.now() + (result.status === 'ready' ? DRIVE_THUMBNAIL_URL_CACHE_MS : DRIVE_THUMBNAIL_MISSING_CACHE_MS)
  })
}

function startDriveThumbnailCacheWarmup(index: DriveIndex, accountId: string | null, generation = driveIndexGeneration): void {
  const warmKey = getDriveIndexSyncKey(accountId)
  const warmSignature = createDriveThumbnailWarmSignature(index)

  if (driveThumbnailWarmSignatures.get(warmKey) === warmSignature || driveThumbnailWarmPromises.has(warmKey)) {
    return
  }

  const candidates = collectDriveThumbnailWarmCandidates(index)

  if (candidates.length === 0) {
    driveThumbnailWarmSignatures.set(warmKey, warmSignature)
    return
  }

  emitGraphActivity({
    level: 'info',
    scope: 'thumbnail',
    title: '썸네일 캐시 준비 시작',
    message: `${candidates.length.toLocaleString('ko-KR')}개 이미지/동영상 썸네일을 백그라운드로 준비합니다.`,
    progress: {
      current: 0,
      total: candidates.length
    }
  })

  const warmPromise = warmDriveThumbnailCache(candidates, accountId, generation)
    .then(() => {
      if (generation === driveIndexGeneration) {
        driveThumbnailWarmSignatures.set(warmKey, warmSignature)
        emitGraphActivity({
          level: 'success',
          scope: 'thumbnail',
          title: '썸네일 캐시 준비 완료',
          message: `${candidates.length.toLocaleString('ko-KR')}개 후보 썸네일 처리를 마쳤습니다.`,
          progress: {
            current: candidates.length,
            total: candidates.length
          }
        })
      }
    })
    .catch((error) => {
      if (generation === driveIndexGeneration) {
        console.error('OneDrive thumbnail cache warmup failed:', error)
      }
    })
    .finally(() => {
      if (driveThumbnailWarmPromises.get(warmKey) === warmPromise) {
        driveThumbnailWarmPromises.delete(warmKey)
      }
    })

  driveThumbnailWarmPromises.set(warmKey, warmPromise)
}

async function warmDriveThumbnailCache(items: CloudDriveItem[], accountId: string | null, generation: number): Promise<void> {
  for (let index = 0; index < items.length; index += DRIVE_THUMBNAIL_WARM_BATCH_SIZE) {
    if (generation !== driveIndexGeneration) {
      return
    }

    const batchItems = items.slice(index, index + DRIVE_THUMBNAIL_WARM_BATCH_SIZE)

    await Promise.allSettled(
      batchItems.map((item) =>
        getDriveItemThumbnail({
          accountId,
          itemId: item.id,
          cacheKey: createDriveThumbnailItemCacheKey(item),
          priority: 'normal',
          size: DRIVE_THUMBNAIL_DEFAULT_SIZE,
          cacheOnly: true
        })
      )
    )

    emitGraphActivity({
      level: 'info',
      scope: 'thumbnail',
      title: '썸네일 캐시 준비 중',
      message: `${Math.min(index + batchItems.length, items.length).toLocaleString('ko-KR')} / ${items.length.toLocaleString(
        'ko-KR'
      )}개 후보를 처리했습니다.`,
      progress: {
        current: Math.min(index + batchItems.length, items.length),
        total: items.length
      }
    })

    if (index + DRIVE_THUMBNAIL_WARM_BATCH_SIZE < items.length) {
      await delay(DRIVE_THUMBNAIL_WARM_BATCH_DELAY_MS)
    }
  }
}

function collectDriveThumbnailWarmCandidates(index: DriveIndex): CloudDriveItem[] {
  return Object.values(index.items).filter(isDriveThumbnailWarmCandidate)
}

function isDriveThumbnailWarmCandidate(item: CloudDriveItem): boolean {
  if (item.type !== 'file') {
    return false
  }

  const mimeType = item.mimeType?.toLocaleLowerCase('en-US') ?? ''

  if (mimeType.startsWith('image/') || mimeType.startsWith('video/')) {
    return true
  }

  const extension = getDriveItemFileExtension(item.name)

  return DRIVE_THUMBNAIL_IMAGE_EXTENSIONS.has(extension) || DRIVE_THUMBNAIL_VIDEO_EXTENSIONS.has(extension)
}

function createDriveThumbnailItemCacheKey(item: CloudDriveItem): string {
  return item.cTag ?? item.eTag ?? `${item.size}:${item.lastModifiedDateTime ?? ''}`
}

function createDriveThumbnailWarmSignature(index: DriveIndex): string {
  return `${index.syncedAt ?? ''}:${Object.keys(index.items).length}`
}

function getDriveItemFileExtension(name: string): string {
  const extension = extname(name).toLocaleLowerCase('en-US')

  return extension.startsWith('.') ? extension.slice(1) : extension
}

async function readCachedDriveThumbnailResult(cacheKey: string, itemId: string): Promise<DriveThumbnailResult | null> {
  try {
    const { dataPath, metadataPath } = getDriveThumbnailCachePaths(cacheKey)
    const metadata = normalizeDriveThumbnailDiskCacheEntry(JSON.parse(await readFile(metadataPath, 'utf8')), itemId)

    if (!metadata) {
      return null
    }

    if (metadata.expiresAtMs <= Date.now()) {
      await removeDriveThumbnailDiskCache(cacheKey)
      return null
    }

    const bytes = await readFile(dataPath)

    if (bytes.byteLength <= 0 || bytes.byteLength > DRIVE_THUMBNAIL_MAX_CACHE_BYTES) {
      return null
    }

    return {
      itemId: metadata.itemId,
      status: 'ready',
      url: createDriveThumbnailDataUrl(metadata.contentType, bytes),
      width: metadata.width,
      height: metadata.height
    }
  } catch {
    return null
  }
}

async function hasFreshDriveThumbnailDiskCache(cacheKey: string, itemId: string): Promise<boolean> {
  try {
    const { dataPath, metadataPath } = getDriveThumbnailCachePaths(cacheKey)
    const metadata = normalizeDriveThumbnailDiskCacheEntry(JSON.parse(await readFile(metadataPath, 'utf8')), itemId)

    if (!metadata) {
      return false
    }

    if (metadata.expiresAtMs <= Date.now()) {
      await removeDriveThumbnailDiskCache(cacheKey)
      return false
    }

    const dataStats = await stat(dataPath)

    return dataStats.isFile() && dataStats.size > 0 && dataStats.size <= DRIVE_THUMBNAIL_MAX_CACHE_BYTES
  } catch {
    return false
  }
}

async function materializeDriveThumbnailResult(
  cacheKey: string,
  result: DriveThumbnailResult,
  options: { cacheOnly?: boolean } = {}
): Promise<DriveThumbnailResult> {
  if (result.status !== 'ready' || !result.url || result.url.startsWith('data:')) {
    return result
  }

  try {
    const response = await fetch(result.url)

    if (!response.ok) {
      return result
    }

    const contentType = normalizeDriveThumbnailContentType(response.headers.get('content-type'))
    const arrayBuffer = await response.arrayBuffer()

    if (arrayBuffer.byteLength <= 0 || arrayBuffer.byteLength > DRIVE_THUMBNAIL_MAX_CACHE_BYTES) {
      return result
    }

    const bytes = Buffer.from(arrayBuffer)
    const cachedResult: DriveThumbnailResult = options.cacheOnly
      ? {
          itemId: result.itemId,
          status: 'ready',
          width: result.width,
          height: result.height
        }
      : {
          ...result,
          url: createDriveThumbnailDataUrl(contentType, bytes)
        }

    try {
      await writeDriveThumbnailDiskCache(cacheKey, cachedResult, contentType, bytes)
    } catch {
      // A disk cache miss should not prevent the thumbnail from rendering.
    }

    return cachedResult
  } catch {
    return result
  }
}

async function writeDriveThumbnailDiskCache(
  cacheKey: string,
  result: DriveThumbnailResult,
  contentType: string,
  bytes: Buffer
): Promise<void> {
  const { dataPath, metadataPath } = getDriveThumbnailCachePaths(cacheKey)
  const temporaryDataPath = `${dataPath}.${process.pid}.${randomUUID()}.tmp`
  const now = Date.now()
  const metadata: DriveThumbnailDiskCacheEntry = {
    version: 1,
    itemId: result.itemId,
    contentType,
    width: result.width,
    height: result.height,
    createdAtMs: now,
    expiresAtMs: now + DRIVE_THUMBNAIL_DISK_CACHE_MS
  }

  await mkdir(dirname(dataPath), { recursive: true })
  await writeFile(temporaryDataPath, bytes)
  await rename(temporaryDataPath, dataPath)
  await writeJsonAtomic(metadataPath, metadata)
}

async function removeDriveThumbnailDiskCache(cacheKey: string): Promise<void> {
  const { dataPath, metadataPath } = getDriveThumbnailCachePaths(cacheKey)

  await Promise.all([unlinkIfExists(dataPath), unlinkIfExists(metadataPath)])
}

function normalizeDriveThumbnailDiskCacheEntry(value: unknown, itemId: string): DriveThumbnailDiskCacheEntry | null {
  const metadata = value as Partial<DriveThumbnailDiskCacheEntry>

  if (
    metadata.version !== 1 ||
    metadata.itemId !== itemId ||
    typeof metadata.contentType !== 'string' ||
    typeof metadata.createdAtMs !== 'number' ||
    !Number.isFinite(metadata.createdAtMs) ||
    typeof metadata.expiresAtMs !== 'number' ||
    !Number.isFinite(metadata.expiresAtMs)
  ) {
    return null
  }

  return {
    version: 1,
    itemId: metadata.itemId,
    contentType: normalizeDriveThumbnailContentType(metadata.contentType),
    width: typeof metadata.width === 'number' && Number.isFinite(metadata.width) ? metadata.width : undefined,
    height: typeof metadata.height === 'number' && Number.isFinite(metadata.height) ? metadata.height : undefined,
    createdAtMs: metadata.createdAtMs,
    expiresAtMs: metadata.expiresAtMs
  }
}

function createDriveThumbnailDataUrl(contentType: string, bytes: Buffer): string {
  return `data:${normalizeDriveThumbnailContentType(contentType)};base64,${bytes.toString('base64')}`
}

function normalizeDriveThumbnailContentType(contentType: string | null): string {
  const normalizedContentType = contentType?.split(';', 1)[0]?.trim().toLocaleLowerCase('en-US')

  return normalizedContentType?.startsWith('image/') ? normalizedContentType : DRIVE_THUMBNAIL_DEFAULT_CONTENT_TYPE
}

function getDriveThumbnailRetryDelayMs(entries: DriveThumbnailQueueEntry[], error: unknown): number | undefined {
  if (entries.length === 0 || !isDriveThumbnailRetryableError(error)) {
    return undefined
  }

  const nextAttemptCount = Math.max(...entries.map((entry) => entry.attemptCount + 1))
  const exponentialDelayMs = DRIVE_THUMBNAIL_RETRY_BASE_DELAY_MS * 2 ** Math.max(0, nextAttemptCount - 1)
  const cappedDelayMs = Math.min(exponentialDelayMs, DRIVE_THUMBNAIL_RETRY_MAX_DELAY_MS)
  const jitterMs = Math.floor(Math.random() * Math.min(5000, cappedDelayMs / 3))

  return Math.max(cappedDelayMs + jitterMs, getRetryAfterDelayMs(error) ?? 0)
}

function retryDriveThumbnailQueueEntries(entries: DriveThumbnailQueueEntry[], delayMs: number): void {
  for (const entry of entries) {
    entry.attemptCount += 1
  }

  requeueDriveThumbnailEntries(entries)
  throttleDriveThumbnails(delayMs)
  scheduleDriveThumbnailQueueFlush(delayMs)
}

function requeueDriveThumbnailEntries(entries: DriveThumbnailQueueEntry[]): void {
  if (entries.length === 0) {
    return
  }

  const firstNormalPriorityIndex = driveThumbnailQueue.findIndex((entry) => entry.request.priority !== 'high')

  if (firstNormalPriorityIndex < 0) {
    driveThumbnailQueue.push(...entries)
  } else {
    driveThumbnailQueue.splice(firstNormalPriorityIndex, 0, ...entries)
  }
}

function recordDriveThumbnailBatchSuccess(successfulItemCount: number): void {
  driveThumbnailSuccessfulItemCount += successfulItemCount

  if (driveThumbnailSuccessfulItemCount < DRIVE_THUMBNAIL_ADAPTIVE_SUCCESS_TARGET) {
    return
  }

  driveThumbnailSuccessfulItemCount = 0
  driveThumbnailBatchLimit = Math.min(DRIVE_THUMBNAIL_BATCH_LIMIT_MAX, driveThumbnailBatchLimit + 1)
  driveThumbnailBatchStepDelayMs = Math.max(
    DRIVE_THUMBNAIL_BATCH_STEP_DELAY_MIN_MS,
    Math.floor(driveThumbnailBatchStepDelayMs * 0.8)
  )
}

function recordDriveThumbnailBatchThrottle(): void {
  driveThumbnailSuccessfulItemCount = 0
  driveThumbnailBatchLimit = Math.max(DRIVE_THUMBNAIL_BATCH_LIMIT_MIN, Math.floor(driveThumbnailBatchLimit / 2))
  driveThumbnailBatchStepDelayMs = Math.min(
    DRIVE_THUMBNAIL_BATCH_STEP_DELAY_MAX_MS,
    Math.max(DRIVE_THUMBNAIL_BATCH_STEP_DELAY_INITIAL_MS, driveThumbnailBatchStepDelayMs * 2)
  )
}

function resetDriveThumbnailAdaptiveRate(): void {
  driveThumbnailBatchLimit = DRIVE_THUMBNAIL_BATCH_LIMIT_INITIAL
  driveThumbnailBatchStepDelayMs = DRIVE_THUMBNAIL_BATCH_STEP_DELAY_INITIAL_MS
  driveThumbnailSuccessfulItemCount = 0
}

function throttleDriveThumbnails(delayMs: number): void {
  const now = Date.now()

  driveThumbnailThrottleUntil = Math.max(driveThumbnailThrottleUntil, now + Math.max(delayMs, DRIVE_THUMBNAIL_BATCH_DELAY_MS))
}

function getDriveThumbnailThrottleDelayMs(): number {
  return Math.max(0, driveThumbnailThrottleUntil - Date.now())
}

function isDriveThumbnailRetryableError(error: unknown): boolean {
  return (
    error instanceof GraphResponseError &&
    (error.status === 429 || error.retryAfterMs !== undefined || (error.status >= 500 && error.status < 600))
  )
}

function formatDriveThumbnailBatchError(response: GraphBatchThumbnailItemResponse): string {
  const body = response.body

  if (body && typeof body === 'object') {
    const errorBody = body as GraphErrorResponse

    if (typeof errorBody.error?.message === 'string' && errorBody.error.message.length > 0) {
      return errorBody.error.message
    }
  }

  return response.status === 429 ? 'OneDrive 썸네일 요청이 일시적으로 제한되었습니다.' : 'OneDrive 썸네일을 가져오지 못했습니다.'
}

function parseRetryAfterMsFromRecord(headers: Record<string, string> | undefined): number | undefined {
  if (!headers) {
    return undefined
  }

  const retryAfterMs = parsePositiveDelayMs(headers['x-ms-retry-after-ms'] ?? headers['X-Ms-Retry-After-Ms'] ?? null)

  if (retryAfterMs !== undefined) {
    return retryAfterMs
  }

  const retryAfter = headers['Retry-After'] ?? headers['retry-after']

  return parseRetryAfterMsFromValue(retryAfter ?? null)
}

function createRootUrl(): URL {
  const url = new URL(`${GRAPH_BASE_URL}/me/drive/root`)

  url.searchParams.set('$select', 'id')
  return url
}

function createDriveUrl(): URL {
  const url = new URL(`${GRAPH_BASE_URL}/me/drive`)

  url.searchParams.set('$select', 'id')
  return url
}

function createDriveUsageUrl(): URL {
  const url = new URL(`${GRAPH_BASE_URL}/me/drive`)

  url.searchParams.set('$select', 'id,quota')
  return url
}

function createGraphBatchUrl(): URL {
  return new URL(`${GRAPH_BASE_URL}/$batch`)
}

function createItemUrl(itemId: string, withSelect = false): URL {
  const url = new URL(`${GRAPH_BASE_URL}/me/drive/items/${encodeURIComponent(itemId)}`)

  if (withSelect) {
    url.searchParams.set('$select', DRIVE_ITEM_SELECT)
  }

  return url
}

function createDownloadContentUrl(itemId: string): URL {
  return new URL(`${GRAPH_BASE_URL}/me/drive/items/${encodeURIComponent(itemId)}/content`)
}

function createThumbnailUrl(itemId: string, size: DriveThumbnailSize): URL {
  const url = new URL(`${GRAPH_BASE_URL}/me/drive/items/${encodeURIComponent(itemId)}/thumbnails`)

  url.searchParams.set('select', size)
  return url
}

function createThumbnailBatchRequestUrl(itemId: string, size: DriveThumbnailSize): string {
  const searchParams = new URLSearchParams()

  searchParams.set('select', size)
  return `/me/drive/items/${encodeURIComponent(itemId)}/thumbnails?${searchParams.toString()}`
}

function createCopyUrl(itemId: string): URL {
  return new URL(`${GRAPH_BASE_URL}/me/drive/items/${encodeURIComponent(itemId)}/copy`)
}

function createDeltaUrl(): URL {
  const url = new URL(`${GRAPH_BASE_URL}/me/drive/root/delta`)

  url.searchParams.set('$select', DRIVE_ITEM_SELECT)
  url.searchParams.set('$top', '200')
  return url
}

function createChildrenUrl(index: DriveIndex, parentId: string): URL {
  const url = createChildrenCollectionUrl(index, parentId)

  url.searchParams.set('$select', DRIVE_ITEM_SELECT)
  url.searchParams.set('$top', '200')
  return url
}

function createChildrenUrlByRootId(rootItemId: string, parentId: string): URL {
  const url = createChildrenCollectionUrlByRootId(rootItemId, parentId)

  url.searchParams.set('$select', DRIVE_ITEM_SELECT)
  url.searchParams.set('$top', '200')
  return url
}

function createChildrenUrlForItem(itemId: string): URL {
  const url = new URL(`${GRAPH_BASE_URL}/me/drive/items/${encodeURIComponent(itemId)}/children`)

  url.searchParams.set('$select', DRIVE_ITEM_SELECT)
  url.searchParams.set('$top', '200')
  return url
}

function createChildrenCollectionUrl(index: DriveIndex, parentId: string): URL {
  const path =
    parentId === index.rootItemId
      ? '/me/drive/root/children'
      : `/me/drive/items/${encodeURIComponent(parentId)}/children`

  return new URL(`${GRAPH_BASE_URL}${path}`)
}

function createChildrenCollectionUrlByRootId(rootItemId: string, parentId: string): URL {
  const path =
    parentId === rootItemId
      ? '/me/drive/root/children'
      : `/me/drive/items/${encodeURIComponent(parentId)}/children`

  return new URL(`${GRAPH_BASE_URL}${path}`)
}

function createUploadSessionUrl(index: DriveIndex, parentId: string, fileName: string): URL {
  const path =
    parentId === index.rootItemId
      ? `/me/drive/root:/${encodeDrivePathSegment(fileName)}:/createUploadSession`
      : `/me/drive/items/${encodeURIComponent(parentId)}:/${encodeDrivePathSegment(fileName)}:/createUploadSession`

  return new URL(`${GRAPH_BASE_URL}${path}`)
}

function createUploadSessionUrlByRootId(rootItemId: string, parentId: string, fileName: string): URL {
  const path =
    parentId === rootItemId
      ? `/me/drive/root:/${encodeDrivePathSegment(fileName)}:/createUploadSession`
      : `/me/drive/items/${encodeURIComponent(parentId)}:/${encodeDrivePathSegment(fileName)}:/createUploadSession`

  return new URL(`${GRAPH_BASE_URL}${path}`)
}

function createUploadContentUrl(index: DriveIndex, parentId: string, fileName: string): URL {
  const path =
    parentId === index.rootItemId
      ? `/me/drive/root:/${encodeDrivePathSegment(fileName)}:/content`
      : `/me/drive/items/${encodeURIComponent(parentId)}:/${encodeDrivePathSegment(fileName)}:/content`
  const url = new URL(`${GRAPH_BASE_URL}${path}`)

  url.searchParams.set('@microsoft.graph.conflictBehavior', 'rename')
  return url
}

function createUploadContentUrlByRootId(
  rootItemId: string,
  parentId: string,
  fileName: string,
  conflictBehavior: GraphConflictBehavior = 'rename'
): URL {
  const path =
    parentId === rootItemId
      ? `/me/drive/root:/${encodeDrivePathSegment(fileName)}:/content`
      : `/me/drive/items/${encodeURIComponent(parentId)}:/${encodeDrivePathSegment(fileName)}:/content`
  const url = new URL(`${GRAPH_BASE_URL}${path}`)

  url.searchParams.set('@microsoft.graph.conflictBehavior', conflictBehavior)
  return url
}

async function graphGet<T>(url: URL, accessToken: string, extraHeaders: Record<string, string> = {}): Promise<T> {
  const response = await graphFetch(url, accessToken, {
    headers: extraHeaders
  })

  return readJsonResponse<T>(response, 'OneDrive 응답을 해석하지 못했습니다.')
}

async function graphSend<T>(url: URL, accessToken: string, init: GraphRequestInit): Promise<T> {
  const response = await graphFetch(url, accessToken, init)

  if (response.status === 204) {
    return undefined as T
  }

  return readJsonResponse<T>(response, 'OneDrive 응답을 해석하지 못했습니다.')
}

async function graphFetch(url: URL, accessToken: string, init: GraphRequestInit): Promise<Response> {
  const response = await fetch(url, {
    ...init,
    headers: {
      Accept: 'application/json',
      Authorization: `Bearer ${accessToken}`,
      ...init.headers
    }
  })

  if (!response.ok) {
    const error = await createGraphResponseError(response)
    emitGraphResponseError(error, url)
    throw error
  }

  return response
}

async function readJsonResponse<T>(response: Response, fallbackMessage: string): Promise<T> {
  const body = await response.text()
  const trimmedBody = body.trim()

  if (!trimmedBody) {
    return undefined as T
  }

  try {
    return JSON.parse(trimmedBody) as T
  } catch (error) {
    throw new GraphResponseError(
      formatInvalidJsonResponseError(response, body, error, fallbackMessage),
      response.status,
      response.headers.get('Location'),
      parseRetryAfterMs(response.headers)
    )
  }
}

async function getDriveId(accessToken: string): Promise<string> {
  const drive = await graphGet<GraphDrive>(createDriveUrl(), accessToken)

  if (!drive.id) {
    throw new Error('OneDrive 드라이브 정보를 확인하지 못했습니다.')
  }

  return drive.id
}

async function getRootItemId(accessToken: string): Promise<string> {
  const root = await graphGet<GraphDriveRoot>(createRootUrl(), accessToken)

  if (!root.id) {
    throw new Error('OneDrive 루트 폴더 정보를 확인하지 못했습니다.')
  }

  return root.id
}

async function createGraphResponseError(response: Response): Promise<GraphResponseError> {
  return new GraphResponseError(
    await formatGraphError(response),
    response.status,
    response.headers.get('Location'),
    parseRetryAfterMs(response.headers)
  )
}

function emitGraphResponseError(error: GraphResponseError, url: URL): void {
  const isThrottle = error.status === 429 || error.retryAfterMs !== undefined

  emitGraphActivity({
    level: isThrottle ? 'warning' : 'error',
    scope: 'graph',
    title: isThrottle ? 'OneDrive 요청 제한' : 'OneDrive 요청 실패',
    message: isThrottle
      ? `${formatRetryDelay(getRetryAfterDelayMs(error) ?? DEFAULT_THROTTLE_RETRY_DELAY_MS)} 후 다시 시도합니다.`
      : error.message,
    detail: `${url.pathname}${url.search ? ` ${redactGraphUrl(url.toString())}` : ''}\n${error.message}`,
    status: error.status,
    retryAfterMs: error.retryAfterMs
  })
}

async function createResponseError(response: Response, fallbackMessage: string): Promise<GraphResponseError> {
  return new GraphResponseError(
    await formatGenericResponseError(response, fallbackMessage),
    response.status,
    response.headers.get('Location'),
    parseRetryAfterMs(response.headers)
  )
}

async function formatGraphError(response: Response): Promise<string> {
  const graphError = parseGraphError(await response.text())

  if (response.status === 401) {
    return '인증이 만료되었습니다. 다시 로그인하세요.'
  }

  if (response.status === 403) {
    return 'OneDrive 파일을 읽을 권한이 없습니다.'
  }

  if (response.status === 410) {
    return 'OneDrive 인덱스가 만료되어 다시 구성해야 합니다.'
  }

  if (response.status === 429) {
    return 'OneDrive 요청이 일시적으로 제한되었습니다.'
  }

  if (response.status >= 500) {
    return 'Microsoft Graph 서비스가 응답하지 않습니다. 잠시 후 다시 시도하세요.'
  }

  return graphError?.error?.message ?? 'OneDrive 파일 목록을 가져오지 못했습니다.'
}

async function formatGenericResponseError(response: Response, fallbackMessage: string): Promise<string> {
  if (response.status === 429) {
    return 'OneDrive 요청이 일시적으로 제한되었습니다.'
  }

  if (response.status >= 500) {
    return 'OneDrive 서비스가 응답하지 않습니다. 잠시 후 다시 시도하세요.'
  }

  return parseGraphError(await response.text())?.error?.message ?? fallbackMessage
}

function parseGraphError(body: string): GraphErrorResponse | null {
  try {
    return JSON.parse(body) as GraphErrorResponse
  } catch {
    return null
  }
}

function formatInvalidJsonResponseError(response: Response, body: string, error: unknown, fallbackMessage: string): string {
  const details = [
    `status=${response.status}`,
    `content-type=${response.headers.get('content-type') ?? 'unknown'}`,
    `url=${redactGraphUrl(response.url)}`,
    `body=${formatBodySnippet(body)}`
  ].join(', ')
  const parseMessage = error instanceof Error ? error.message : fallbackMessage

  return `Microsoft Graph 응답이 JSON 형식이 아닙니다. ${parseMessage}. ${details}`
}

function formatBodySnippet(body: string): string {
  const snippet = body.trim().replace(/\s+/g, ' ').slice(0, 500)

  return JSON.stringify(snippet || '<empty>')
}

function redactGraphUrl(value: string): string {
  try {
    const url = new URL(value)

    for (const key of [...url.searchParams.keys()]) {
      const normalizedKey = key.toLocaleLowerCase('en-US')

      if (normalizedKey.includes('token')) {
        url.searchParams.set(key, '<redacted>')
      }
    }

    return url.toString()
  } catch {
    return value
  }
}

function parseRetryAfterMs(headers: Headers): number | undefined {
  const retryAfterMs = parsePositiveDelayMs(headers.get('x-ms-retry-after-ms'))

  if (retryAfterMs !== undefined) {
    return retryAfterMs
  }

  return parseRetryAfterMsFromValue(headers.get('Retry-After'))
}

function parseRetryAfterMsFromValue(retryAfter: string | null): number | undefined {
  const retryAfterSeconds = Number(retryAfter)

  if (Number.isFinite(retryAfterSeconds) && retryAfterSeconds > 0) {
    return Math.ceil(retryAfterSeconds * 1000)
  }

  if (retryAfter) {
    const retryAt = Date.parse(retryAfter)

    if (Number.isFinite(retryAt)) {
      return Math.max(TRANSFER_DISPATCH_DELAY_MS, retryAt - Date.now())
    }
  }

  return undefined
}

function parsePositiveDelayMs(value: string | null): number | undefined {
  if (!value) {
    return undefined
  }

  const delayMs = Number(value)

  if (!Number.isFinite(delayMs) || delayMs <= 0) {
    return undefined
  }

  return Math.ceil(delayMs)
}

async function createDownloadResponse(itemId: string, offset: number, signal?: AbortSignal, accessToken?: string): Promise<Response> {
  const token = accessToken ?? (await getGraphAccessToken())
  const rangeHeader: Record<string, string> = offset > 0 ? { Range: `bytes=${offset}-` } : {}
  await waitForDriveTransferThrottle(signal)
  const graphResponse = await fetch(createDownloadContentUrl(itemId), {
    headers: {
      Authorization: `Bearer ${token}`,
      ...rangeHeader
    },
    redirect: 'manual',
    signal
  })

  if (graphResponse.status >= 300 && graphResponse.status < 400) {
    const location = graphResponse.headers.get('Location')

    if (!location) {
      throw new Error('다운로드 주소를 가져오지 못했습니다.')
    }

    const downloadResponse = await fetch(location, {
      headers: rangeHeader,
      signal
    })

    if (!downloadResponse.ok) {
      throw await createResponseError(downloadResponse, 'OneDrive 파일을 다운로드하지 못했습니다.')
    }

    return downloadResponse
  }

  if (!graphResponse.ok) {
    throw await createGraphResponseError(graphResponse)
  }

  return graphResponse
}

async function getExistingFileSize(path: string): Promise<number> {
  try {
    return (await stat(path)).size
  } catch {
    return 0
  }
}

async function createUniqueDownloadPath(directoryPath: string, fileName: string, reservedPaths: Set<string>): Promise<string> {
  const extension = extname(fileName)
  const stem = extension ? fileName.slice(0, -extension.length) : fileName

  for (let sequence = 0; sequence < 10_000; sequence += 1) {
    const candidateName = sequence === 0 ? fileName : `${stem} (${sequence})${extension}`
    const candidatePath = join(directoryPath, candidateName)
    const normalizedCandidatePath = normalizeLocalPathKey(candidatePath)

    if (
      !reservedPaths.has(normalizedCandidatePath) &&
      !(await pathExists(candidatePath)) &&
      !(await pathExists(`${candidatePath}.onedrive-download`))
    ) {
      return candidatePath
    }
  }

  throw new Error('중복되지 않는 다운로드 파일 이름을 만들지 못했습니다.')
}

function normalizeDownloadDirectoryPath(directoryPath: string): string {
  const trimmedPath = directoryPath.trim()

  if (!trimmedPath) {
    return ''
  }

  if (process.platform !== 'win32') {
    return trimmedPath
  }

  const driveOnlyMatch = /^([a-zA-Z]):[\\/]?$/.exec(trimmedPath)

  if (driveOnlyMatch?.[1]) {
    return `${driveOnlyMatch[1].toUpperCase()}:\\`
  }

  const slashDriveColonMatch = /^\/([a-zA-Z]):[\\/]?(.*)$/.exec(trimmedPath)

  if (slashDriveColonMatch?.[1]) {
    const restPath = slashDriveColonMatch[2] ? slashDriveColonMatch[2].replace(/[\\/]+/g, '\\') : ''
    return win32.normalize(`${slashDriveColonMatch[1].toUpperCase()}:\\${restPath}`)
  }

  const slashDriveRootMatch = /^\/([a-zA-Z])(?:[\\/](.*))?$/.exec(trimmedPath)

  if (slashDriveRootMatch?.[1]) {
    const restPath = slashDriveRootMatch[2] ? slashDriveRootMatch[2].replace(/[\\/]+/g, '\\') : ''
    return win32.normalize(`${slashDriveRootMatch[1].toUpperCase()}:\\${restPath}`)
  }

  return win32.normalize(trimmedPath)
}

async function ensureLocalDirectory(directoryPath: string, errorMessage: string): Promise<void> {
  try {
    const existingStat = await stat(directoryPath)

    if (existingStat.isDirectory()) {
      return
    }

    throw new Error('같은 경로에 폴더가 아닌 항목이 이미 있습니다.')
  } catch (error) {
    if (!isFileSystemNotFoundError(error)) {
      throw new Error(`${errorMessage}: ${directoryPath}${formatFileSystemError(error)}`)
    }
  }

  try {
    await mkdir(directoryPath, { recursive: true })
  } catch (error) {
    throw new Error(`${errorMessage}: ${directoryPath}${formatFileSystemError(error)}`)
  }
}

function isFileSystemNotFoundError(error: unknown): boolean {
  return typeof error === 'object' && error !== null && 'code' in error && (error as { code?: unknown }).code === 'ENOENT'
}

function normalizeLocalPathKey(localPath: string): string {
  const normalizedPath = process.platform === 'win32' ? win32.normalize(localPath) : localPath
  return normalizedPath.normalize('NFC').toLocaleLowerCase('en-US')
}

function formatFileSystemError(error: unknown): string {
  if (!error) {
    return ''
  }

  const code =
    typeof error === 'object' && 'code' in error && typeof (error as { code?: unknown }).code === 'string'
      ? (error as { code: string }).code
      : ''
  const message = error instanceof Error ? error.message : String(error)

  if (!message) {
    return code ? ` (${code})` : ''
  }

  return code ? ` (${code}: ${message})` : ` (${message})`
}

async function pathExists(path: string): Promise<boolean> {
  try {
    await stat(path)
    return true
  } catch {
    return false
  }
}

function inferDownloadTotalBytes(response: Response, offset: number, fallbackTotalBytes: number): number {
  const contentRange = response.headers.get('Content-Range')
  const rangeTotalBytes = contentRange?.match(/\/(\d+)$/)?.[1]

  if (rangeTotalBytes) {
    return Number(rangeTotalBytes)
  }

  const contentLength = Number(response.headers.get('Content-Length') ?? 0)

  if (contentLength > 0) {
    return offset + contentLength
  }

  return fallbackTotalBytes
}

async function saveTransferProgress(
  task: InternalDriveTransferTask,
  state: DriveTransferState,
  lastProgressSaveAt: number,
  onProgress?: DriveTransferProgressListener
): Promise<number> {
  const now = Date.now()

  if (now - lastProgressSaveAt < TRANSFER_PROGRESS_SAVE_INTERVAL_MS && task.transferredBytes < task.totalBytes) {
    return lastProgressSaveAt
  }

  updateTransferSpeed(task, now)
  task.updatedAt = new Date(now).toISOString()
  await writeDriveTransferState(state, onProgress)
  return now
}

function updateTransferSpeed(task: InternalDriveTransferTask, now: number): void {
  const previousSampleAt = task.speedSampleAt ?? now
  const previousSampleBytes = task.speedSampleBytes ?? task.transferredBytes
  const elapsedSeconds = (now - previousSampleAt) / 1000

  if (elapsedSeconds <= 0) {
    return
  }

  task.bytesPerSecond = Math.max(0, Math.round((task.transferredBytes - previousSampleBytes) / elapsedSeconds))
  task.speedSampleAt = now
  task.speedSampleBytes = task.transferredBytes
}

async function pauseTransferTask(
  task: InternalDriveTransferTask,
  state: DriveTransferState,
  onProgress?: DriveTransferProgressListener
): Promise<void> {
  task.status = 'paused'
  task.bytesPerSecond = 0
  task.message = '전송이 중지되었습니다.'
  task.updatedAt = new Date().toISOString()
  await writeDriveTransferState(state, onProgress)
}

async function removeCompletedDriveTransferTask(task: InternalDriveTransferTask, onProgress?: DriveTransferProgressListener): Promise<void> {
  task.status = 'completed'
  task.transferredBytes = task.totalBytes > 0 ? task.totalBytes : task.transferredBytes
  task.bytesPerSecond = 0
  task.nextRetryAt = undefined
  task.lastError = undefined
  task.failureStage = undefined
  task.updatedAt = new Date().toISOString()

  await markAccountTransferGroupTaskCompleted(task)
  await recordCompletedTransferTask(task)
  await cleanupTransferTask(task)
  await deleteDriveTransferTaskFile(task.id)
  await removeDriveTransferIndexTaskIds([task.id])
  await resetCompletedTransferSummaryIfQueueFinished(task)
  transferScanCursor = 0
  await recordSuccessfulTransferCompletion()
  await emitTransferSnapshot(onProgress)
}

async function markAccountTransferGroupTaskCompleted(task: InternalDriveTransferTask): Promise<void> {
  if (task.kind !== 'account-transfer' || !task.transferGroupId || task.cleanupOnly) {
    return
  }

  await withTransferMetadataMutation(async () => {
    const index = await readDriveTransferIndex()

    for (const taskId of index.taskIds) {
      if (taskId === task.id) {
        continue
      }

      const candidate = await readDriveTransferTask(taskId)

      if (!candidate?.cleanupOnly || candidate.transferGroupId !== task.transferGroupId) {
        continue
      }

      candidate.completedGroupTaskIds = [...new Set([...(candidate.completedGroupTaskIds ?? []), task.id])]
      candidate.updatedAt = new Date().toISOString()
      await writeDriveTransferTask(candidate)
    }
  })
}

async function cleanupTransferTask(task: InternalDriveTransferTask): Promise<void> {
  if (task.uploadUrl) {
    try {
      await fetch(task.uploadUrl, { method: 'DELETE' })
    } catch {
      // Upload sessions may already be expired or aborted.
    }
  }

  const tempPath = task.tempPath ?? (task.kind === 'download' && task.localPath ? `${task.localPath}.onedrive-download` : null)

  if (tempPath) {
    try {
      await unlink(tempPath)
    } catch {
      // Partial download file may not exist.
    }
  }
}

async function failTransferTask(
  task: InternalDriveTransferTask,
  state: DriveTransferState,
  message: string,
  onProgress?: DriveTransferProgressListener
): Promise<void> {
  task.status = 'failed'
  task.bytesPerSecond = 0
  task.message = message
  task.updatedAt = new Date().toISOString()
  await writeDriveTransferState(state, onProgress)
}

async function scheduleRetryTransferTask(
  task: InternalDriveTransferTask,
  state: DriveTransferState,
  message: string,
  failureStage: TransferFailureStage,
  onProgress?: DriveTransferProgressListener,
  retryAfterDelayMs?: number,
  shouldThrottleTransfers = false
): Promise<void> {
  const attemptCount = (task.attemptCount ?? 0) + 1
  const retryDelayMs = getRetryDelayMs(attemptCount, retryAfterDelayMs)
  const nextRetryAt = new Date(Date.now() + retryDelayMs).toISOString()

  if (shouldThrottleTransfers) {
    await throttleDriveTransfers(retryDelayMs)
  }

  task.status = 'retrying'
  task.bytesPerSecond = 0
  task.attemptCount = attemptCount
  task.nextRetryAt = nextRetryAt
  task.lastError = message
  task.failureStage = failureStage
  task.message = `${getFailureStageLabel(failureStage)} 실패: ${message} · ${formatRetryDelay(retryDelayMs)} 후 자동 재시도`
  task.updatedAt = new Date().toISOString()

  await writeDriveTransferState(state, onProgress)
  scheduleDriveTransferRetryWorker(retryDelayMs)
}

function getRetryDelayMs(attemptCount: number, minimumDelayMs?: number): number {
  const exponentialDelayMs = TRANSFER_RETRY_BASE_DELAY_MS * 2 ** Math.max(0, attemptCount - 1)
  const cappedDelayMs = Math.min(exponentialDelayMs, TRANSFER_RETRY_MAX_DELAY_MS)
  const jitterMs = Math.floor(Math.random() * Math.min(5000, cappedDelayMs / 3))

  return Math.max(cappedDelayMs + jitterMs, minimumDelayMs ?? 0)
}

async function throttleDriveTransfers(delayMs: number): Promise<void> {
  const now = Date.now()
  const normalizedDelayMs = Math.max(delayMs, TRANSFER_DISPATCH_DELAY_MS)

  transferThrottleUntil = Math.max(transferThrottleUntil, now + normalizedDelayMs)
  emitGraphActivity({
    level: 'warning',
    scope: 'transfer',
    title: '전송 요청 제한',
    message: `OneDrive 전송 요청이 제한되어 ${formatRetryDelay(normalizedDelayMs)} 동안 새 요청을 늦춥니다.`
  })
}

function getDriveTransferThrottleDelayMs(): number {
  return Math.max(0, transferThrottleUntil - Date.now())
}

async function waitForDriveTransferThrottle(signal?: AbortSignal): Promise<void> {
  const delayMs = getDriveTransferThrottleDelayMs()

  if (delayMs <= 0) {
    return
  }

  await waitForDelay(delayMs, signal)
}

async function waitForDelay(delayMs: number, signal?: AbortSignal): Promise<void> {
  if (signal?.aborted) {
    throw createAbortError()
  }

  await new Promise<void>((resolve, reject) => {
    const cleanup = (): void => {
      signal?.removeEventListener('abort', handleAbort)
    }
    const timer = setTimeout(() => {
      cleanup()
      resolve()
    }, delayMs)
    const handleAbort = (): void => {
      clearTimeout(timer)
      cleanup()
      reject(createAbortError())
    }

    signal?.addEventListener('abort', handleAbort, { once: true })
  })
}

function createAbortError(): Error {
  const error = new Error('전송이 중지되었습니다.')
  error.name = 'AbortError'
  return error
}

async function recordSuccessfulTransferCompletion(): Promise<void> {
  scheduleDriveTransferRetryWorker(TRANSFER_DISPATCH_DELAY_MS)
}

function getEffectiveTransferSlotLimit(configuredSlotLimit: number): number {
  return Math.max(TRANSFER_SLOT_MIN, Math.floor(configuredSlotLimit))
}

function scheduleDriveTransferRetryWorker(delayMs = TRANSFER_RETRY_IDLE_DELAY_MS): void {
  const normalizedDelayMs = Math.max(TRANSFER_DISPATCH_DELAY_MS, delayMs)
  const nextDueAt = Date.now() + normalizedDelayMs

  if (transferRetryTimer && transferRetryTimerDueAt <= nextDueAt) {
    return
  }

  if (transferRetryTimer) {
    windowClearTimeout(transferRetryTimer)
    transferRetryTimer = null
  }

  transferRetryTimerDueAt = nextDueAt
  transferRetryTimer = setTimeout(() => {
    transferRetryTimer = null
    transferRetryTimerDueAt = 0
    void runDriveTransferRetryWorker()
  }, normalizedDelayMs)
}

function stopDriveTransferRetryScheduler(): void {
  if (transferRetryTimer) {
    windowClearTimeout(transferRetryTimer)
    transferRetryTimer = null
    transferRetryTimerDueAt = 0
  }
}

async function runDriveTransferRetryWorker(): Promise<void> {
  if (transferRetryWorkerPromise) {
    return transferRetryWorkerPromise
  }

  let startedCount = 0
  let startedRetryingCount = 0

  transferRetryWorkerPromise = (async () => {
    const throttleDelayMs = getDriveTransferThrottleDelayMs()

    if (throttleDelayMs > 0) {
      scheduleDriveTransferRetryWorker(throttleDelayMs)
      return
    }

    const availableSlots = await getAvailableTransferSlots()

    if (availableSlots <= 0) {
      return
    }

    const tasks = await findRetryableTransferTasks(Math.min(TRANSFER_RETRY_BATCH_LIMIT, availableSlots), false, {
      mode: 'non-retrying'
    })

    for (const task of tasks) {
      if (task.status !== 'paused' && (await startDriveTransferTask(task, transferRetryProgressListener))) {
        startedCount += 1
      }
    }

    const remainingSlots = await getAvailableTransferSlots()

    if (remainingSlots <= 0) {
      return
    }

    const retryingStepDelayMs = getRetryingTransferDispatchDelayMs()

    if (retryingStepDelayMs > 0) {
      scheduleDriveTransferRetryWorker(retryingStepDelayMs)
      return
    }

    const [retryingTask] = await findRetryableTransferTasks(1, false, {
      mode: 'retrying',
      minimumRetryingWaitMs: TRANSFER_RETRY_READY_WAIT_MS
    })

    if (retryingTask && (await startDriveTransferTask(retryingTask, transferRetryProgressListener))) {
      startedCount += 1
      startedRetryingCount += 1
      nextRetryingTransferDispatchAt = Date.now() + TRANSFER_RETRY_STEP_DELAY_MS
    }
  })().finally(async () => {
    transferRetryWorkerPromise = null

    const throttleDelayMs = getDriveTransferThrottleDelayMs()

    if (throttleDelayMs > 0) {
      scheduleDriveTransferRetryWorker(throttleDelayMs)
      return
    }

    if ((await getAvailableTransferSlots()) <= 0) {
      return
    }

    if (startedCount > 0) {
      scheduleDriveTransferRetryWorker(startedRetryingCount > 0 ? TRANSFER_RETRY_STEP_DELAY_MS : TRANSFER_DISPATCH_DELAY_MS)
      return
    }

    const nextRetryDelayMs = await getNextRetryingTransferReadyDelayMs(TRANSFER_RETRY_READY_WAIT_MS)

    if (nextRetryDelayMs !== undefined) {
      scheduleDriveTransferRetryWorker(Math.max(nextRetryDelayMs, getRetryingTransferDispatchDelayMs()))
      return
    }

    if (getOccupiedTransferSlotCount() === 0) {
      scheduleDriveTransferRetryWorker(TRANSFER_RETRY_IDLE_DELAY_MS)
    }
  })

  return transferRetryWorkerPromise
}

async function getAvailableTransferSlots(): Promise<number> {
  const settings = await getTransferSettings()
  const effectiveSlotLimit = getEffectiveTransferSlotLimit(settings.maxConcurrentTransfers)

  return Math.max(0, effectiveSlotLimit - getOccupiedTransferSlotCount())
}

function getOccupiedTransferSlotCount(): number {
  return new Set([...transferWorkerSlotIds, ...runningTransferIds]).size
}

function getRetryingTransferDispatchDelayMs(): number {
  return Math.max(0, nextRetryingTransferDispatchAt - Date.now())
}

async function startDriveTransferTask(task: InternalDriveTransferTask, onProgress?: DriveTransferProgressListener): Promise<boolean> {
  if (transferWorkerSlotIds.has(task.id) || runningTransferIds.has(task.id)) {
    return false
  }

  if ((await getAvailableTransferSlots()) <= 0) {
    return false
  }

  transferWorkerSlotIds.add(task.id)
  void runDriveTransferTask(task, onProgress)
    .catch(() => {
      // The task file records retry, pause, and failure details.
    })
    .finally(() => {
      transferWorkerSlotIds.delete(task.id)
      scheduleDriveTransferRetryWorker(TRANSFER_DISPATCH_DELAY_MS)
    })

  return true
}

async function runDriveTransferTask(task: InternalDriveTransferTask, onProgress?: DriveTransferProgressListener): Promise<void> {
  if (task.kind === 'upload') {
    await processUploadTask(task, { version: 1, tasks: [task] }, onProgress)
    return
  }

  if (task.kind === 'account-transfer') {
    await processAccountTransferTask(task, { version: 1, tasks: [task] }, onProgress)
    return
  }

  await processDownloadTask(task, { version: 1, tasks: [task] }, onProgress)
}

function windowClearTimeout(timer: ReturnType<typeof setTimeout>): void {
  clearTimeout(timer)
}

function getTransferKindLabelForMessage(kind: DriveTransferTask['kind']): string {
  if (kind === 'upload') {
    return '업로드'
  }

  if (kind === 'download') {
    return '다운로드'
  }

  return '계정 간 전송'
}

function getFailureStageLabel(stage: TransferFailureStage): string {
  const labels: Record<TransferFailureStage, string> = {
    auth: '인증',
    metadata: '메타데이터',
    'local-file': '로컬 파일',
    'upload-session': '업로드 세션',
    'upload-chunk': '업로드 조각',
    'download-link': '다운로드 주소',
    'download-stream': '다운로드 스트림',
    finalize: '마무리 저장',
    unknown: '알 수 없는 지점'
  }

  return labels[stage]
}

function formatRetryDelay(delayMs: number): string {
  const seconds = Math.max(1, Math.round(delayMs / 1000))

  if (seconds < 60) {
    return `${seconds}초`
  }

  return `${Math.round(seconds / 60)}분`
}

function getErrorMessage(error: unknown, fallbackMessage: string): string {
  return error instanceof Error ? error.message : fallbackMessage
}

function getRetryAfterDelayMs(error: unknown): number | undefined {
  if (!(error instanceof GraphResponseError)) {
    return undefined
  }

  if (error.retryAfterMs !== undefined) {
    return Math.max(TRANSFER_DISPATCH_DELAY_MS, error.retryAfterMs)
  }

  return error.status === 429 ? DEFAULT_THROTTLE_RETRY_DELAY_MS : undefined
}

function isResponseThrottleError(error: unknown): boolean {
  return error instanceof GraphResponseError && (error.status === 429 || error.retryAfterMs !== undefined)
}

function throwIfTransferPaused(taskId: string): void {
  if (transferPauseRequests.has(taskId)) {
    throw new TransferPausedError()
  }
}

function isTransferPauseError(error: unknown, taskId: string): boolean {
  return error instanceof TransferPausedError || transferPauseRequests.has(taskId) || (error instanceof Error && error.name === 'AbortError')
}

function parseGraphUrl(value: string, errorMessage: string): URL {
  const url = new URL(value)

  if (url.hostname !== 'graph.microsoft.com') {
    throw new Error(errorMessage)
  }

  return url
}

function parseDriveDeltaUrl(value: string, errorMessage: string): URL {
  const url = parseGraphUrl(value, errorMessage)

  if (!isDriveDeltaUrl(url)) {
    throw new Error(errorMessage)
  }

  return url
}

function parseStoredDriveDeltaUrl(value: string | null | undefined): URL | null {
  if (!value) {
    return null
  }

  try {
    return parseDriveDeltaUrl(value, '저장된 OneDrive delta cursor가 올바르지 않습니다.')
  } catch {
    return null
  }
}

function normalizeDriveDeltaLink(value: string | undefined): string | undefined {
  return parseStoredDriveDeltaUrl(value)?.toString()
}

function isDriveDeltaUrl(url: URL): boolean {
  if (url.protocol !== 'https:' || url.hostname !== 'graph.microsoft.com') {
    return false
  }

  const pathname = url.pathname.toLocaleLowerCase('en-US')

  return /\/drive\/root\/delta(?:$|\()/.test(pathname) || /\/drives\/[^/]+\/root\/delta(?:$|\()/.test(pathname)
}

function parseTrustedCopyMonitorUrl(value: string): URL {
  const url = new URL(value)
  const hostname = url.hostname.toLowerCase()

  if (
    url.protocol !== 'https:' ||
    (hostname !== 'graph.microsoft.com' && !hostname.endsWith('.sharepoint.com') && hostname !== 'api.onedrive.com')
  ) {
    throw new Error('OneDrive 복사 진행 상태 주소가 올바르지 않습니다.')
  }

  return url
}

function getAvailableCopyName(name: string, type: CloudDriveItemType, occupiedNames: Set<string>): string | undefined {
  if (!occupiedNames.has(normalizeDriveItemNameForConflict(name))) {
    return undefined
  }

  const extension = type === 'file' ? extname(name) : ''
  const baseName = extension ? name.slice(0, -extension.length) : name
  let candidate = `${baseName} - 복사본${extension}`
  let copyIndex = 2

  while (occupiedNames.has(normalizeDriveItemNameForConflict(candidate))) {
    candidate = `${baseName} - 복사본 (${copyIndex})${extension}`
    copyIndex += 1
  }

  return candidate
}

function normalizeDriveItemNameForConflict(name: string): string {
  return name.trim().toLocaleLowerCase('ko-KR')
}

function delay(delayMs: number): Promise<void> {
  return new Promise((resolve) => {
    setTimeout(resolve, delayMs)
  })
}

function validateDriveItemId(itemId: string): string {
  const normalizedItemId = itemId.trim()

  if (!normalizedItemId) {
    throw new Error('OneDrive 항목을 선택하세요.')
  }

  return normalizedItemId
}

function validateAccountId(accountId: string, errorMessage: string): string {
  const normalizedAccountId = accountId.trim()

  if (!normalizedAccountId) {
    throw new Error(errorMessage)
  }

  return normalizedAccountId
}

function validateDriveItemName(name: string): string {
  const normalizedName = name.trim()

  if (!normalizedName) {
    throw new Error('파일 이름을 입력하세요.')
  }

  if (/[<>:"/\\|?*\u0000-\u001f]/.test(normalizedName)) {
    throw new Error('파일 이름에 사용할 수 없는 문자가 있습니다.')
  }

  return normalizedName
}

function parseUploadRangeStart(value: string | undefined): number | null {
  if (!value) {
    return null
  }

  const start = Number(value.split('-')[0])

  return Number.isFinite(start) ? start : null
}

function isExpired(value: string | undefined): boolean {
  if (!value) {
    return false
  }

  return new Date(value).getTime() <= Date.now() + 30_000
}

function encodeDrivePathSegment(value: string): string {
  return encodeURIComponent(value).replaceAll('%20', ' ')
}

async function readDriveTransferState(): Promise<DriveTransferState> {
  await migrateLegacyDriveTransfers()

  const activeAccountId = await getActiveAccountId()
  const index = await readDriveTransferIndex()
  const visibleTaskIds = [...new Set([...runningTransferIds, ...index.taskIds.slice(0, TRANSFER_VISIBLE_LIMIT)])]
  const tasks = (await Promise.all(visibleTaskIds.map((taskId) => readDriveTransferTask(taskId)))).flatMap((task) =>
    task && isDriveTransferTaskForAccount(task, activeAccountId) ? [normalizeLoadedTransferTask(task)] : []
  )

  return {
    version: 1,
    tasks: sortTransferTasks(tasks)
  }
}

async function writeDriveTransferState(state: DriveTransferState, onProgress?: DriveTransferProgressListener): Promise<void> {
  for (const task of state.tasks) {
    await writeDriveTransferTask(task)
  }

  await emitTransferSnapshot(onProgress)
}

function publicTransferTasks(tasks: InternalDriveTransferTask[]): DriveTransferTask[] {
  return tasks.map((task) => ({
    id: task.id,
    kind: task.kind,
    status: task.status,
    name: task.name,
    transferredBytes: task.transferredBytes,
    totalBytes: task.totalBytes,
    bytesPerSecond: task.bytesPerSecond,
    attemptCount: task.attemptCount,
    nextRetryAt: task.nextRetryAt,
    lastError: task.lastError,
    failureStage: task.failureStage,
    createdAt: task.createdAt,
    updatedAt: task.updatedAt,
    localPath: task.localPath,
    message: task.message
  }))
}

function applyTransferTaskToSummary(summary: DriveTransferSummary, task: InternalDriveTransferTask): void {
  summary.activeCount += 1
  summary.totalCount += 1
  summary.totalBytes += Math.max(task.totalBytes, 0)
  summary.transferredBytes += Math.max(task.transferredBytes, 0)
  summary.bytesPerSecond += task.status === 'running' ? Math.max(task.bytesPerSecond ?? 0, 0) : 0

  if (task.status === 'queued') {
    summary.queuedCount += 1
    return
  }

  if (task.status === 'running') {
    summary.runningCount += 1
    return
  }

  if (task.status === 'paused') {
    summary.pausedCount += 1
    return
  }

  if (task.status === 'retrying') {
    summary.retryingCount += 1
    return
  }

  if (task.status === 'failed') {
    summary.failedCount += 1
  }
}

async function emitTransferSnapshot(onProgress?: DriveTransferProgressListener): Promise<void> {
  onProgress?.(publicTransferTasks((await readDriveTransferState()).tasks))
}

async function withTransferMetadataMutation<T>(mutation: () => Promise<T>): Promise<T> {
  const previousMutation = transferMetadataMutationPromise
  let releaseMutation!: () => void

  transferMetadataMutationPromise = new Promise<void>((resolve) => {
    releaseMutation = resolve
  })

  await previousMutation

  try {
    return await mutation()
  } finally {
    releaseMutation()
  }
}

async function registerDriveTransferTasks(tasks: InternalDriveTransferTask[]): Promise<void> {
  if (tasks.length === 0) {
    return
  }

  await withTransferMetadataMutation(async () => {
    const index = await readDriveTransferIndex()
    const existingTaskIds = new Set(index.taskIds)
    const nextTaskIds = tasks.map((task) => task.id).filter((taskId) => !existingTaskIds.has(taskId))

    if (nextTaskIds.length === 0) {
      return
    }

    index.taskIds = [...nextTaskIds, ...index.taskIds]
    transferScanCursor = 0
    await writeDriveTransferIndex(index)
  })
}

async function removeDriveTransferIndexTaskIds(taskIds: string[]): Promise<void> {
  if (taskIds.length === 0) {
    return
  }

  await withTransferMetadataMutation(async () => {
    const taskIdSet = new Set(taskIds)
    const index = await readDriveTransferIndex()
    index.taskIds = index.taskIds.filter((taskId) => !taskIdSet.has(taskId))
    await writeDriveTransferIndex(index)
  })
}

async function readDriveTransferIndex(): Promise<DriveTransferIndex> {
  await migrateLegacyDriveTransfers()

  if (transferIndexCache) {
    return transferIndexCache
  }

  transferIndexCache = await readDriveTransferIndexFile()
  return transferIndexCache
}

async function readDriveTransferIndexFile(): Promise<DriveTransferIndex> {
  try {
    const index = JSON.parse(await readFile(getDriveTransferIndexPath(), 'utf8')) as DriveTransferIndex

    if (index.version === 1 && Array.isArray(index.taskIds)) {
      return {
        version: 1,
        taskIds: [...new Set(index.taskIds.filter((taskId) => typeof taskId === 'string' && taskId.trim()))]
      }
    }
  } catch {
    // Missing or invalid queue indexes start empty.
  }

  return {
    version: 1,
    taskIds: []
  }
}

async function writeDriveTransferIndex(index: DriveTransferIndex): Promise<void> {
  transferIndexCache = {
    version: 1,
    taskIds: [...new Set(index.taskIds)]
  }
  await writeJsonAtomic(getDriveTransferIndexPath(), transferIndexCache)
}

async function readCompletedTransferSummary(accountId: string | null): Promise<DriveTransferCompletedSummary> {
  const summary = await readDriveTransferSummaryFile()

  return summary.accounts[getTransferSummaryAccountKey(accountId)] ?? {
    completedCount: 0,
    completedBytes: 0
  }
}

async function recordCompletedTransferTask(task: InternalDriveTransferTask): Promise<void> {
  if (task.cleanupOnly) {
    return
  }

  await withTransferMetadataMutation(async () => {
    const summary = await readDriveTransferSummaryFile()
    const accountKey = getTransferSummaryAccountKey(task.accountId ?? (await getActiveAccountId()))
    const currentSummary = summary.accounts[accountKey] ?? {
      completedCount: 0,
      completedBytes: 0
    }

    summary.accounts[accountKey] = {
      completedCount: currentSummary.completedCount + 1,
      completedBytes: currentSummary.completedBytes + Math.max(task.totalBytes, task.transferredBytes, 0)
    }
    await writeDriveTransferSummaryFile(summary)
  })
}

async function resetCompletedTransferSummary(accountId: string | null): Promise<DriveTransferCompletedSummary> {
  await withTransferMetadataMutation(async () => {
    const summary = await readDriveTransferSummaryFile()

    delete summary.accounts[getTransferSummaryAccountKey(accountId)]
    await writeDriveTransferSummaryFile(summary)
  })

  return {
    completedCount: 0,
    completedBytes: 0
  }
}

async function resetCompletedTransferSummaryIfQueueFinished(task: InternalDriveTransferTask): Promise<void> {
  const accountId = task.accountId ?? (await getActiveAccountId())
  const index = await readDriveTransferIndex()

  if (await hasIndexedTransferTaskForAccount(index, accountId)) {
    return
  }

  await resetCompletedTransferSummary(accountId)
}

async function hasIndexedTransferTaskForAccount(index: DriveTransferIndex, accountId: string | null): Promise<boolean> {
  for (const taskId of index.taskIds) {
    const task = await readDriveTransferTask(taskId)

    if (task && isDriveTransferTaskForAccount(task, accountId)) {
      return true
    }
  }

  return false
}

async function readDriveTransferSummaryFile(): Promise<DriveTransferSummaryFile> {
  try {
    const summary = JSON.parse(await readFile(getDriveTransferSummaryPath(), 'utf8')) as DriveTransferSummaryFile

    if (summary.version === 1 && summary.accounts && typeof summary.accounts === 'object') {
      return {
        version: 1,
        accounts: Object.fromEntries(
          Object.entries(summary.accounts).map(([accountId, accountSummary]) => [
            accountId,
            {
              completedCount: Math.max(0, Math.floor(accountSummary.completedCount ?? 0)),
              completedBytes: Math.max(0, Math.floor(accountSummary.completedBytes ?? 0))
            }
          ])
        )
      }
    }
  } catch {
    // Missing or invalid summaries start empty.
  }

  return {
    version: 1,
    accounts: {}
  }
}

async function writeDriveTransferSummaryFile(summary: DriveTransferSummaryFile): Promise<void> {
  await writeJsonAtomic(getDriveTransferSummaryPath(), {
    version: 1,
    accounts: summary.accounts
  })
}

function getTransferSummaryAccountKey(accountId: string | null | undefined): string {
  return accountId?.trim() || 'unscoped'
}

async function readDriveTransferTask(taskId: string): Promise<InternalDriveTransferTask | null> {
  const normalizedTaskId = taskId.trim()

  if (!normalizedTaskId) {
    return null
  }

  try {
    const task = JSON.parse(await readFile(getDriveTransferTaskPath(normalizedTaskId), 'utf8')) as InternalDriveTransferTask

    if (!task.id || task.id !== normalizedTaskId || !task.kind || !task.status) {
      return null
    }

    if (task.status === 'running' && !transferWorkerSlotIds.has(task.id) && !runningTransferIds.has(task.id)) {
      task.status = 'retrying'
      task.bytesPerSecond = 0
      task.message = '앱이 종료되어 전송이 끊겼습니다. 자동으로 다시 시도합니다.'
      task.nextRetryAt = new Date().toISOString()
      task.updatedAt = new Date().toISOString()
      await writeDriveTransferTask(task)
    }

    return task
  } catch {
    return null
  }
}

async function writeDriveTransferTask(task: InternalDriveTransferTask): Promise<void> {
  await writeJsonAtomic(getDriveTransferTaskPath(task.id), task)
}

async function deleteDriveTransferTaskFile(taskId: string): Promise<void> {
  await unlinkIfExists(getDriveTransferTaskPath(taskId))
}

async function* iterateAllDriveTransferTasks(): AsyncGenerator<InternalDriveTransferTask> {
  const index = await readDriveTransferIndex()

  for (const taskId of index.taskIds) {
    const task = await readDriveTransferTask(taskId)

    if (task) {
      yield task
    }
  }
}

async function findRetryableTransferTasks(
  limit: number,
  includePaused: boolean,
  options: TransferRetryCandidateOptions = {}
): Promise<InternalDriveTransferTask[]> {
  const index = await readDriveTransferIndex()
  const tasks: InternalDriveTransferTask[] = []
  const activeAccountId = await getActiveAccountId()

  if (index.taskIds.length === 0) {
    return tasks
  }

  let scannedCount = 0

  while (scannedCount < index.taskIds.length && tasks.length < limit) {
    const taskId = index.taskIds[transferScanCursor % index.taskIds.length]
    transferScanCursor = (transferScanCursor + 1) % index.taskIds.length
    scannedCount += 1

    if (!taskId || transferWorkerSlotIds.has(taskId) || runningTransferIds.has(taskId)) {
      continue
    }

    const task = await readDriveTransferTask(taskId)

    if (task && isDriveTransferTaskForAccount(task, activeAccountId) && isTransferReadyToRun(task, includePaused, options)) {
      tasks.push(task)
    }
  }

  return tasks
}

async function getNextRetryingTransferReadyDelayMs(minimumRetryingWaitMs: number): Promise<number | undefined> {
  const index = await readDriveTransferIndex()
  const activeAccountId = await getActiveAccountId()
  const now = Date.now()
  let nextDelayMs: number | undefined

  for (const taskId of index.taskIds) {
    if (!taskId || transferWorkerSlotIds.has(taskId) || runningTransferIds.has(taskId)) {
      continue
    }

    const task = await readDriveTransferTask(taskId)

    if (!task || task.status !== 'retrying' || !isDriveTransferTaskForAccount(task, activeAccountId)) {
      continue
    }

    const delayMs = Math.max(TRANSFER_DISPATCH_DELAY_MS, getRetryingTransferReadyAt(task, minimumRetryingWaitMs) - now)
    nextDelayMs = nextDelayMs === undefined ? delayMs : Math.min(nextDelayMs, delayMs)
  }

  return nextDelayMs
}

function isDriveTransferTaskForAccount(task: InternalDriveTransferTask, activeAccountId: string | null): boolean {
  if (!activeAccountId) {
    return false
  }

  return !task.accountId || task.accountId === activeAccountId
}

function isTransferReadyToRun(
  task: InternalDriveTransferTask,
  includePaused: boolean,
  options: TransferRetryCandidateOptions = {}
): boolean {
  if (task.status === 'completed' || transferWorkerSlotIds.has(task.id) || runningTransferIds.has(task.id)) {
    return false
  }

  if (options.mode === 'non-retrying' && task.status === 'retrying') {
    return false
  }

  if (options.mode === 'retrying' && task.status !== 'retrying') {
    return false
  }

  if (task.status === 'paused') {
    return includePaused
  }

  if (task.status === 'running') {
    return true
  }

  if (task.status === 'retrying' && getRetryingTransferReadyAt(task, options.minimumRetryingWaitMs ?? 0) > Date.now()) {
    return false
  }

  if (!task.nextRetryAt) {
    return task.status === 'queued' || task.status === 'retrying' || task.status === 'failed'
  }

  return new Date(task.nextRetryAt).getTime() <= Date.now()
}

function getRetryingTransferReadyAt(task: InternalDriveTransferTask, minimumRetryingWaitMs: number): number {
  const nextRetryAt = task.nextRetryAt ? Date.parse(task.nextRetryAt) : 0
  const updatedAt = Date.parse(task.updatedAt)
  const minimumReadyAt = Number.isFinite(updatedAt) ? updatedAt + minimumRetryingWaitMs : 0

  return Math.max(Number.isFinite(nextRetryAt) ? nextRetryAt : 0, minimumReadyAt)
}

function sortTransferTasks(tasks: InternalDriveTransferTask[]): InternalDriveTransferTask[] {
  return [...tasks].sort((left, right) => {
    const progressDifference = getTransferProgressRatio(right) - getTransferProgressRatio(left)

    if (progressDifference !== 0) {
      return progressDifference
    }

    return new Date(right.updatedAt).getTime() - new Date(left.updatedAt).getTime()
  })
}

function getTransferProgressRatio(task: InternalDriveTransferTask): number {
  if (task.totalBytes <= 0) {
    return task.status === 'completed' ? 1 : 0
  }

  return Math.min(1, Math.max(0, task.transferredBytes / task.totalBytes))
}

function normalizeLoadedTransferTask(task: InternalDriveTransferTask): InternalDriveTransferTask {
  return {
    ...task,
    bytesPerSecond: task.status === 'running' ? task.bytesPerSecond : 0
  }
}

async function migrateLegacyDriveTransfers(): Promise<void> {
  if (hasCheckedLegacyDriveTransfers) {
    return
  }

  const legacyPath = getLegacyDriveTransfersPath()

  if (!(await pathExists(legacyPath))) {
    hasCheckedLegacyDriveTransfers = true
    return
  }

  try {
    const state = JSON.parse(await readFile(legacyPath, 'utf8')) as DriveTransferState

    if (state.version === 1 && Array.isArray(state.tasks)) {
      const index = await readDriveTransferIndexFile()
      const existingTaskIds = new Set(index.taskIds)
      const taskIdsToAdd: string[] = []

      for (const task of state.tasks) {
        if (!task.id) {
          continue
        }

        await writeDriveTransferTask(task)

        if (!existingTaskIds.has(task.id)) {
          taskIdsToAdd.push(task.id)
        }
      }

      if (taskIdsToAdd.length > 0) {
        index.taskIds = [...taskIdsToAdd, ...index.taskIds]
        await writeDriveTransferIndex(index)
      }
    }

    await unlinkIfExists(legacyPath)
    hasCheckedLegacyDriveTransfers = true
  } catch {
    hasCheckedLegacyDriveTransfers = true
    // A corrupt legacy queue should not block the new durable queue.
  }
}

async function writeJsonAtomic(path: string, value: unknown): Promise<void> {
  await mkdir(dirname(path), { recursive: true })

  const temporaryPath = `${path}.${process.pid}.${randomUUID()}.tmp`

  await writeFile(temporaryPath, JSON.stringify(value), 'utf8')
  await rename(temporaryPath, path)
}

async function unlinkIfExists(path: string): Promise<void> {
  try {
    await unlink(path)
  } catch {
    // The file may already be gone.
  }
}

async function readDriveIndex(accountId?: string | null): Promise<DriveIndex> {
  const normalizedAccountId = accountId ?? (await getActiveAccountId())

  if (!normalizedAccountId) {
    return createEmptyDriveIndex()
  }

  try {
    const index = await readDriveIndexFromStore(normalizedAccountId)

    if (index) {
      return index
    }
  } catch {
    // Missing or invalid SQLite indexes are rebuilt from Graph delta.
  }

  try {
    const legacyIndex = normalizeStoredDriveIndex(JSON.parse(await readFile(getLegacyAccountDriveIndexPath(normalizedAccountId), 'utf8')))

    if (legacyIndex) {
      await writeDriveIndex(legacyIndex, normalizedAccountId)
      await unlinkIfExists(getLegacyAccountDriveIndexPath(normalizedAccountId))
      return legacyIndex
    }
  } catch {
    // Per-account JSON indexes were used before the SQLite index store.
  }

  try {
    const legacyIndex = normalizeStoredDriveIndex(JSON.parse(await readFile(getLegacyDriveIndexPath(), 'utf8')))

    if (legacyIndex) {
      await writeDriveIndex(legacyIndex, normalizedAccountId)
      await unlinkIfExists(getLegacyDriveIndexPath())
      return legacyIndex
    }
  } catch {
    // Legacy index may not exist.
  }

  return createEmptyDriveIndex()
}

async function getCurrentDriveIndex(accountId?: string | null): Promise<DriveIndex> {
  const normalizedAccountId = accountId ?? (await getActiveAccountId())

  if (activeDriveIndexSnapshot && activeDriveIndexSnapshotAccountId === normalizedAccountId) {
    return activeDriveIndexSnapshot
  }

  const index = await readDriveIndex(normalizedAccountId)

  if (normalizedAccountId) {
    activeDriveIndexSnapshot = index
    activeDriveIndexSnapshotAccountId = normalizedAccountId
  }

  return index
}

async function writeDriveIndex(index: DriveIndex, accountId?: string | null): Promise<void> {
  const normalizedAccountId = accountId ?? (await getActiveAccountId())

  activeDriveIndexSnapshot = index
  activeDriveIndexSnapshotAccountId = normalizedAccountId
  await writeDriveIndexToStore(index, normalizedAccountId)
}

function normalizeStoredDriveIndex(value: unknown): DriveIndex | null {
  const index = value as Partial<DriveIndex>

  if (index.version !== 1 || !index.items) {
    return null
  }

  return {
    ...index,
    version: 1,
    expandedFolderIds: index.expandedFolderIds ?? {},
    items: index.items
  }
}

function getLegacyDriveIndexPath(): string {
  return join(app.getPath('userData'), DRIVE_INDEX_LEGACY_FILE_NAME)
}

function getLegacyDriveTransfersPath(): string {
  return join(app.getPath('userData'), DRIVE_TRANSFERS_LEGACY_FILE_NAME)
}

function getDriveTransfersDirectory(): string {
  return join(app.getPath('userData'), DRIVE_TRANSFERS_DIR_NAME)
}

function getDriveThumbnailCacheDirectory(): string {
  return join(app.getPath('userData'), DRIVE_THUMBNAIL_CACHE_DIR_NAME)
}

function getDriveThumbnailCachePaths(cacheKey: string): { dataPath: string; metadataPath: string } {
  const cacheId = createHash('sha256').update(cacheKey).digest('hex')
  const cacheDirectory = getDriveThumbnailCacheDirectory()

  return {
    dataPath: join(cacheDirectory, `${cacheId}.bin`),
    metadataPath: join(cacheDirectory, `${cacheId}.json`)
  }
}

function getAccountTransferTempPath(taskId: string): string {
  return join(getDriveTransfersDirectory(), DRIVE_TRANSFERS_TEMP_DIR_NAME, `${taskId}.bin`)
}

function getDriveTransferIndexPath(): string {
  return join(getDriveTransfersDirectory(), DRIVE_TRANSFERS_INDEX_FILE_NAME)
}

function getDriveTransferSummaryPath(): string {
  return join(getDriveTransfersDirectory(), DRIVE_TRANSFERS_SUMMARY_FILE_NAME)
}

function getDriveTransferTaskPath(taskId: string): string {
  return join(getDriveTransfersDirectory(), DRIVE_TRANSFERS_TASKS_DIR_NAME, taskId.slice(0, 2), `${taskId}.json`)
}
