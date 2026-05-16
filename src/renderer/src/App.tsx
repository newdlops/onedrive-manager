import { useCallback, useEffect, useRef, useState } from 'react'
import type {
  Dispatch,
  DragEvent,
  FormEvent,
  MouseEvent as ReactMouseEvent,
  PointerEvent as ReactPointerEvent,
  ReactElement,
  SetStateAction,
  WheelEvent as ReactWheelEvent
} from 'react'
import type {
  AppEnvironment,
  AuthAccount,
  AuthSession,
  CloudDriveItem,
  DriveAccountUsage,
  DriveFolderCompareDifference,
  DriveFolderCompareEndpoint,
  DriveFolderCompareResult,
  DriveFolderReconcilePriority,
  DriveTransferListResult,
  DriveTransferSummary,
  DriveTransferTask,
  DriveIndexStatus,
  DriveFolderRef,
  DriveThumbnailSize,
  DriveSettings,
  GraphActivityEvent,
  MicrosoftAuthSettingsSource,
  TransferDriveItemRef,
  TransferSettings
} from '@shared/types'
import './styles.css'

type EnvironmentState =
  | { status: 'loading' }
  | { status: 'ready'; environment: AppEnvironment }
  | { status: 'error'; message: string }

type SessionState =
  | { status: 'loading' }
  | { status: 'ready'; session: AuthSession }
  | { status: 'error'; message: string }

type DriveState =
  | { status: 'idle'; items: CloudDriveItem[]; nextLink?: string }
  | { status: 'loading'; items: CloudDriveItem[]; nextLink?: string }
  | { status: 'ready'; items: CloudDriveItem[]; nextLink?: string }
  | { status: 'error'; items: CloudDriveItem[]; nextLink?: string; message: string }

type CachedDriveFolder = {
  items: CloudDriveItem[]
  nextLink?: string
}

type LoadDriveFolderOptions = {
  append?: boolean
  accountId?: string | null
  forceRefresh?: boolean
  nextLink?: string
  tabId?: string
}

type IndexState =
  | { status: 'idle' }
  | { status: 'syncing'; index?: DriveIndexStatus }
  | { status: 'ready'; index: DriveIndexStatus }
  | { status: 'error'; message: string }

type FileOperationState =
  | { status: 'idle' }
  | { status: 'working'; message: string }
  | { status: 'success'; message: string }
  | { status: 'error'; message: string }

type AccountUsageState =
  | { status: 'idle'; usages: Record<string, DriveAccountUsage> }
  | { status: 'loading'; usages: Record<string, DriveAccountUsage> }
  | { status: 'ready'; usages: Record<string, DriveAccountUsage> }
  | { status: 'error'; usages: Record<string, DriveAccountUsage>; message: string }

type TransferViewerState =
  | { status: 'closed'; offset: number; limit: number; result: DriveTransferListResult | null }
  | { status: 'loading'; offset: number; limit: number; result: DriveTransferListResult | null }
  | { status: 'ready'; offset: number; limit: number; result: DriveTransferListResult }
  | { status: 'error'; offset: number; limit: number; result: DriveTransferListResult | null; message: string }

type DrivePreviewState = {
  tabId: string
  accountId: string | null
  item: CloudDriveItem
}

type DriveSelectionMode = 'replace' | 'toggle' | 'range' | 'context'

type DriveClipboardMode = 'copy' | 'cut'

type DriveClipboard = {
  mode: DriveClipboardMode
  items: CloudDriveItem[]
  sourceAccountId: string
  sourceFolderId: string | null
}

type DriveDragPayload = {
  itemIds: string[]
  sourceAccountId: string | null
  items: TransferDriveItemRef[]
}

type BoxSelectionState = {
  originX: number
  originY: number
  currentX: number
  currentY: number
}

type BoxSelectionSession = {
  originX: number
  originY: number
  baseSelectedItemIds: string[]
  isAdditive: boolean
  hasMoved: boolean
}

type DriveSortField = 'name' | 'modified' | 'type' | 'size'

type DriveSortDirection = 'asc' | 'desc'

type DriveSortOptions = {
  field: DriveSortField | null
  direction: DriveSortDirection
  foldersFirst: boolean
}

type DriveViewMode = 'details' | 'large-icons'

type ThumbnailPreviewState =
  | { status: 'idle' | 'loading' | 'missing' | 'error' }
  | { status: 'ready'; url: string; width?: number; height?: number }

type PreviewTransform = {
  scale: number
  x: number
  y: number
}

type PreviewPanSession = {
  pointerId: number
  originX: number
  originY: number
  startX: number
  startY: number
}

type DriveColumnKey = 'name' | 'modified' | 'type' | 'size'

type DriveColumnWidths = Record<DriveColumnKey, number>

type ColumnResizeSession = {
  columnKey: DriveColumnKey
  originX: number
  originWidth: number
  initialWidths: DriveColumnWidths
  bodyCursor: string
  bodyUserSelect: string
}

type PaneResizeSession = {
  leftTabId: string
  rightTabId: string
  originX: number
  leftWidth: number
  rightWidth: number
  leftSize: number
  rightSize: number
  bodyCursor: string
  bodyUserSelect: string
}

type DriveTab = {
  id: string
  accountId: string | null
  folderPath: DriveFolderRef[]
  driveState: DriveState
  indexState: IndexState
  sortOptions: DriveSortOptions
  viewMode: DriveViewMode
  columnWidths: DriveColumnWidths
  paneSize: number
}

type StateUpdate<T> = T | ((current: T) => T)

type ContextMenuState = {
  x: number
  y: number
  targetItemId?: string | null
}

type FolderCompareEndpointView = DriveFolderCompareEndpoint & {
  accountLabel: string
  pathLabel: string
}

type FolderCompareState =
  | { status: 'closed'; source: FolderCompareEndpointView | null; target: FolderCompareEndpointView | null; result: DriveFolderCompareResult | null }
  | { status: 'loading'; source: FolderCompareEndpointView; target: FolderCompareEndpointView; result: DriveFolderCompareResult | null }
  | { status: 'ready'; source: FolderCompareEndpointView; target: FolderCompareEndpointView; result: DriveFolderCompareResult }
  | { status: 'error'; source: FolderCompareEndpointView; target: FolderCompareEndpointView; result: DriveFolderCompareResult | null; message: string }

type DriveNameDialogState = {
  title: string
  label: string
  initialValue: string
  confirmLabel: string
}

const DRIVE_ITEM_DRAG_TYPE = 'application/x-onedrive-manager-items'

const rootFolder: DriveFolderRef = { id: null, name: '내 OneDrive' }

const platformLabels: Record<string, string> = {
  darwin: 'macOS',
  win32: 'Windows',
  linux: 'Linux'
}

const itemTypeLabels: Record<CloudDriveItem['type'], string> = {
  file: '파일',
  folder: '폴더',
  package: '패키지'
}

const driveSortFieldLabels: Record<DriveSortField, string> = {
  name: '이름',
  modified: '수정한 날짜',
  type: '유형',
  size: '크기'
}

const driveItemTypeOrder: Record<CloudDriveItem['type'], number> = {
  folder: 0,
  package: 1,
  file: 2
}

const driveItemNameCollator = new Intl.Collator('ko-KR', {
  numeric: true,
  sensitivity: 'base'
})

const defaultDriveSortOptions: DriveSortOptions = {
  field: 'name',
  direction: 'asc',
  foldersFirst: true
}

const driveColumnLabels: Record<DriveColumnKey, string> = {
  name: '이름',
  modified: '수정한 날짜',
  type: '유형',
  size: '크기'
}

const driveViewModeLabels: Record<DriveViewMode, string> = {
  details: '목록',
  'large-icons': '큰 아이콘'
}

const imageFileExtensions = new Set(['avif', 'bmp', 'gif', 'heic', 'heif', 'jpeg', 'jpg', 'png', 'tif', 'tiff', 'webp'])
const videoFileExtensions = new Set(['avi', 'm2ts', 'm4v', 'mkv', 'mov', 'mp4', 'mpeg', 'mpg', 'webm', 'wmv'])
const driveFileIconCache = new Map<string, string | null>()
const driveFileIconPromises = new Map<string, Promise<string | null>>()

const defaultDriveColumnWidths: DriveColumnWidths = {
  name: 280,
  modified: 150,
  type: 160,
  size: 90
}

const minDriveColumnWidths: DriveColumnWidths = {
  name: 160,
  modified: 120,
  type: 100,
  size: 72
}

const maxDriveColumnWidths: DriveColumnWidths = {
  name: 720,
  modified: 320,
  type: 320,
  size: 220
}

const DEFAULT_DRIVE_PANE_SIZE = 1
const MIN_DRIVE_PANE_WIDTH = 320
const PANE_RESIZE_HANDLE_WIDTH = 6
const TRANSFER_VIEW_PAGE_LIMIT = 100
const TRANSFER_UI_REFRESH_INTERVAL_MS = 3_000
const FOLDER_COMPARE_PAGE_SIZE = 200
const DRIVE_PREVIEW_THUMBNAIL_MIN_WIDTH = 640
const DRIVE_PREVIEW_THUMBNAIL_MIN_HEIGHT = 480
const DRIVE_PREVIEW_THUMBNAIL_MAX_WIDTH = 1280
const DRIVE_PREVIEW_THUMBNAIL_MAX_HEIGHT = 960
const DRIVE_PREVIEW_THUMBNAIL_MAX_PIXEL_RATIO = 1.5
const DRIVE_PREVIEW_THUMBNAIL_WAIT_MS = 7_000
const DRIVE_PREVIEW_ZOOM_MIN = 1
const DRIVE_PREVIEW_ZOOM_MAX = 5
const DRIVE_PREVIEW_ZOOM_STEP = 0.25
const DEFAULT_TRANSFER_SETTINGS: TransferSettings = {
  maxConcurrentTransfers: 4,
  minConcurrentTransfers: 1,
  maxAllowedConcurrentTransfers: 64
}
const DEFAULT_DRIVE_SETTINGS: DriveSettings = {
  indexMode: 'automatic'
}
const MAX_GRAPH_ACTIVITY_EVENTS = 200

const PLACEHOLDER_CLIENT_ID = '00000000-0000-0000-0000-000000000000'
const CLIENT_ID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
const DEFAULT_TENANT_ID = 'consumers'

async function listTransferPageSnapshot(offset: number, limit: number): Promise<DriveTransferListResult> {
  if (typeof window.oneDriveManager.listTransferPage === 'function') {
    return window.oneDriveManager.listTransferPage({ offset, limit })
  }

  const tasks = await window.oneDriveManager.listTransfers()
  const normalizedOffset = Math.max(0, Math.floor(offset))
  const normalizedLimit = Math.max(1, Math.floor(limit))

  return {
    tasks: tasks.slice(normalizedOffset, normalizedOffset + normalizedLimit),
    summary: createTransferSummaryFromTasks(tasks),
    offset: normalizedOffset,
    limit: normalizedLimit,
    totalTaskCount: tasks.length
  }
}

function createTransferSummaryFromTasks(tasks: DriveTransferTask[]): DriveTransferSummary {
  const summary: DriveTransferSummary = {
    totalCount: 0,
    activeCount: 0,
    completedCount: 0,
    queuedCount: 0,
    runningCount: 0,
    pausedCount: 0,
    retryingCount: 0,
    failedCount: 0,
    totalBytes: 0,
    transferredBytes: 0,
    bytesPerSecond: 0
  }

  for (const task of tasks) {
    summary.totalCount += 1
    summary.totalBytes += Math.max(task.totalBytes, 0)
    summary.transferredBytes += Math.max(task.transferredBytes, 0)

    if (task.status === 'completed') {
      summary.completedCount += 1
      continue
    }

    summary.activeCount += 1
    summary.bytesPerSecond += task.status === 'running' ? Math.max(task.bytesPerSecond ?? 0, 0) : 0

    if (task.status === 'queued') {
      summary.queuedCount += 1
    } else if (task.status === 'running') {
      summary.runningCount += 1
    } else if (task.status === 'paused') {
      summary.pausedCount += 1
    } else if (task.status === 'retrying') {
      summary.retryingCount += 1
    } else if (task.status === 'failed') {
      summary.failedCount += 1
    }
  }

  return summary
}

export function App(): ReactElement {
  const initialTabRef = useRef<DriveTab | null>(null)

  if (!initialTabRef.current) {
    initialTabRef.current = createDriveTab(null)
  }

  const [environmentState, setEnvironmentState] = useState<EnvironmentState>({ status: 'loading' })
  const [sessionState, setSessionState] = useState<SessionState>({ status: 'loading' })
  const [tabs, setTabs] = useState<DriveTab[]>(() => [initialTabRef.current as DriveTab])
  const [activeTabId, setActiveTabId] = useState(() => (initialTabRef.current as DriveTab).id)
  const [connectionError, setConnectionError] = useState<string | null>(null)
  const [settingsError, setSettingsError] = useState<string | null>(null)
  const [settingsMessage, setSettingsMessage] = useState<string | null>(null)
  const [settingsForm, setSettingsForm] = useState({ clientId: '', tenantId: DEFAULT_TENANT_ID })
  const [fileOperationState, setFileOperationState] = useState<FileOperationState>({ status: 'idle' })
  const [accountUsageState, setAccountUsageState] = useState<AccountUsageState>({ status: 'idle', usages: {} })
  const [transferTasks, setTransferTasks] = useState<DriveTransferTask[]>([])
  const [transferSummary, setTransferSummary] = useState<DriveTransferSummary | null>(null)
  const [transferSettings, setTransferSettings] = useState<TransferSettings>(DEFAULT_TRANSFER_SETTINGS)
  const [driveSettings, setDriveSettings] = useState<DriveSettings>(DEFAULT_DRIVE_SETTINGS)
  const [graphActivityEvents, setGraphActivityEvents] = useState<GraphActivityEvent[]>([])
  const [isGraphActivityLogOpen, setIsGraphActivityLogOpen] = useState(false)
  const [transferViewerState, setTransferViewerState] = useState<TransferViewerState>({
    status: 'closed',
    offset: 0,
    limit: TRANSFER_VIEW_PAGE_LIMIT,
    result: null
  })
  const [selectedItemIds, setSelectedItemIds] = useState<string[]>([])
  const [selectionAnchorItemId, setSelectionAnchorItemId] = useState<string | null>(null)
  const [driveClipboard, setDriveClipboard] = useState<DriveClipboard | null>(null)
  const [folderCompareSource, setFolderCompareSource] = useState<FolderCompareEndpointView | null>(null)
  const [folderCompareState, setFolderCompareState] = useState<FolderCompareState>({
    status: 'closed',
    source: null,
    target: null,
    result: null
  })
  const [drivePreview, setDrivePreview] = useState<DrivePreviewState | null>(null)
  const [contextMenu, setContextMenu] = useState<ContextMenuState | null>(null)
  const [driveNameDialog, setDriveNameDialog] = useState<DriveNameDialogState | null>(null)
  const [dragOverFolderId, setDragOverFolderId] = useState<string | null>(null)
  const [isConnecting, setIsConnecting] = useState(false)
  const [isSavingSettings, setIsSavingSettings] = useState(false)
  const [isResettingSettings, setIsResettingSettings] = useState(false)
  const folderCacheRef = useRef(new Map<string, CachedDriveFolder>())
  const driveNameDialogResolveRef = useRef<((value: string | null) => void) | null>(null)
  const paneResizeSessionRef = useRef<PaneResizeSession | null>(null)
  const session = sessionState.status === 'ready' ? sessionState.session : null
  const accountUsageKey = session?.accounts.map((account) => account.homeAccountId).join('|') ?? ''
  const activeTab = tabs.find((tab) => tab.id === activeTabId) ?? tabs[0] ?? initialTabRef.current
  const driveState = activeTab?.driveState ?? createIdleDriveState()
  const folderPath = activeTab?.folderPath ?? [rootFolder]
  const indexState = activeTab?.indexState ?? { status: 'idle' }

  const loadDriveFolder = useCallback(async (path: DriveFolderRef[], options: LoadDriveFolderOptions = {}): Promise<void> => {
    const { append = false, forceRefresh = false, nextLink } = options
    const tabId = options.tabId ?? activeTabId
    const accountId = options.accountId ?? getDriveTabAccountId(tabs, tabId) ?? session?.activeAccountId ?? null
    const folder = path.at(-1) ?? rootFolder
    const currentFolder = tabs.find((tab) => tab.id === tabId)?.folderPath.at(-1) ?? rootFolder
    const isSameFolderRequest = currentFolder.id === folder.id
    const cacheKey = getFolderCacheKey(accountId, folder)
    const cachedFolder = folderCacheRef.current.get(cacheKey)

    if (forceRefresh) {
      clearFolderCacheForAccount(folderCacheRef.current, accountId)
    }

    if (!append && !forceRefresh && !nextLink && cachedFolder) {
      setFolderPathForTab(setTabs, tabId, path)

      if (tabId === activeTabId) {
        clearDriveSelection()
      }

      setDriveStateForTab(setTabs, tabId, {
        status: 'ready',
        items: cachedFolder.items,
        nextLink: cachedFolder.nextLink
      })
      return
    }

    setDriveStateForTab(setTabs, tabId, (currentState) => ({
      status: 'loading',
      items: append ? currentState.items : (isSameFolderRequest ? currentState.items : cachedFolder?.items ?? []),
      nextLink: append ? currentState.nextLink : cachedFolder?.nextLink
    }))

    try {
      const result = await window.oneDriveManager.listDriveChildren({
        folderId: folder.id,
        nextLink,
        forceRefresh
      })
      const baseItems = append ? (folderCacheRef.current.get(cacheKey)?.items ?? []) : []
      const items = append ? [...baseItems, ...result.items] : result.items

      folderCacheRef.current.set(cacheKey, {
        items,
        nextLink: result.nextLink
      })

      setFolderPathForTab(setTabs, tabId, path)

      if (tabId === activeTabId) {
        clearDriveSelection()
      }

      setDriveStateForTab(setTabs, tabId, {
        status: 'ready',
        items,
        nextLink: result.nextLink
      })
    } catch (error) {
      setDriveStateForTab(setTabs, tabId, (currentState) => ({
        status: 'error',
        items: currentState.items,
        nextLink: currentState.nextLink,
        message: error instanceof Error ? error.message : 'OneDrive 파일 목록을 가져오지 못했습니다.'
      }))
    }
  }, [activeTabId, session?.activeAccountId, tabs])

  const loadSession = useCallback(async (): Promise<AuthSession | null> => {
    setSessionState({ status: 'loading' })

    try {
      const session = await window.oneDriveManager.getAuthSession()
      setSessionState({ status: 'ready', session })
      setSettingsForm({
        clientId: session.settings.clientId,
        tenantId: session.settings.tenantId
      })
      return session
    } catch (error) {
      setSessionState({
        status: 'error',
        message: error instanceof Error ? error.message : '계정 상태를 확인하지 못했습니다.'
      })
      return null
    }
  }, [])

  const loadTransferSettings = useCallback(async (): Promise<void> => {
    if (typeof window.oneDriveManager.getTransferSettings !== 'function') {
      return
    }

    try {
      setTransferSettings(await window.oneDriveManager.getTransferSettings())
    } catch {
      setTransferSettings(DEFAULT_TRANSFER_SETTINGS)
    }
  }, [])

  const loadDriveSettings = useCallback(async (): Promise<void> => {
    if (typeof window.oneDriveManager.getDriveSettings !== 'function') {
      return
    }

    try {
      setDriveSettings(await window.oneDriveManager.getDriveSettings())
    } catch {
      setDriveSettings(DEFAULT_DRIVE_SETTINGS)
    }
  }, [])

  const refreshEnvironment = useCallback(async (): Promise<void> => {
    setEnvironmentState({ status: 'loading' })

    try {
      const environment = await window.oneDriveManager.getEnvironment()
      setEnvironmentState({ status: 'ready', environment })
    } catch (error) {
      setEnvironmentState({
        status: 'error',
        message: error instanceof Error ? error.message : '환경 정보를 확인하지 못했습니다.'
      })
    }
  }, [])

  const warmNavigationIndex = useCallback(async (forceRefresh = false, tabId = activeTabId, isManualRequest = false): Promise<boolean> => {
    if (driveSettings.indexMode === 'manual' && !isManualRequest) {
      return true
    }

    setIndexStateForTab(setTabs, tabId, (currentState) => {
      if (currentState.status === 'ready') {
        return { status: 'syncing', index: currentState.index }
      }

      if (currentState.status === 'syncing') {
        return currentState
      }

      return { status: 'syncing' }
    })

    try {
      const index = await window.oneDriveManager.warmDriveIndex({ forceRefresh })
      setIndexStateForTab(setTabs, tabId, index.isSyncing ? { status: 'syncing', index } : { status: 'ready', index })
      return true
    } catch (error) {
      setIndexStateForTab(setTabs, tabId, {
        status: 'error',
        message: error instanceof Error ? error.message : 'OneDrive 탐색 인덱스를 구성하지 못했습니다.'
      })
      return false
    }
  }, [activeTabId, driveSettings.indexMode])

  const initialize = useCallback(async (): Promise<void> => {
    await refreshEnvironment()
    await loadTransferSettings()
    await loadDriveSettings()
    const session = await loadSession()

    if (session?.isAuthenticated) {
      const accountId = session.activeAccountId ?? null

      setAccountForTab(setTabs, activeTabId, accountId)
      void warmNavigationIndex(false, activeTabId)
      await loadDriveFolder([rootFolder], { accountId, tabId: activeTabId })
    }
  }, [loadDriveFolder, loadDriveSettings, loadSession, loadTransferSettings, refreshEnvironment, warmNavigationIndex])

  useEffect(() => {
    void initialize()
  }, [])

  useEffect(() => {
    if (indexState.status !== 'syncing') {
      return
    }

    const refreshTimer = window.setTimeout(() => {
      void warmNavigationIndex(false)
    }, 1500)

    return () => window.clearTimeout(refreshTimer)
  }, [indexState, warmNavigationIndex])

  useEffect(() => {
    const availableItemIds = new Set(driveState.items.map((item) => item.id))

    setSelectedItemIds((currentItemIds) => currentItemIds.filter((itemId) => availableItemIds.has(itemId)))
    setSelectionAnchorItemId((currentItemId) => (currentItemId && availableItemIds.has(currentItemId) ? currentItemId : null))
  }, [driveState.items])

  useEffect(() => {
    if (!contextMenu) {
      return
    }

    const closeMenu = (): void => setContextMenu(null)
    const closeMenuOnEscape = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        closeMenu()
      }
    }

    window.addEventListener('click', closeMenu)
    window.addEventListener('blur', closeMenu)
    window.addEventListener('resize', closeMenu)
    window.addEventListener('scroll', closeMenu, true)
    window.addEventListener('keydown', closeMenuOnEscape)

    return () => {
      window.removeEventListener('click', closeMenu)
      window.removeEventListener('blur', closeMenu)
      window.removeEventListener('resize', closeMenu)
      window.removeEventListener('scroll', closeMenu, true)
      window.removeEventListener('keydown', closeMenuOnEscape)
    }
  }, [contextMenu])

  useEffect(() => {
    if (!driveClipboard) {
      return
    }

    const cancelClipboardOnEscape = (event: KeyboardEvent): void => {
      if (event.key !== 'Escape') {
        return
      }

      setDriveClipboard(null)
      setFileOperationState({ status: 'success', message: '복사/잘라내기 준비를 취소했습니다.' })
    }

    window.addEventListener('keydown', cancelClipboardOnEscape)

    return () => window.removeEventListener('keydown', cancelClipboardOnEscape)
  }, [driveClipboard])

  useEffect(() => {
    let transferUpdateTimer: number | null = null
    let pendingTransferTasks: DriveTransferTask[] | null = null
    let lastTransferUpdateAt = 0

    function applyTransferUpdate(): void {
      if (pendingTransferTasks) {
        setTransferTasks(pendingTransferTasks)
        pendingTransferTasks = null
      }

      lastTransferUpdateAt = Date.now()
      void refreshTransferPreview()
    }

    void refreshTransferPreview()

    const dispose = window.oneDriveManager.onTransfersUpdated((tasks) => {
      pendingTransferTasks = tasks
      const elapsedMs = Date.now() - lastTransferUpdateAt

      if (elapsedMs >= TRANSFER_UI_REFRESH_INTERVAL_MS) {
        if (transferUpdateTimer) {
          window.clearTimeout(transferUpdateTimer)
          transferUpdateTimer = null
        }

        applyTransferUpdate()
        return
      }

      if (!transferUpdateTimer) {
        transferUpdateTimer = window.setTimeout(() => {
          transferUpdateTimer = null
          applyTransferUpdate()
        }, TRANSFER_UI_REFRESH_INTERVAL_MS - elapsedMs)
      }
    })

    return () => {
      if (transferUpdateTimer) {
        window.clearTimeout(transferUpdateTimer)
      }

      dispose()
    }
  }, [])

  useEffect(() => {
    if (typeof window.oneDriveManager.onGraphActivity !== 'function') {
      return undefined
    }

    return window.oneDriveManager.onGraphActivity((event) => {
      setGraphActivityEvents((currentEvents) => [event, ...currentEvents].slice(0, MAX_GRAPH_ACTIVITY_EVENTS))
    })
  }, [])

  useEffect(() => {
    if (!session?.isAuthenticated || session.accounts.length === 0) {
      setAccountUsageState({ status: 'idle', usages: {} })
      return
    }

    let didCancel = false

    setAccountUsageState((currentState) => ({ status: 'loading', usages: currentState.usages }))
    void window.oneDriveManager
      .listAccountUsage()
      .then((usages) => {
        if (didCancel) {
          return
        }

        setAccountUsageState({
          status: 'ready',
          usages: Object.fromEntries(usages.map((usage) => [usage.accountId, usage]))
        })
      })
      .catch((error) => {
        if (didCancel) {
          return
        }

        setAccountUsageState((currentState) => ({
          status: 'error',
          usages: currentState.usages,
          message: error instanceof Error ? error.message : '계정 사용량을 확인하지 못했습니다.'
        }))
      })

    return () => {
      didCancel = true
    }
  }, [session?.isAuthenticated, accountUsageKey])

  async function connectAccount(): Promise<void> {
    setIsConnecting(true)
    setConnectionError(null)
    setSettingsError(null)

    try {
      if (!(await ensureAuthSettingsBeforeLogin())) {
        return
      }

      const session = await window.oneDriveManager.connectAccount()
      const accountId = session.activeAccountId ?? null

      setSessionState({ status: 'ready', session })
      setAccountForTab(setTabs, activeTabId, accountId)
      setIndexStateForTab(setTabs, activeTabId, { status: 'idle' })
      setFolderPathForTab(setTabs, activeTabId, [rootFolder])
      setDriveStateForTab(setTabs, activeTabId, createIdleDriveState())
      setTransferTasks(await window.oneDriveManager.listTransfers())
      clearDriveSelection()
      void warmNavigationIndex(false, activeTabId)
      await loadDriveFolder([rootFolder], { accountId, tabId: activeTabId })
    } catch (error) {
      setConnectionError(error instanceof Error ? error.message : '로그인을 완료하지 못했습니다.')
    } finally {
      setIsConnecting(false)
    }
  }

  async function ensureAuthSettingsBeforeLogin(): Promise<boolean> {
    const activeSettings = sessionState.status === 'ready' ? sessionState.session.settings : null
    const nextSettings = {
      clientId: settingsForm.clientId.trim(),
      tenantId: settingsForm.tenantId.trim() || DEFAULT_TENANT_ID
    }
    const validationError = validateAuthSettings(nextSettings.clientId)

    if (validationError) {
      setSettingsError(validationError)
      return false
    }

    const isDirty =
      !activeSettings?.isConfigured ||
      activeSettings.clientId !== nextSettings.clientId ||
      activeSettings.tenantId !== nextSettings.tenantId

    if (!isDirty) {
      return true
    }

    try {
      const settings = await window.oneDriveManager.updateMicrosoftAuthSettings(nextSettings)
      folderCacheRef.current.clear()
      setSettingsForm({
        clientId: settings.clientId,
        tenantId: settings.tenantId
      })
      setSettingsMessage('로그인 설정을 저장했습니다.')
      const session = await loadSession()
      return Boolean(session?.isConfigured)
    } catch (error) {
      setSettingsError(error instanceof Error ? error.message : '로그인 설정을 저장하지 못했습니다.')
      return false
    }
  }

  async function switchAccount(accountId: string): Promise<void> {
    const currentAccountId = sessionState.status === 'ready' ? sessionState.session.activeAccountId : null

    if (!accountId || accountId === currentAccountId || isFileOperationBusy) {
      return
    }

    setConnectionError(null)
    setSettingsError(null)
    setContextMenu(null)
    setFileOperationState({ status: 'working', message: '계정 전환 중' })

    try {
      const session = await window.oneDriveManager.switchAccount({ accountId })
      const nextAccountId = session.activeAccountId ?? accountId

      setSessionState({ status: 'ready', session })
      setAccountForTab(setTabs, activeTabId, nextAccountId)
      setIndexStateForTab(setTabs, activeTabId, { status: 'idle' })
      setFolderPathForTab(setTabs, activeTabId, [rootFolder])
      setDriveStateForTab(setTabs, activeTabId, createIdleDriveState())
      clearDriveSelection()
      setTransferTasks(await window.oneDriveManager.listTransfers())

      if (session.isAuthenticated) {
        void warmNavigationIndex(false, activeTabId)
        await loadDriveFolder([rootFolder], { accountId: nextAccountId, tabId: activeTabId })
      }

      setFileOperationState({ status: 'success', message: '계정 전환 완료' })
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '계정을 전환하지 못했습니다.'
      })
    }
  }

  async function disconnectAccount(): Promise<void> {
    setConnectionError(null)

    try {
      const disconnectedAccountId = sessionState.status === 'ready' ? sessionState.session.activeAccountId ?? null : null
      const session = await window.oneDriveManager.disconnectAccount()
      const nextAccountId = session.activeAccountId ?? null

      setSessionState({ status: 'ready', session })
      setTabs((currentTabs) =>
        currentTabs.map((tab) =>
          tab.accountId === disconnectedAccountId
            ? {
                ...tab,
                accountId: nextAccountId,
                folderPath: [rootFolder],
                driveState: createIdleDriveState(),
                indexState: { status: 'idle' }
              }
            : tab
        )
      )
      clearDriveSelection()
      setDriveClipboard(null)
      setTransferTasks(await window.oneDriveManager.listTransfers())

      if (session.isAuthenticated) {
        void warmNavigationIndex(false, activeTabId)
        await loadDriveFolder([rootFolder], { accountId: nextAccountId, tabId: activeTabId })
      }
    } catch (error) {
      setConnectionError(error instanceof Error ? error.message : '로그아웃하지 못했습니다.')
    }
  }

  async function openFolder(item: CloudDriveItem): Promise<void> {
    await openFolderInTab(activeTabId, item)
  }

  async function openBreadcrumb(index: number): Promise<void> {
    await openBreadcrumbInTab(activeTabId, index)
  }

  async function openParentFolder(): Promise<void> {
    await openParentFolderInTab(activeTabId)
  }

  async function loadMore(): Promise<void> {
    await loadMoreInTab(activeTabId)
  }

  async function openFolderInTab(tabId: string, item: CloudDriveItem): Promise<void> {
    if (item.type !== 'folder') {
      return
    }

    const tab = await activateDriveTabById(tabId)

    if (!tab) {
      return
    }

    await loadDriveFolder([...tab.folderPath, { id: item.id, name: item.name }], { accountId: tab.accountId, tabId: tab.id })
  }

  async function openBreadcrumbInTab(tabId: string, index: number): Promise<void> {
    const tab = await activateDriveTabById(tabId)

    if (!tab) {
      return
    }

    await loadDriveFolder(tab.folderPath.slice(0, index + 1), { accountId: tab.accountId, tabId: tab.id })
  }

  async function openParentFolderInTab(tabId: string): Promise<void> {
    const tab = await activateDriveTabById(tabId)

    if (!tab || tab.folderPath.length <= 1) {
      return
    }

    await loadDriveFolder(tab.folderPath.slice(0, tab.folderPath.length - 1), { accountId: tab.accountId, tabId: tab.id })
  }

  async function loadMoreInTab(tabId: string): Promise<void> {
    const tab = await activateDriveTabById(tabId)

    if (tab?.driveState.nextLink) {
      await loadDriveFolder(tab.folderPath, { accountId: tab.accountId, append: true, nextLink: tab.driveState.nextLink, tabId: tab.id })
    }
  }

  function previewItemInTab(tabId: string, item: CloudDriveItem): void {
    const tab = tabs.find((candidate) => candidate.id === tabId)

    setContextMenu(null)
    setActiveTabId(tabId)
    setDrivePreview({
      tabId,
      accountId: tab?.accountId ?? session?.activeAccountId ?? null,
      item
    })
  }

  function moveDrivePreview(direction: -1 | 1): void {
    setDrivePreview((currentPreview) => {
      if (!currentPreview) {
        return null
      }

      const tab = tabs.find((candidate) => candidate.id === currentPreview.tabId)
      const previewItems = tab ? getDrivePreviewItems(tab) : [currentPreview.item]

      if (previewItems.length <= 1) {
        return currentPreview
      }

      const currentIndex = previewItems.findIndex((candidate) => candidate.id === currentPreview.item.id)
      const nextIndex = currentIndex >= 0 ? (currentIndex + direction + previewItems.length) % previewItems.length : 0
      const nextItem = previewItems[nextIndex] ?? currentPreview.item

      return {
        tabId: currentPreview.tabId,
        accountId: tab?.accountId ?? currentPreview.accountId,
        item: nextItem
      }
    })
  }

  async function renameItemInTab(tabId: string, item: CloudDriveItem): Promise<void> {
    if (await activateDriveTabById(tabId)) {
      await renameSelectedItem(item)
    }
  }

  async function deleteItemInTab(tabId: string, item: CloudDriveItem): Promise<void> {
    if (await activateDriveTabById(tabId)) {
      await deleteSelectedItems([item])
    }
  }

  async function moveItemsToFolderInTab(tabId: string, folder: CloudDriveItem, payload: DriveDragPayload): Promise<void> {
    const targetTab = await activateDriveTabById(tabId)

    if (!targetTab) {
      return
    }

    if (payload.sourceAccountId && targetTab.accountId && payload.sourceAccountId !== targetTab.accountId) {
      setContextMenu(null)
      setDragOverFolderId(null)
      setFileOperationState({ status: 'working', message: '계정 간 이동 등록 중' })

      try {
        const result = await window.oneDriveManager.transferDriveItemsBetweenAccounts({
          sourceAccountId: payload.sourceAccountId,
          targetAccountId: targetTab.accountId,
          targetParentId: folder.id,
          deleteSourceOnComplete: true,
          items: payload.items
        })

        setTransferTasks(await window.oneDriveManager.listTransfers())
        folderCacheRef.current.clear()
        await loadDriveFolder(targetTab.folderPath, { accountId: targetTab.accountId, forceRefresh: true, tabId: targetTab.id })
        void warmNavigationIndex(true, targetTab.id)
        setFileOperationState({
          status: 'success',
          message: `${result.queuedCount.toLocaleString('ko-KR')}개 파일을 계정 간 이동 큐에 등록했습니다.`
        })
      } catch (error) {
        setFileOperationState({
          status: 'error',
          message: error instanceof Error ? error.message : '계정 간 이동을 등록하지 못했습니다.'
        })
      }

      return
    }

    if (payload.itemIds.length > 0) {
      setContextMenu(null)
      setDragOverFolderId(null)
      setFileOperationState({ status: 'working', message: '항목 이동 중' })

      try {
        const movedItems = await window.oneDriveManager.moveDriveItems({
          itemIds: payload.itemIds,
          parentId: folder.id
        })

        folderCacheRef.current.clear()
        clearDriveSelection()
        await loadDriveFolder(targetTab.folderPath, { accountId: targetTab.accountId, forceRefresh: true, tabId: targetTab.id })
        void warmNavigationIndex(true, targetTab.id)
        setFileOperationState({
          status: 'success',
          message: `${movedItems.length.toLocaleString('ko-KR')}개 항목 이동 완료`
        })
      } catch (error) {
        setFileOperationState({
          status: 'error',
          message: error instanceof Error ? error.message : '항목을 이동하지 못했습니다.'
        })
      }
    }
  }

  async function uploadDroppedFilesInTab(tabId: string, parentId: string | null, files: File[]): Promise<void> {
    if (await activateDriveTabById(tabId)) {
      await uploadDroppedFiles(parentId, files)
    }
  }

  async function createNewDriveTab(): Promise<void> {
    const accountId = session?.activeAccountId ?? null
    const nextTab = createDriveTab(accountId)

    setTabs((currentTabs) => [...currentTabs, nextTab])
    setActiveTabId(nextTab.id)
    clearDriveSelection()

    if (session?.isAuthenticated) {
      void warmNavigationIndex(false, nextTab.id)
      await loadDriveFolder([rootFolder], { accountId, tabId: nextTab.id })
    }
  }

  async function activateDriveTab(tabId: string): Promise<void> {
    await activateDriveTabById(tabId)
  }

  async function activateDriveTabById(tabId: string): Promise<DriveTab | null> {
    const nextTab = tabs.find((tab) => tab.id === tabId)

    if (!nextTab || isFileOperationBusy) {
      return null
    }

    if (nextTab.id === activeTabId) {
      return nextTab
    }

    return activateDriveTabEntry(nextTab)
  }

  async function activateDriveTabEntry(nextTab: DriveTab): Promise<DriveTab | null> {
    setActiveTabId(nextTab.id)
    clearDriveSelection()

    if (!nextTab.accountId || nextTab.accountId === session?.activeAccountId) {
      return nextTab
    }

    setFileOperationState({ status: 'working', message: '탭 계정 전환 중' })

    try {
      const nextSession = await window.oneDriveManager.switchAccount({ accountId: nextTab.accountId })

      setSessionState({ status: 'ready', session: nextSession })
      setTransferTasks(await window.oneDriveManager.listTransfers())
      void warmNavigationIndex(false, nextTab.id)
      setFileOperationState({ status: 'success', message: '탭 전환 완료' })
      return nextTab
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '탭 계정으로 전환하지 못했습니다.'
      })
      return null
    }
  }

  async function closeDriveTab(tabId: string): Promise<void> {
    if (tabs.length <= 1 || isFileOperationBusy) {
      return
    }

    const closedTabIndex = tabs.findIndex((tab) => tab.id === tabId)

    if (closedTabIndex < 0) {
      return
    }

    const nextTabs = tabs.filter((tab) => tab.id !== tabId)
    const nextActiveTab = tabId === activeTabId ? nextTabs[Math.min(closedTabIndex, nextTabs.length - 1)] : null

    setTabs(nextTabs)

    if (nextActiveTab) {
      await activateDriveTabEntry(nextActiveTab)
    }
  }

  function startDrivePaneResize(event: ReactMouseEvent<HTMLDivElement>, leftTabId: string, rightTabId: string): void {
    const leftPane = event.currentTarget.previousElementSibling
    const rightPane = event.currentTarget.nextElementSibling

    if (!(leftPane instanceof HTMLElement) || !(rightPane instanceof HTMLElement)) {
      return
    }

    const leftTab = tabs.find((tab) => tab.id === leftTabId)
    const rightTab = tabs.find((tab) => tab.id === rightTabId)

    if (!leftTab || !rightTab) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    paneResizeSessionRef.current = {
      leftTabId,
      rightTabId,
      originX: event.clientX,
      leftWidth: leftPane.getBoundingClientRect().width,
      rightWidth: rightPane.getBoundingClientRect().width,
      leftSize: leftTab.paneSize,
      rightSize: rightTab.paneSize,
      bodyCursor: document.body.style.cursor,
      bodyUserSelect: document.body.style.userSelect
    }
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    window.addEventListener('mousemove', handleDrivePaneResizeMouseMove)
    window.addEventListener('mouseup', handleDrivePaneResizeMouseUp, { once: true })
  }

  function handleDrivePaneResizeMouseMove(event: MouseEvent): void {
    const session = paneResizeSessionRef.current

    if (!session) {
      return
    }

    const totalWidth = session.leftWidth + session.rightWidth

    if (totalWidth <= MIN_DRIVE_PANE_WIDTH * 2) {
      return
    }

    const nextLeftWidth = clampNumber(session.leftWidth + event.clientX - session.originX, MIN_DRIVE_PANE_WIDTH, totalWidth - MIN_DRIVE_PANE_WIDTH)
    const totalSize = session.leftSize + session.rightSize
    const nextLeftSize = Math.max(0.2, totalSize * (nextLeftWidth / totalWidth))
    const nextRightSize = Math.max(0.2, totalSize - nextLeftSize)

    setTabs((currentTabs) =>
      currentTabs.map((tab) => {
        if (tab.id === session.leftTabId) {
          return { ...tab, paneSize: nextLeftSize }
        }

        if (tab.id === session.rightTabId) {
          return { ...tab, paneSize: nextRightSize }
        }

        return tab
      })
    )
  }

  function handleDrivePaneResizeMouseUp(): void {
    const session = paneResizeSessionRef.current

    window.removeEventListener('mousemove', handleDrivePaneResizeMouseMove)
    paneResizeSessionRef.current = null

    if (session) {
      document.body.style.cursor = session.bodyCursor
      document.body.style.userSelect = session.bodyUserSelect
    }
  }

  async function resumeTransfers(taskId?: string): Promise<void> {
    setFileOperationState({ status: 'working', message: '전송 재개 중' })

    try {
      const tasks = await window.oneDriveManager.resumeTransfers(taskId)

      setTransferTasks(tasks)
      setFileOperationState({ status: 'success', message: '전송 재개 완료' })
      folderCacheRef.current.clear()
      await loadDriveFolder(folderPath, { forceRefresh: true })
      void warmNavigationIndex(true)
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '전송을 재개하지 못했습니다.'
      })
    }
  }

  async function stopTransfer(taskId: string): Promise<void> {
    try {
      const tasks = await window.oneDriveManager.stopTransfer(taskId)

      setTransferTasks(tasks)
      void refreshTransferPreview()
      setFileOperationState({ status: 'success', message: '전송을 중지했습니다.' })
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '전송을 중지하지 못했습니다.'
      })
    }
  }

  async function deleteTransfer(taskId?: string): Promise<void> {
    const confirmMessage = taskId
      ? '이 전송 기록과 이어받기 데이터를 삭제할까요?'
      : '모든 전송 기록과 이어받기 데이터를 삭제할까요?'

    if (!window.confirm(confirmMessage)) {
      return
    }

    try {
      const tasks = await window.oneDriveManager.deleteTransfer(taskId)

      setTransferTasks(tasks)
      void refreshTransferPreview()
      if (transferViewerState.status !== 'closed') {
        void loadTransferViewerPage(transferViewerState.offset)
      }
      setFileOperationState({ status: 'success', message: '전송 기록을 삭제했습니다.' })
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '전송 기록을 삭제하지 못했습니다.'
      })
    }
  }

  async function updateTransferConcurrency(maxConcurrentTransfers: number): Promise<void> {
    const normalizedValue = Math.min(
      transferSettings.maxAllowedConcurrentTransfers,
      Math.max(transferSettings.minConcurrentTransfers, Math.floor(maxConcurrentTransfers))
    )

    setTransferSettings((currentSettings) => ({
      ...currentSettings,
      maxConcurrentTransfers: normalizedValue
    }))

    if (typeof window.oneDriveManager.updateTransferSettings !== 'function') {
      return
    }

    try {
      const settings = await window.oneDriveManager.updateTransferSettings({ maxConcurrentTransfers: normalizedValue })

      setTransferSettings(settings)
      void refreshTransferPreview()
    } catch (error) {
      setTransferSettings((currentSettings) => ({
        ...currentSettings,
        maxConcurrentTransfers: transferSettings.maxConcurrentTransfers
      }))
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '전송 설정을 저장하지 못했습니다.'
      })
    }
  }

  async function refreshTransferPreview(): Promise<void> {
    try {
      const result = await listTransferPageSnapshot(0, TRANSFER_VIEW_PAGE_LIMIT)

      setTransferTasks(result.tasks)
      setTransferSummary(result.summary)
    } catch {
      try {
        setTransferTasks(await window.oneDriveManager.listTransfers())
      } catch {
        // Transfer preview is best-effort; operation handlers surface actionable errors.
      }
    }
  }

  async function openTransferViewer(): Promise<void> {
    setTransferViewerState((currentState) => ({
      status: 'loading',
      offset: currentState.offset,
      limit: currentState.limit,
      result: currentState.result
    }))
    await loadTransferViewerPage(transferViewerState.offset)
  }

  function closeTransferViewer(): void {
    setTransferViewerState((currentState) => ({
      status: 'closed',
      offset: currentState.offset,
      limit: currentState.limit,
      result: currentState.result
    }))
  }

  function requestDriveName(dialog: DriveNameDialogState): Promise<string | null> {
    driveNameDialogResolveRef.current?.(null)

    return new Promise((resolve) => {
      driveNameDialogResolveRef.current = resolve
      setDriveNameDialog(dialog)
    })
  }

  function resolveDriveNameDialog(value: string | null): void {
    const resolve = driveNameDialogResolveRef.current

    driveNameDialogResolveRef.current = null
    setDriveNameDialog(null)
    resolve?.(value)
  }

  async function loadTransferViewerPage(offset: number): Promise<void> {
    const normalizedOffset = Math.max(0, offset)

    setTransferViewerState((currentState) => ({
      status: 'loading',
      offset: normalizedOffset,
      limit: currentState.limit,
      result: currentState.result
    }))

    try {
      const result = await listTransferPageSnapshot(normalizedOffset, TRANSFER_VIEW_PAGE_LIMIT)

      setTransferSummary(result.summary)
      setTransferViewerState({
        status: 'ready',
        offset: result.offset,
        limit: result.limit,
        result
      })
    } catch (error) {
      setTransferViewerState((currentState) => ({
        status: 'error',
        offset: normalizedOffset,
        limit: currentState.limit,
        result: currentState.result,
        message: error instanceof Error ? error.message : '전송 목록을 가져오지 못했습니다.'
      }))
    }
  }

  async function uploadFiles(): Promise<void> {
    if (!canUseDrive || fileOperationState.status === 'working') {
      return
    }

    const currentFolder = folderPath.at(-1) ?? rootFolder

    setContextMenu(null)
    setFileOperationState({ status: 'working', message: '업로드 중' })

    try {
      const result = await window.oneDriveManager.uploadDriveItems({ parentId: currentFolder.id })

      if (result.cancelled) {
        setFileOperationState({ status: 'idle' })
        return
      }

      folderCacheRef.current.clear()
      await loadDriveFolder(folderPath, { forceRefresh: true })
      void warmNavigationIndex(true)
      setFileOperationState({
        status: 'success',
        message: '업로드 작업을 큐에 등록했습니다.'
      })
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '파일을 업로드하지 못했습니다.'
      })
    }
  }

  async function createFolder(): Promise<void> {
    if (!canUseDrive || fileOperationState.status === 'working') {
      return
    }

    if (typeof window.oneDriveManager.createDriveFolder !== 'function') {
      setFileOperationState({ status: 'error', message: '앱을 다시 실행한 뒤 폴더 생성을 사용할 수 있습니다.' })
      return
    }

    setContextMenu(null)
    const currentFolder = folderPath.at(-1) ?? rootFolder
    const nextName = await requestDriveName({
      title: '새 폴더',
      label: '폴더 이름',
      initialValue: getAvailableNewFolderName(driveState.items),
      confirmLabel: '만들기'
    })

    if (!nextName) {
      return
    }

    const validationError = validateDriveItemName(nextName)

    if (validationError) {
      setContextMenu(null)
      setFileOperationState({ status: 'error', message: validationError })
      return
    }

    setFileOperationState({ status: 'working', message: '폴더 생성 중' })

    try {
      const createdFolder = await window.oneDriveManager.createDriveFolder({
        parentId: currentFolder.id,
        name: nextName
      })

      folderCacheRef.current.clear()
      await loadDriveFolder(folderPath, { forceRefresh: true })
      void warmNavigationIndex(true)
      setFileOperationState({ status: 'success', message: `폴더 생성 완료: ${createdFolder.name}` })
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '폴더를 만들지 못했습니다.'
      })
    }
  }

  async function uploadDroppedFiles(parentId: string | null, files: File[]): Promise<void> {
    if (!canUseDrive || fileOperationState.status === 'working' || files.length === 0) {
      return
    }

    const paths = window.oneDriveManager.getDroppedFilePaths(files)

    if (paths.length === 0) {
      setFileOperationState({ status: 'error', message: '드롭한 파일 경로를 확인하지 못했습니다.' })
      return
    }

    setContextMenu(null)
    setDragOverFolderId(null)
    setFileOperationState({ status: 'working', message: '드롭한 파일 업로드 등록 중' })

    try {
      await window.oneDriveManager.uploadDroppedItems({ parentId, paths })
      setFileOperationState({
        status: 'success',
        message: `${paths.length.toLocaleString('ko-KR')}개 파일/폴더 업로드 작업을 큐에 등록했습니다.`
      })
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '드롭한 파일을 업로드 큐에 등록하지 못했습니다.'
      })
    }
  }

  function copySelectedItemsToClipboard(mode: DriveClipboardMode, items = selectedItems): void {
    const sourceAccountId = session?.activeAccountId

    if (!canUseDrive || !sourceAccountId || items.length === 0 || fileOperationState.status === 'working') {
      return
    }

    setContextMenu(null)
    setDriveClipboard({
      mode,
      items,
      sourceAccountId,
      sourceFolderId: currentFolder.id
    })
    setFileOperationState({
      status: 'success',
      message:
        mode === 'copy'
          ? `${items.length.toLocaleString('ko-KR')}개 항목 복사 준비 완료`
          : `${items.length.toLocaleString('ko-KR')}개 항목 잘라내기 준비 완료`
    })
  }

  async function pasteClipboardItems(targetParentId: string | null): Promise<void> {
    const targetAccountId = session?.activeAccountId

    if (!canUseDrive || !targetAccountId || !driveClipboard || driveClipboard.items.length === 0 || fileOperationState.status === 'working') {
      return
    }

    const itemIds = driveClipboard.items.map((item) => item.id)
    const isCrossAccountPaste = driveClipboard.sourceAccountId !== targetAccountId

    setContextMenu(null)

    if (isCrossAccountPaste) {
      setFileOperationState({
        status: 'working',
        message: driveClipboard.mode === 'cut' ? '계정 간 이동 등록 중' : '계정 간 복사 등록 중'
      })

      try {
        const result = await window.oneDriveManager.transferDriveItemsBetweenAccounts({
          sourceAccountId: driveClipboard.sourceAccountId,
          targetAccountId,
          targetParentId,
          deleteSourceOnComplete: driveClipboard.mode === 'cut',
          items: driveClipboard.items.map((item) => ({
            itemId: item.id,
            name: item.name,
            type: item.type,
            size: item.size
          }))
        })

        setTransferTasks(await window.oneDriveManager.listTransfers())
        folderCacheRef.current.clear()
        await loadDriveFolder(folderPath, { forceRefresh: true })
        void warmNavigationIndex(true)
        setDriveClipboard(driveClipboard.mode === 'cut' ? null : driveClipboard)
        setFileOperationState({
          status: 'success',
          message:
            result.queuedCount > 0
              ? `${result.queuedCount.toLocaleString('ko-KR')}개 파일을 계정 간 전송 큐에 등록했습니다.`
              : '폴더 구조를 대상 계정에 만들었습니다.'
        })
      } catch (error) {
        setFileOperationState({
          status: 'error',
          message: error instanceof Error ? error.message : '계정 간 붙여넣기를 등록하지 못했습니다.'
        })
      }

      return
    }

    if (driveClipboard.mode === 'cut') {
      const movableItems = driveClipboard.items.filter((item) => item.id !== targetParentId && item.parentId !== targetParentId)

      if (movableItems.length === 0) {
        setFileOperationState({ status: 'success', message: '이미 대상 폴더에 있습니다.' })
        return
      }

      setFileOperationState({ status: 'working', message: '항목 붙여넣는 중' })

      try {
        const movedItems = await window.oneDriveManager.moveDriveItems({
          itemIds,
          parentId: targetParentId
        })
        folderCacheRef.current.clear()
        clearDriveSelection()
        setDriveClipboard(null)
        await loadDriveFolder(folderPath, { forceRefresh: true })
        void warmNavigationIndex(true)
        setFileOperationState({
          status: 'success',
          message: `${movedItems.length.toLocaleString('ko-KR')}개 항목 이동 완료`
        })
      } catch (error) {
        setFileOperationState({
          status: 'error',
          message: error instanceof Error ? error.message : '항목을 붙여넣지 못했습니다.'
        })
      }

      return
    }

    setFileOperationState({ status: 'working', message: '항목 복사 중' })

    try {
      const result = await window.oneDriveManager.copyDriveItems({
        itemIds,
        parentId: targetParentId
      })
      folderCacheRef.current.clear()
      await loadDriveFolder(folderPath, { forceRefresh: true })
      void warmNavigationIndex(true)
      setFileOperationState({
        status: 'success',
        message: result.hasPendingOperations
          ? `${driveClipboard.items.length.toLocaleString('ko-KR')}개 항목 복사 요청 완료. 큰 폴더는 완료까지 시간이 걸릴 수 있습니다.`
          : `${driveClipboard.items.length.toLocaleString('ko-KR')}개 항목 복사 완료`
      })
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '항목을 복사하지 못했습니다.'
      })
    }
  }

  function setSelectedFolderAsCompareSource(): void {
    const folder = singleSelectedItem?.type === 'folder' ? singleSelectedItem : null
    const source = folder ? createFolderCompareEndpointFromItem(folder, activeTab) : null

    if (!source) {
      setFileOperationState({ status: 'error', message: '비교 기준으로 사용할 폴더를 선택하세요.' })
      return
    }

    setContextMenu(null)
    setFolderCompareSource(source)
    setFileOperationState({ status: 'success', message: `'${source.pathLabel}' 폴더를 비교 기준으로 선택했습니다.` })
  }

  async function compareFolderWithSource(): Promise<void> {
    const source = folderCompareSource
    const target = getFolderCompareTargetEndpoint()

    if (!source) {
      setFileOperationState({ status: 'error', message: '먼저 비교 기준 폴더를 선택하세요.' })
      return
    }

    if (!target) {
      setFileOperationState({ status: 'error', message: '비교 대상 폴더를 확인하지 못했습니다.' })
      return
    }

    if (typeof window.oneDriveManager.compareDriveFolders !== 'function') {
      setFileOperationState({ status: 'error', message: '앱을 다시 실행한 뒤 폴더 비교를 사용할 수 있습니다.' })
      return
    }

    setContextMenu(null)
    setFolderCompareState({ status: 'loading', source, target, result: null })
    setFileOperationState({ status: 'working', message: '폴더 비교 중' })

    try {
      const result = await window.oneDriveManager.compareDriveFolders({
        source: createDriveFolderCompareEndpoint(source),
        target: createDriveFolderCompareEndpoint(target)
      })

      setFolderCompareState({ status: 'ready', source, target, result })
      setFileOperationState({
        status: 'success',
        message:
          result.differenceCount === 0
            ? '두 폴더의 차이를 찾지 못했습니다.'
            : `${result.differenceCount.toLocaleString('ko-KR')}개 차이를 찾았습니다.`
      })
    } catch (error) {
      setFolderCompareState({
        status: 'error',
        source,
        target,
        result: null,
        message: error instanceof Error ? error.message : '폴더를 비교하지 못했습니다.'
      })
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '폴더를 비교하지 못했습니다.'
      })
    }
  }

  async function reconcileComparedFolders(priority: DriveFolderReconcilePriority): Promise<void> {
    if (folderCompareState.status !== 'ready') {
      setFileOperationState({ status: 'error', message: '먼저 폴더 비교를 완료하세요.' })
      return
    }

    if (typeof window.oneDriveManager.reconcileDriveFolders !== 'function') {
      setFileOperationState({ status: 'error', message: '앱을 다시 실행한 뒤 폴더 맞추기를 사용할 수 있습니다.' })
      return
    }

    setFileOperationState({ status: 'working', message: '폴더 맞추기 전송 큐 등록 중' })

    try {
      const result = await window.oneDriveManager.reconcileDriveFolders({
        compare: folderCompareState.result,
        priority
      })
      const skippedMessage = result.skippedCount > 0 ? `, ${result.skippedCount.toLocaleString('ko-KR')}개는 종류 충돌 또는 미지원 항목으로 건너뜀` : ''
      const resultMessage =
        result.queuedCount > 0
          ? `${result.queuedCount.toLocaleString('ko-KR')}개 파일을 맞추기 전송 큐에 등록했습니다${skippedMessage}.`
          : result.createdFolderCount > 0
            ? `${result.createdFolderCount.toLocaleString('ko-KR')}개 폴더를 만들었습니다${skippedMessage}.`
            : `전송 큐에 등록할 항목이 없습니다${skippedMessage}.`

      setTransferTasks(await window.oneDriveManager.listTransfers())
      void refreshTransferPreview()
      folderCacheRef.current.clear()
      setFileOperationState({
        status: 'success',
        message: resultMessage
      })
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '폴더 맞추기를 등록하지 못했습니다.'
      })
    }
  }

  function clearFolderCompareSource(): void {
    setFolderCompareSource(null)
    setFileOperationState({ status: 'success', message: '폴더 비교 기준을 지웠습니다.' })
  }

  function closeFolderCompareDialog(): void {
    setFolderCompareState((currentState) => ({
      status: 'closed',
      source: currentState.source,
      target: currentState.target,
      result: currentState.result
    }))
  }

  function getFolderCompareTargetEndpoint(): FolderCompareEndpointView | null {
    if (!activeTab?.accountId) {
      return null
    }

    const targetFolder = contextMenuTargetItem?.type === 'folder' ? contextMenuTargetItem : null

    if (targetFolder) {
      return createFolderCompareEndpointFromItem(targetFolder, activeTab)
    }

    return createFolderCompareEndpointFromPath(folderPath, activeTab)
  }

  function createFolderCompareEndpointFromItem(item: CloudDriveItem, tab: DriveTab | null | undefined): FolderCompareEndpointView | null {
    if (!tab?.accountId) {
      return null
    }

    return createFolderCompareEndpointFromPath([...tab.folderPath, { id: item.id, name: item.name }], tab)
  }

  function createFolderCompareEndpointFromPath(path: DriveFolderRef[], tab: DriveTab | null | undefined): FolderCompareEndpointView | null {
    if (!tab?.accountId) {
      return null
    }

    const folder = path.at(-1) ?? rootFolder
    const account = getAccountForTab(session, tab.accountId)

    return {
      accountId: tab.accountId,
      accountLabel: account ? getAccountOptionLabel(account) : tab.accountId,
      folderId: folder.id,
      folderName: folder.name,
      pathLabel: getFolderPathLabel(path)
    }
  }

  async function downloadSelectedItems(items = selectedItems): Promise<void> {
    if (items.length === 0 || fileOperationState.status === 'working') {
      return
    }

    setContextMenu(null)
    setFileOperationState({ status: 'working', message: items.length === 1 ? '다운로드 등록 중' : '여러 항목 다운로드 등록 중' })

    try {
      if (items.length === 1 && items[0]?.type === 'file') {
        const [item] = items

        if (!item) {
          setFileOperationState({ status: 'idle' })
          return
        }

        const result = await window.oneDriveManager.downloadDriveItem({
          itemId: item.id,
          name: item.name,
          type: item.type,
          size: item.size
        })

        if (result.cancelled) {
          setFileOperationState({ status: 'idle' })
          return
        }

        setFileOperationState({
          status: 'success',
          message: `다운로드 작업을 큐에 등록했습니다: ${result.localPath ?? item.name}`
        })
        return
      }

      const result = await window.oneDriveManager.downloadDriveItems({
        items: items.map((item) => ({
          itemId: item.id,
          name: item.name,
          type: item.type,
          size: item.size
        }))
      })

      if (result.cancelled) {
        setFileOperationState({ status: 'idle' })
        return
      }

      const queuedCount = result.queuedCount ?? items.filter((item) => item.type === 'file').length
      const createdFolderCount = result.createdFolderCount ?? 0
      const skippedCount = result.skippedCount ?? 0
      const skippedMessage = skippedCount > 0 ? `, ${skippedCount.toLocaleString('ko-KR')}개 건너뜀` : ''

      setFileOperationState({
        status: 'success',
        message:
          queuedCount > 0
            ? `${queuedCount.toLocaleString('ko-KR')}개 파일 다운로드 작업을 큐에 등록했습니다. 폴더 ${createdFolderCount.toLocaleString('ko-KR')}개 생성${skippedMessage}.`
            : `폴더 ${createdFolderCount.toLocaleString('ko-KR')}개를 만들었습니다${skippedMessage}.`
      })
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '항목을 다운로드하지 못했습니다.'
      })
    }
  }

  async function renameSelectedItem(item = singleSelectedItem): Promise<void> {
    if (!item || fileOperationState.status === 'working') {
      return
    }

    setContextMenu(null)
    const nextName = await requestDriveName({
      title: '이름 변경',
      label: '새 이름',
      initialValue: item.name,
      confirmLabel: '변경'
    })

    if (!nextName || nextName === item.name) {
      return
    }

    const validationError = validateDriveItemName(nextName)

    if (validationError) {
      setFileOperationState({ status: 'error', message: validationError })
      return
    }

    setFileOperationState({ status: 'working', message: '이름 변경 중' })

    try {
      await window.oneDriveManager.renameDriveItem({
        itemId: item.id,
        name: nextName
      })
      folderCacheRef.current.clear()
      await loadDriveFolder(folderPath, { forceRefresh: true })
      void warmNavigationIndex(true)
      setFileOperationState({ status: 'success', message: '이름 변경 완료' })
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '이름을 변경하지 못했습니다.'
      })
    }
  }

  async function deleteSelectedItems(items = selectedItems): Promise<void> {
    if (items.length === 0 || fileOperationState.status === 'working') {
      return
    }

    const confirmMessage =
      items.length === 1
        ? `'${items[0]?.name ?? '선택한'}' 항목을 OneDrive 휴지통으로 이동할까요?`
        : `${items.length.toLocaleString('ko-KR')}개 항목을 OneDrive 휴지통으로 이동할까요?`

    if (!window.confirm(confirmMessage)) {
      return
    }

    setContextMenu(null)
    setFileOperationState({ status: 'working', message: '삭제 중' })

    try {
      for (const item of items) {
        await window.oneDriveManager.deleteDriveItem({ itemId: item.id })
      }

      if (driveClipboard?.items.some((clipboardItem) => items.some((item) => item.id === clipboardItem.id))) {
        setDriveClipboard(null)
      }

      folderCacheRef.current.clear()
      await loadDriveFolder(folderPath, { forceRefresh: true })
      void warmNavigationIndex(true)
      setFileOperationState({
        status: 'success',
        message: items.length === 1 ? '삭제 완료' : `${items.length.toLocaleString('ko-KR')}개 항목 삭제 완료`
      })
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '항목을 삭제하지 못했습니다.'
      })
    }
  }

  async function refreshCurrentFolder(): Promise<void> {
    setContextMenu(null)
    folderCacheRef.current.clear()
    void warmNavigationIndex(true)
    await loadDriveFolder(folderPath, { forceRefresh: true })
  }

  async function runManualIndexRefresh(): Promise<void> {
    if (!canUseDrive || isFileOperationBusy) {
      return
    }

    setContextMenu(null)
    setFileOperationState({ status: 'working', message: '탐색 인덱스 수동 갱신 중' })

    const didStart = await warmNavigationIndex(true, activeTabId, true)

    setFileOperationState({
      status: didStart ? 'success' : 'error',
      message: didStart ? '탐색 인덱스 갱신을 시작했습니다.' : '탐색 인덱스 갱신을 시작하지 못했습니다.'
    })
  }

  async function toggleDriveIndexMode(): Promise<void> {
    if (typeof window.oneDriveManager.updateDriveSettings !== 'function') {
      return
    }

    const nextIndexMode = driveSettings.indexMode === 'automatic' ? 'manual' : 'automatic'

    try {
      const nextSettings = await window.oneDriveManager.updateDriveSettings({ indexMode: nextIndexMode })

      setDriveSettings(nextSettings)
      setFileOperationState({
        status: 'success',
        message: nextSettings.indexMode === 'automatic' ? '탐색 인덱스를 자동으로 갱신합니다.' : '탐색 인덱스를 수동으로만 갱신합니다.'
      })

      if (nextSettings.indexMode === 'automatic') {
        void warmNavigationIndex(false, activeTabId, true)
      }
    } catch (error) {
      setFileOperationState({
        status: 'error',
        message: error instanceof Error ? error.message : '탐색 인덱스 설정을 저장하지 못했습니다.'
      })
    }
  }

  async function saveAuthSettings(): Promise<void> {
    setIsSavingSettings(true)
    setSettingsError(null)
    setSettingsMessage(null)
    setConnectionError(null)

    try {
      const nextSettings = {
        clientId: settingsForm.clientId.trim(),
        tenantId: settingsForm.tenantId.trim() || DEFAULT_TENANT_ID
      }
      const validationError = validateAuthSettings(nextSettings.clientId)

      if (validationError) {
        setSettingsError(validationError)
        return
      }

      const settings = await window.oneDriveManager.updateMicrosoftAuthSettings(nextSettings)
      folderCacheRef.current.clear()
      setSettingsForm({
        clientId: settings.clientId,
        tenantId: settings.tenantId
      })
      setSettingsMessage('로그인 설정을 저장했습니다.')
      await loadSession()
    } catch (error) {
      setSettingsError(error instanceof Error ? error.message : '로그인 설정을 저장하지 못했습니다.')
    } finally {
      setIsSavingSettings(false)
    }
  }

  async function resetAllSettings(): Promise<void> {
    if (!window.confirm('저장된 로그인 정보, 토큰 캐시, 앱 설정을 모두 초기화할까요?')) {
      return
    }

    setIsResettingSettings(true)
    setConnectionError(null)
    setSettingsError(null)
    setSettingsMessage(null)

    try {
      const session = await window.oneDriveManager.resetAllSettings()
      const nextTab = createDriveTab(null)

      setSessionState({ status: 'ready', session })
      folderCacheRef.current.clear()
      setTabs([nextTab])
      setActiveTabId(nextTab.id)
      setTransferTasks([])
      setSettingsForm({
        clientId: session.settings.clientId,
        tenantId: session.settings.tenantId
      })
      clearDriveSelection()
      setDriveClipboard(null)
      setFolderCompareSource(null)
      closeFolderCompareDialog()
      setSettingsMessage('저장된 설정과 캐시를 초기화했습니다.')
    } catch (error) {
      setConnectionError(error instanceof Error ? error.message : '설정과 캐시를 초기화하지 못했습니다.')
    } finally {
      setIsResettingSettings(false)
    }
  }

  function clearDriveSelection(): void {
    setSelectedItemIds([])
    setSelectionAnchorItemId(null)
    setContextMenu(null)
  }

  function selectDriveItem(item: CloudDriveItem, mode: DriveSelectionMode): void {
    selectDriveItemInTab(activeTabId, item, mode)
  }

  function selectDriveItemInTab(tabId: string, item: CloudDriveItem, mode: DriveSelectionMode): void {
    setContextMenu(null)
    setActiveTabId(tabId)
    const tab = tabs.find((tab) => tab.id === tabId)
    const tabItems = tab ? getSortedDriveTabItems(tab) : driveState.items

    if (mode === 'context' && tabId === activeTabId && selectedItemIds.includes(item.id)) {
      return
    }

    if (mode === 'toggle') {
      setSelectionAnchorItemId(item.id)
      setSelectedItemIds((currentItemIds) =>
        currentItemIds.includes(item.id) ? currentItemIds.filter((itemId) => itemId !== item.id) : [...currentItemIds, item.id]
      )
      return
    }

    if (mode === 'range') {
      const anchorId = selectionAnchorItemId ?? selectedItemIds.at(-1) ?? item.id
      const anchorIndex = tabItems.findIndex((candidate) => candidate.id === anchorId)
      const itemIndex = tabItems.findIndex((candidate) => candidate.id === item.id)

      if (anchorIndex >= 0 && itemIndex >= 0) {
        const startIndex = Math.min(anchorIndex, itemIndex)
        const endIndex = Math.max(anchorIndex, itemIndex)

        setSelectedItemIds(tabItems.slice(startIndex, endIndex + 1).map((candidate) => candidate.id))
        return
      }
    }

    setSelectionAnchorItemId(item.id)
    setSelectedItemIds([item.id])
  }

  function openContextMenuInTab(tabId: string, point: ContextMenuState): void {
    setActiveTabId(tabId)
    openContextMenu(point)
  }

  function openContextMenu(point: ContextMenuState): void {
    setContextMenu({
      x: Math.max(4, Math.min(point.x, window.innerWidth - 220)),
      y: Math.max(4, Math.min(point.y, window.innerHeight - 320)),
      targetItemId: point.targetItemId ?? null
    })
  }

  function getSelectedItems(): CloudDriveItem[] {
    const selectedIds = new Set(selectedItemIds)

    return activeTab ? getSortedDriveTabItems(activeTab).filter((item) => selectedIds.has(item.id)) : driveState.items.filter((item) => selectedIds.has(item.id))
  }

  const environment = environmentState.status === 'ready' ? environmentState.environment : null
  const selectedItems = getSelectedItems()
  const singleSelectedItem = selectedItems.length === 1 ? (selectedItems[0] ?? null) : null
  const currentFolder = folderPath.at(-1) ?? rootFolder
  const isFileOperationBusy = fileOperationState.status === 'working'
  const resumableTransferCount = transferTasks.filter(
    (task) => task.status === 'failed' || task.status === 'queued' || task.status === 'paused' || task.status === 'retrying'
  ).length
  const canStartSignIn = sessionState.status !== 'loading' && !isConnecting && !isResettingSettings
  const canUseDrive = Boolean(session?.isAuthenticated)
  const canUseActiveTabDrive = Boolean(session?.isAuthenticated && activeTab?.accountId)
  const shouldShowAuthSettings = !session?.isAuthenticated && !session?.isConfigured
  const statusMessage = connectionError ?? (sessionState.status === 'error' ? sessionState.message : null)
  const contextMenuTargetItem = contextMenu?.targetItemId ? driveState.items.find((item) => item.id === contextMenu.targetItemId) ?? null : null
  const pasteTargetFolderId =
    contextMenuTargetItem?.type === 'folder' && selectedItems.length === 1 && selectedItems[0]?.id === contextMenuTargetItem.id
      ? contextMenuTargetItem.id
      : currentFolder.id
  const canSetFolderCompareSource = Boolean(canUseActiveTabDrive && singleSelectedItem?.type === 'folder')
  const canCompareFolder = Boolean(
    canUseActiveTabDrive && folderCompareSource && (!contextMenuTargetItem || contextMenuTargetItem.type === 'folder')
  )
  const folderCompareTargetLabel = contextMenuTargetItem?.type === 'folder' ? '선택 폴더와 비교' : '현재 폴더와 비교'
  const previewItemCount = drivePreview ? getDrivePreviewItems(tabs.find((tab) => tab.id === drivePreview.tabId)).length : 0
  const cutItemIds = driveClipboard?.mode === 'cut' ? driveClipboard.items.map((item) => item.id) : []
  const latestGraphActivityEvent = graphActivityEvents[0] ?? null

  return (
    <main className="explorer-shell">
      <header className="title-bar">
        <div className="app-title">
          <span className="app-icon" aria-hidden="true" />
          <span>OneDrive 관리자</span>
        </div>
        <div className="title-account">{getAccountLine(session, isConnecting)}</div>
      </header>

      <section className="command-bar" aria-label="파일 관리자 명령">
        <button className="command-button" type="button" disabled={!canUseDrive || isFileOperationBusy} onClick={() => void uploadFiles()}>
          업로드
        </button>
        <button
          className="command-button"
          type="button"
          disabled={!canUseDrive || selectedItems.length === 0 || isFileOperationBusy}
          onClick={() => void downloadSelectedItems()}
        >
          다운로드
        </button>
        <button
          className="command-button"
          type="button"
          disabled={!canUseDrive || !singleSelectedItem || isFileOperationBusy}
          onClick={() => void renameSelectedItem(singleSelectedItem)}
        >
          이름 변경
        </button>
        <button
          className="command-button danger"
          type="button"
          disabled={!canUseDrive || selectedItems.length === 0 || isFileOperationBusy}
          onClick={() => void deleteSelectedItems()}
        >
          삭제
        </button>
        <button
          className="command-button"
          type="button"
          disabled={!canUseDrive || driveState.status === 'loading' || isFileOperationBusy}
          onClick={() => void refreshCurrentFolder()}
        >
          새로고침
        </button>
        <button
          className="command-button"
          type="button"
          disabled={!canUseDrive || isFileOperationBusy || indexState.status === 'syncing'}
          onClick={() => void runManualIndexRefresh()}
        >
          인덱스 갱신
        </button>
        <button className="command-button" type="button" disabled={isFileOperationBusy} onClick={() => void toggleDriveIndexMode()}>
          {driveSettings.indexMode === 'automatic' ? '인덱스 자동' : '인덱스 수동'}
        </button>
        <button className="command-button" type="button" disabled={!canUseDrive || isFileOperationBusy} onClick={() => void createNewDriveTab()}>
          새 탭
        </button>
        <button
          className="command-button"
          type="button"
          disabled={tabs.length <= 1 || isFileOperationBusy}
          onClick={() => void closeDriveTab(activeTabId)}
        >
          탭 닫기
        </button>
        <button
          className="command-button"
          type="button"
          disabled={!canUseDrive || resumableTransferCount === 0 || isFileOperationBusy}
          onClick={() => void resumeTransfers()}
        >
          전송 재개
        </button>
        {fileOperationState.status !== 'idle' ? (
          <span className="operation-status" data-state={fileOperationState.status}>
            {fileOperationState.message}
          </span>
        ) : null}
        {session?.isAuthenticated ? (
          <>
            <button className="command-button primary" type="button" disabled={!canStartSignIn} onClick={() => void connectAccount()}>
              {isConnecting ? '로그인 중' : '계정 추가'}
            </button>
            <button className="command-button" type="button" disabled={isResettingSettings || isConnecting} onClick={() => void disconnectAccount()}>
              현재 계정 로그아웃
            </button>
          </>
        ) : (
          <button className="command-button primary" type="button" disabled={!canStartSignIn} onClick={() => void connectAccount()}>
            {isConnecting ? '로그인 중' : '로그인'}
          </button>
        )}
      </section>

      <section className="manager-layout" aria-label="OneDrive 파일 관리자">
        <aside className="navigation-pane" aria-label="탐색 창">
          <button
            className="tree-item selected"
            type="button"
            disabled={!canUseActiveTabDrive || driveState.status === 'loading'}
            onClick={() => void openBreadcrumb(0)}
          >
            <span className="tree-icon drive" aria-hidden="true" />
            <span>내 OneDrive</span>
          </button>

          <div className="sidebar-section">
            <div className="sidebar-heading">계정</div>
            <AccountSwitcher
              session={session}
              usageState={accountUsageState}
              isBusy={isFileOperationBusy || isConnecting || driveState.status === 'loading'}
              onSwitch={(accountId) => void switchAccount(accountId)}
              onAdd={() => void connectAccount()}
              onDisconnect={() => void disconnectAccount()}
            />
            <button
              className="command-button danger full reset-button"
              type="button"
              disabled={isConnecting || isSavingSettings || isResettingSettings}
              onClick={() => void resetAllSettings()}
            >
              {isResettingSettings ? '초기화 중' : '설정 초기화'}
            </button>
          </div>

          {folderCompareSource ? (
            <div className="sidebar-section compare-source-section">
              <div className="sidebar-heading">폴더 비교 기준</div>
              <div className="compare-source-card">
                <strong title={folderCompareSource.pathLabel}>{folderCompareSource.folderName}</strong>
                <span title={folderCompareSource.pathLabel}>{folderCompareSource.pathLabel}</span>
                <small>{folderCompareSource.accountLabel}</small>
                <button className="transfer-clear-button" type="button" onClick={clearFolderCompareSource}>
                  기준 지우기
                </button>
              </div>
            </div>
          ) : null}

          <TransferPanel
            tasks={transferTasks}
            summary={transferSummary}
            settings={transferSettings}
            isBusy={isFileOperationBusy}
            onOpenViewer={() => void openTransferViewer()}
            onConcurrencyChange={(value) => void updateTransferConcurrency(value)}
            onResume={(taskId) => void resumeTransfers(taskId)}
            onStop={(taskId) => void stopTransfer(taskId)}
            onDelete={(taskId) => void deleteTransfer(taskId)}
          />

          {shouldShowAuthSettings ? (
            <AuthSettingsForm
              clientId={settingsForm.clientId}
              tenantId={settingsForm.tenantId}
              source={session?.settings.source ?? 'missing'}
              isSaving={isSavingSettings}
              isDisabled={isConnecting || sessionState.status === 'loading'}
              error={settingsError}
              message={settingsMessage}
              onChange={(nextForm) => {
                setSettingsForm(nextForm)
                setSettingsMessage(null)
                setSettingsError(null)
              }}
              onSave={() => void saveAuthSettings()}
            />
          ) : null}

          {!shouldShowAuthSettings && settingsMessage ? <p className="inline-message sidebar-message">{settingsMessage}</p> : null}

          {isConnecting ? (
            <div className="web-login-box" aria-live="polite">
              <strong>웹 로그인 진행 중</strong>
              <span>브라우저에서 5분 안에 완료하면 목록을 불러옵니다.</span>
            </div>
          ) : null}

          {statusMessage ? <p className="inline-error">{statusMessage}</p> : null}
        </aside>

        <section className="content-pane" aria-label="파일 목록">
          <div className="drive-workspace" data-pane-count={tabs.length} style={{ gridTemplateColumns: getDriveWorkspaceGridTemplateColumns(tabs) }}>
            {tabs.flatMap((tab, index) => {
              const isActive = tab.id === activeTabId
              const tabCurrentFolder = tab.folderPath.at(-1) ?? rootFolder
              const tabCanUseDrive = Boolean(session?.isAuthenticated && tab.accountId)
              const pane = (
                <DriveTabPane
                  key={tab.id}
                  tab={tab}
                  account={getAccountForTab(session, tab.accountId)}
                  isActive={isActive}
                  canUseDrive={tabCanUseDrive}
                  canSignIn={canStartSignIn}
                  canClose={tabs.length > 1}
                  isConnecting={isConnecting}
                  isBusy={isFileOperationBusy}
                  selectedItemIds={isActive ? selectedItemIds : []}
                  cutItemIds={cutItemIds}
                  dragOverFolderId={isActive ? dragOverFolderId : null}
                  currentFolderId={tabCurrentFolder.id}
                  onActivate={() => void activateDriveTab(tab.id)}
                  onClose={() => void closeDriveTab(tab.id)}
                  onNewTab={() => void createNewDriveTab()}
                  onSignIn={() => void connectAccount()}
                  onOpenParent={() => void openParentFolderInTab(tab.id)}
                  onOpenBreadcrumb={(index) => void openBreadcrumbInTab(tab.id, index)}
                  onClearSelection={() => {
                    setActiveTabId(tab.id)
                    clearDriveSelection()
                  }}
                  onBoxSelect={(itemIds) => {
                    setActiveTabId(tab.id)
                    setContextMenu(null)
                    setSelectedItemIds(itemIds)
                    setSelectionAnchorItemId(itemIds.at(-1) ?? null)
                  }}
                  onSortOptionsChange={(nextSortOptions) => {
                    setSortOptionsForTab(setTabs, tab.id, nextSortOptions)
                  }}
                  onColumnWidthsChange={(nextColumnWidths) => {
                    setColumnWidthsForTab(setTabs, tab.id, nextColumnWidths)
                  }}
                  onViewModeChange={(nextViewMode) => {
                    setViewModeForTab(setTabs, tab.id, nextViewMode)
                  }}
                  onSelectItem={(item, mode) => selectDriveItemInTab(tab.id, item, mode)}
                  onOpenFolder={(item) => void openFolderInTab(tab.id, item)}
                  onPreviewItem={(item) => previewItemInTab(tab.id, item)}
                  onRenameItem={(item) => void renameItemInTab(tab.id, item)}
                  onDeleteItem={(item) => void deleteItemInTab(tab.id, item)}
                  onDragOverFolderChange={setDragOverFolderId}
                  onMoveItemsToFolder={(item, payload) => void moveItemsToFolderInTab(tab.id, item, payload)}
                  onUploadDroppedFiles={(parentId, files) => void uploadDroppedFilesInTab(tab.id, parentId, files)}
                  onContextMenuItem={(item, point) => {
                    selectDriveItemInTab(tab.id, item, 'context')
                    openContextMenuInTab(tab.id, { ...point, targetItemId: item.id })
                  }}
                  onContextMenuBackground={(point) => openContextMenuInTab(tab.id, point)}
                  onLoadMore={() => void loadMoreInTab(tab.id)}
                />
              )

              if (index === tabs.length - 1) {
                return [pane]
              }

              const rightTab = tabs[index + 1]

              return [
                pane,
                <div
                  className="drive-pane-resize-handle"
                  key={`${tab.id}-resize-${rightTab?.id ?? 'end'}`}
                  role="separator"
                  aria-orientation="vertical"
                  aria-label="탭 크기 조절"
                  onMouseDown={(event) => {
                    if (rightTab) {
                      startDrivePaneResize(event, tab.id, rightTab.id)
                    }
                  }}
                />
              ]
            })}
          </div>
        </section>
      </section>

      {contextMenu ? (
        <ExplorerContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          canUseDrive={canUseDrive}
          isBusy={isFileOperationBusy}
          selectedItems={selectedItems}
          singleSelectedItem={singleSelectedItem}
          clipboard={driveClipboard}
          compareSource={folderCompareSource}
          canSetCompareSource={canSetFolderCompareSource}
          canCompareFolder={canCompareFolder}
          compareTargetLabel={folderCompareTargetLabel}
          onOpenFolder={(item) => void openFolder(item)}
          onDownload={() => void downloadSelectedItems()}
          onCopy={() => copySelectedItemsToClipboard('copy')}
          onCut={() => copySelectedItemsToClipboard('cut')}
          onPaste={() => void pasteClipboardItems(pasteTargetFolderId)}
          onSetCompareSource={setSelectedFolderAsCompareSource}
          onCompareFolder={() => void compareFolderWithSource()}
          onRename={() => void renameSelectedItem()}
          onDelete={() => void deleteSelectedItems()}
          onCreateFolder={() => void createFolder()}
          onUpload={() => void uploadFiles()}
          onRefresh={() => void refreshCurrentFolder()}
        />
      ) : null}

      {driveNameDialog ? (
        <DriveNameDialog
          dialog={driveNameDialog}
          onCancel={() => resolveDriveNameDialog(null)}
          onSubmit={(value) => resolveDriveNameDialog(value)}
        />
      ) : null}

      {drivePreview ? (
        <DrivePreviewDialog
          preview={drivePreview}
          hasMultipleItems={previewItemCount > 1}
          onClose={() => setDrivePreview(null)}
          onPrevious={() => moveDrivePreview(-1)}
          onNext={() => moveDrivePreview(1)}
        />
      ) : null}

      {transferViewerState.status !== 'closed' ? (
        <TransferViewerDialog
          state={transferViewerState}
          isBusy={isFileOperationBusy}
          onClose={closeTransferViewer}
          onLoadPage={(offset) => void loadTransferViewerPage(offset)}
          onResume={(taskId) => void resumeTransfers(taskId)}
          onStop={(taskId) => void stopTransfer(taskId)}
          onDelete={(taskId) => void deleteTransfer(taskId)}
        />
      ) : null}

      {folderCompareState.status !== 'closed' ? (
        <FolderCompareDialog
          state={folderCompareState}
          isBusy={fileOperationState.status === 'working'}
          onClose={closeFolderCompareDialog}
          onReconcile={(priority) => void reconcileComparedFolders(priority)}
        />
      ) : null}

      {isGraphActivityLogOpen ? (
        <GraphActivityDialog events={graphActivityEvents} onClose={() => setIsGraphActivityLogOpen(false)} />
      ) : null}

      <footer className="status-bar">
        <button className="status-bar-main" type="button" onClick={() => setIsGraphActivityLogOpen(true)}>
          <span>{latestGraphActivityEvent ? formatGraphActivityStatus(latestGraphActivityEvent) : getDriveStatusText(driveState, indexState)}</span>
          {latestGraphActivityEvent?.progress ? <GraphActivityProgressBar event={latestGraphActivityEvent} /> : null}
        </button>
        <span>{getDriveStatusText(driveState, indexState)}</span>
        <span>{environment ? platformLabels[environment.platform.name] ?? environment.platform.name : getEnvironmentStatus(environmentState)}</span>
      </footer>
    </main>
  )
}

function DriveNameDialog({
  dialog,
  onCancel,
  onSubmit
}: {
  dialog: DriveNameDialogState
  onCancel: () => void
  onSubmit: (value: string) => void
}): ReactElement {
  const [value, setValue] = useState(dialog.initialValue)
  const [error, setError] = useState<string | null>(null)
  const inputRef = useRef<HTMLInputElement | null>(null)

  useEffect(() => {
    setValue(dialog.initialValue)
    setError(null)
  }, [dialog])

  useEffect(() => {
    inputRef.current?.focus()
    inputRef.current?.select()
  }, [dialog])

  useEffect(() => {
    function handleKeyDown(event: KeyboardEvent): void {
      if (event.key === 'Escape') {
        onCancel()
      }
    }

    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [onCancel])

  function handleSubmit(event: FormEvent<HTMLFormElement>): void {
    event.preventDefault()
    const normalizedValue = value.trim()
    const validationError = validateDriveItemName(normalizedValue)

    if (validationError) {
      setError(validationError)
      return
    }

    onSubmit(normalizedValue)
  }

  return (
    <div className="name-dialog-backdrop" role="presentation" onMouseDown={onCancel}>
      <form
        className="name-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="drive-name-dialog-title"
        onMouseDown={(event) => event.stopPropagation()}
        onSubmit={handleSubmit}
      >
        <header>
          <h2 id="drive-name-dialog-title">{dialog.title}</h2>
        </header>
        <label>
          <span>{dialog.label}</span>
          <input
            ref={inputRef}
            type="text"
            value={value}
            onChange={(event) => {
              setValue(event.target.value)
              if (error) {
                setError(null)
              }
            }}
          />
        </label>
        {error ? <p>{error}</p> : null}
        <footer>
          <button type="button" onClick={onCancel}>
            취소
          </button>
          <button type="submit">{dialog.confirmLabel}</button>
        </footer>
      </form>
    </div>
  )
}

function AccountSwitcher({
  session,
  usageState,
  isBusy,
  onSwitch,
  onAdd,
  onDisconnect
}: {
  session: AuthSession | null
  usageState: AccountUsageState
  isBusy: boolean
  onSwitch: (accountId: string) => void
  onAdd: () => void
  onDisconnect: () => void
}): ReactElement {
  const activeAccount = session?.account
  const accounts = session?.accounts ?? []

  return (
    <div className="account-card">
      {accounts.length > 0 ? (
        <div className="account-list" aria-label="Microsoft 계정 목록">
          {accounts.map((account) => {
            const usage = usageState.usages[account.homeAccountId]
            const isActive = account.homeAccountId === activeAccount?.homeAccountId

            return (
              <button
                className="account-option-card"
                type="button"
                key={account.homeAccountId}
                aria-pressed={isActive}
                disabled={isBusy || isActive}
                onClick={() => onSwitch(account.homeAccountId)}
              >
                <span className="account-card-main">
                  <strong>{account.name || account.username}</strong>
                  <span>{account.username}</span>
                </span>
                <AccountUsageMeter usage={usage} isLoading={usageState.status === 'loading'} />
              </button>
            )
          })}
        </div>
      ) : (
        <div className="account-empty-card">
          <strong>로그인 필요</strong>
          <span>Microsoft 계정을 연결하세요.</span>
        </div>
      )}
      {usageState.status === 'error' ? <p className="account-usage-error">{usageState.message}</p> : null}
      {session?.isConfigured ? (
        <div className="account-actions">
          <button type="button" disabled={isBusy} onClick={onAdd}>
            계정 추가
          </button>
          <button type="button" disabled={!activeAccount || isBusy} onClick={onDisconnect}>
            제거
          </button>
        </div>
      ) : null}
    </div>
  )
}

function AccountUsageMeter({ usage, isLoading }: { usage: DriveAccountUsage | undefined; isLoading: boolean }): ReactElement {
  const ratio = usage?.total && usage.total > 0 ? Math.min(100, Math.max(0, Math.round((usage.used / usage.total) * 100))) : 0
  const label = getAccountUsageLabel(usage, isLoading)

  return (
    <span className="account-usage">
      <span className="account-usage-row">
        <span>{label}</span>
        {usage?.state ? <span>{getQuotaStateLabel(usage.state)}</span> : null}
      </span>
      <span className="account-usage-track" aria-label={`OneDrive 사용량 ${label}`}>
        <span style={{ width: `${ratio}%` }} />
      </span>
    </span>
  )
}

function ExplorerContextMenu({
  x,
  y,
  canUseDrive,
  isBusy,
  selectedItems,
  singleSelectedItem,
  clipboard,
  compareSource,
  canSetCompareSource,
  canCompareFolder,
  compareTargetLabel,
  onOpenFolder,
  onDownload,
  onCopy,
  onCut,
  onPaste,
  onSetCompareSource,
  onCompareFolder,
  onRename,
  onDelete,
  onCreateFolder,
  onUpload,
  onRefresh
}: {
  x: number
  y: number
  canUseDrive: boolean
  isBusy: boolean
  selectedItems: CloudDriveItem[]
  singleSelectedItem: CloudDriveItem | null
  clipboard: DriveClipboard | null
  compareSource: FolderCompareEndpointView | null
  canSetCompareSource: boolean
  canCompareFolder: boolean
  compareTargetLabel: string
  onOpenFolder: (item: CloudDriveItem) => void
  onDownload: () => void
  onCopy: () => void
  onCut: () => void
  onPaste: () => void
  onSetCompareSource: () => void
  onCompareFolder: () => void
  onRename: () => void
  onDelete: () => void
  onCreateFolder: () => void
  onUpload: () => void
  onRefresh: () => void
}): ReactElement {
  const selectedFolder = singleSelectedItem?.type === 'folder' ? singleSelectedItem : null
  const pasteLabel = clipboard
    ? `${clipboard.mode === 'cut' ? '잘라낸' : '복사한'} 항목 붙여넣기${clipboard.items.length > 1 ? ` (${clipboard.items.length.toLocaleString('ko-KR')})` : ''}`
    : '붙여넣기'

  return (
    <div
      className="context-menu"
      role="menu"
      style={{ left: x, top: y }}
      onContextMenu={(event) => event.preventDefault()}
      onClick={(event) => event.stopPropagation()}
    >
      {selectedFolder ? (
        <button type="button" role="menuitem" disabled={!canUseDrive || isBusy} onClick={() => onOpenFolder(selectedFolder)}>
          열기
        </button>
      ) : null}
      <button type="button" role="menuitem" disabled={!canUseDrive || selectedItems.length === 0 || isBusy} onClick={onDownload}>
        {selectedItems.length > 1 ? `${selectedItems.length.toLocaleString('ko-KR')}개 다운로드` : '다운로드'}
      </button>
      <span className="context-menu-separator" role="separator" />
      <button type="button" role="menuitem" disabled={!canUseDrive || selectedItems.length === 0 || isBusy} onClick={onCopy}>
        복사
      </button>
      <button type="button" role="menuitem" disabled={!canUseDrive || selectedItems.length === 0 || isBusy} onClick={onCut}>
        잘라내기
      </button>
      <button type="button" role="menuitem" disabled={!canUseDrive || !clipboard || clipboard.items.length === 0 || isBusy} onClick={onPaste}>
        {pasteLabel}
      </button>
      <span className="context-menu-separator" role="separator" />
      <button type="button" role="menuitem" disabled={!canUseDrive || !canSetCompareSource || isBusy} onClick={onSetCompareSource}>
        비교 기준으로 선택
      </button>
      <button type="button" role="menuitem" disabled={!canUseDrive || !compareSource || !canCompareFolder || isBusy} onClick={onCompareFolder}>
        {compareTargetLabel}
      </button>
      <span className="context-menu-separator" role="separator" />
      <button type="button" role="menuitem" disabled={!canUseDrive || !singleSelectedItem || isBusy} onClick={onRename}>
        이름 변경
      </button>
      <button
        className="danger"
        type="button"
        role="menuitem"
        disabled={!canUseDrive || selectedItems.length === 0 || isBusy}
        onClick={onDelete}
      >
        삭제
      </button>
      <span className="context-menu-separator" role="separator" />
      <button type="button" role="menuitem" disabled={!canUseDrive || isBusy} onClick={onCreateFolder}>
        새 폴더
      </button>
      <button type="button" role="menuitem" disabled={!canUseDrive || isBusy} onClick={onUpload}>
        업로드
      </button>
      <button type="button" role="menuitem" disabled={!canUseDrive || isBusy} onClick={onRefresh}>
        새로고침
      </button>
    </div>
  )
}

function DrivePreviewDialog({
  preview,
  hasMultipleItems,
  onClose,
  onPrevious,
  onNext
}: {
  preview: DrivePreviewState
  hasMultipleItems: boolean
  onClose: () => void
  onPrevious: () => void
  onNext: () => void
}): ReactElement {
  const item = preview.item
  const previewKind = getDriveItemPreviewKind(item)
  const thumbnailCacheKey = `${getDriveThumbnailCacheKey(item)}:${getDrivePreviewThumbnailSize()}`
  const previewStageRef = useRef<HTMLDivElement | null>(null)
  const previewImageRef = useRef<HTMLImageElement | null>(null)
  const previewPanSessionRef = useRef<PreviewPanSession | null>(null)
  const [thumbnailState, setThumbnailState] = useState<ThumbnailPreviewState>({ status: 'loading' })
  const [previewTransform, setPreviewTransform] = useState<PreviewTransform>({ scale: 1, x: 0, y: 0 })
  const canTransformPreview = thumbnailState.status === 'ready'
  const canPanPreview = canTransformPreview && previewTransform.scale > DRIVE_PREVIEW_ZOOM_MIN

  useEffect(() => {
    let isCancelled = false
    let hasTimedOut = false
    const size = getDrivePreviewThumbnailSize()

    setThumbnailState({ status: 'loading' })

    const previewTimeout = window.setTimeout(() => {
      hasTimedOut = true

      if (!isCancelled) {
        setThumbnailState({ status: 'missing' })
      }
    }, DRIVE_PREVIEW_THUMBNAIL_WAIT_MS)

    void window.oneDriveManager
      .getDriveThumbnail({
        accountId: preview.accountId,
        itemId: item.id,
        cacheKey: `${getDriveThumbnailCacheKey(item)}:${size}`,
        priority: 'high',
        size
      })
      .then((thumbnail) => {
        if (isCancelled) {
          return
        }

        window.clearTimeout(previewTimeout)

        if (thumbnail.status === 'ready' && thumbnail.url) {
          setThumbnailState({
            status: 'ready',
            url: thumbnail.url,
            width: thumbnail.width,
            height: thumbnail.height
          })
          return
        }

        setThumbnailState({ status: 'missing' })
      })
      .catch(() => {
        window.clearTimeout(previewTimeout)

        if (!isCancelled && !hasTimedOut) {
          setThumbnailState({ status: 'error' })
        }
      })

    return () => {
      isCancelled = true
      window.clearTimeout(previewTimeout)
    }
  }, [item.id, preview.accountId, thumbnailCacheKey])

  useEffect(() => {
    previewPanSessionRef.current = null
    setPreviewTransform({ scale: 1, x: 0, y: 0 })
  }, [item.id, thumbnailCacheKey])

  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        event.preventDefault()
        onClose()
        return
      }

      if (!hasMultipleItems) {
        return
      }

      if (event.key === 'ArrowRight' || event.key === 'PageDown') {
        event.preventDefault()
        onNext()
      }

      if (event.key === 'ArrowLeft' || event.key === 'PageUp') {
        event.preventDefault()
        onPrevious()
      }
    }

    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [hasMultipleItems, onClose, onNext, onPrevious])

  function setPreviewZoom(nextScale: number): void {
    setPreviewTransform((currentTransform) => clampPreviewTransform({ ...currentTransform, scale: nextScale }))
  }

  function resetPreviewTransform(): void {
    previewPanSessionRef.current = null
    setPreviewTransform({ scale: 1, x: 0, y: 0 })
  }

  function handlePreviewWheel(event: ReactWheelEvent<HTMLDivElement>): void {
    if (!canTransformPreview) {
      return
    }

    event.preventDefault()
    const direction = event.deltaY > 0 ? -1 : 1
    setPreviewZoom(previewTransform.scale + direction * DRIVE_PREVIEW_ZOOM_STEP)
  }

  function handlePreviewPointerDown(event: ReactPointerEvent<HTMLDivElement>): void {
    if (!canPanPreview || event.button !== 0) {
      return
    }

    event.preventDefault()
    previewPanSessionRef.current = {
      pointerId: event.pointerId,
      originX: event.clientX,
      originY: event.clientY,
      startX: previewTransform.x,
      startY: previewTransform.y
    }
    event.currentTarget.setPointerCapture(event.pointerId)
  }

  function handlePreviewPointerMove(event: ReactPointerEvent<HTMLDivElement>): void {
    const panSession = previewPanSessionRef.current

    if (!panSession || panSession.pointerId !== event.pointerId) {
      return
    }

    setPreviewTransform((currentTransform) =>
      clampPreviewTransform({
        ...currentTransform,
        x: panSession.startX + event.clientX - panSession.originX,
        y: panSession.startY + event.clientY - panSession.originY
      })
    )
  }

  function handlePreviewPointerEnd(event: ReactPointerEvent<HTMLDivElement>): void {
    const panSession = previewPanSessionRef.current

    if (!panSession || panSession.pointerId !== event.pointerId) {
      return
    }

    previewPanSessionRef.current = null

    if (event.currentTarget.hasPointerCapture(event.pointerId)) {
      event.currentTarget.releasePointerCapture(event.pointerId)
    }
  }

  function clampPreviewTransform(transform: PreviewTransform): PreviewTransform {
    const scale = clampNumber(transform.scale, DRIVE_PREVIEW_ZOOM_MIN, DRIVE_PREVIEW_ZOOM_MAX)

    if (scale <= DRIVE_PREVIEW_ZOOM_MIN) {
      return { scale: DRIVE_PREVIEW_ZOOM_MIN, x: 0, y: 0 }
    }

    const imageWidth = previewImageRef.current?.clientWidth ?? previewStageRef.current?.clientWidth ?? 0
    const imageHeight = previewImageRef.current?.clientHeight ?? previewStageRef.current?.clientHeight ?? 0
    const maxX = Math.max(0, (imageWidth * (scale - 1)) / 2)
    const maxY = Math.max(0, (imageHeight * (scale - 1)) / 2)

    return {
      scale,
      x: clampNumber(transform.x, -maxX, maxX),
      y: clampNumber(transform.y, -maxY, maxY)
    }
  }

  return (
    <div className="preview-dialog-backdrop" role="presentation" onMouseDown={onClose}>
      <section
        className="preview-dialog"
        role="dialog"
        aria-modal="true"
        aria-label="파일 미리보기"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="preview-dialog-header">
          <div>
            <h2 title={item.name}>{item.name}</h2>
            <span>{`${getTypeLabel(item)} · ${item.type === 'file' ? formatBytes(item.size) : formatDate(item.lastModifiedDateTime)}`}</span>
          </div>
          <button className="preview-dialog-close" type="button" aria-label="닫기" onClick={onClose}>
            x
          </button>
        </header>

        <div
          ref={previewStageRef}
          className="preview-stage"
          data-can-pan={canPanPreview ? 'true' : 'false'}
          data-preview-kind={previewKind}
          data-thumbnail-state={thumbnailState.status}
          onPointerDown={handlePreviewPointerDown}
          onPointerMove={handlePreviewPointerMove}
          onPointerUp={handlePreviewPointerEnd}
          onPointerCancel={handlePreviewPointerEnd}
          onWheel={handlePreviewWheel}
        >
          {thumbnailState.status === 'ready' ? (
            <img
              ref={previewImageRef}
              className="preview-image"
              src={thumbnailState.url}
              width={thumbnailState.width}
              height={thumbnailState.height}
              style={{
                transform: `translate3d(${previewTransform.x}px, ${previewTransform.y}px, 0) scale(${previewTransform.scale})`
              }}
              alt=""
              aria-hidden="true"
              decoding="async"
              draggable={false}
              referrerPolicy="no-referrer"
              onError={() => setThumbnailState({ status: 'error' })}
            />
          ) : (
            <div className="preview-placeholder">
              <DriveItemIcon item={item} />
              <span>{thumbnailState.status === 'loading' ? '미리보기 로딩 중' : '미리보기 없음'}</span>
            </div>
          )}
        </div>

        <footer className="preview-dialog-footer">
          <button type="button" disabled={!hasMultipleItems} onClick={onPrevious}>
            이전
          </button>
          <div className="preview-zoom-controls" role="group" aria-label="미리보기 확대/축소">
            <button
              type="button"
              aria-label="축소"
              disabled={!canTransformPreview || previewTransform.scale <= DRIVE_PREVIEW_ZOOM_MIN}
              onClick={() => setPreviewZoom(previewTransform.scale - DRIVE_PREVIEW_ZOOM_STEP)}
            >
              -
            </button>
            <span>{`${Math.round(previewTransform.scale * 100)}%`}</span>
            <button
              type="button"
              aria-label="확대"
              disabled={!canTransformPreview || previewTransform.scale >= DRIVE_PREVIEW_ZOOM_MAX}
              onClick={() => setPreviewZoom(previewTransform.scale + DRIVE_PREVIEW_ZOOM_STEP)}
            >
              +
            </button>
            <button type="button" disabled={!canTransformPreview || previewTransform.scale === 1} onClick={resetPreviewTransform}>
              100%
            </button>
          </div>
          <button type="button" disabled={!hasMultipleItems} onClick={onNext}>
            다음
          </button>
        </footer>
      </section>
    </div>
  )
}

function TransferPanel({
  tasks,
  summary,
  settings,
  isBusy,
  onOpenViewer,
  onConcurrencyChange,
  onResume,
  onStop,
  onDelete
}: {
  tasks: DriveTransferTask[]
  summary: DriveTransferSummary | null
  settings: TransferSettings
  isBusy: boolean
  onOpenViewer: () => void
  onConcurrencyChange: (value: number) => void
  onResume: (taskId: string) => void
  onStop: (taskId: string) => void
  onDelete: (taskId?: string) => void
}): ReactElement {
  const visibleTasks = tasks.slice(0, 5)
  const progress = summary ? getTransferSummaryProgress(summary) : 0

  return (
    <div className="sidebar-section transfer-section">
      <div className="transfer-heading-row">
        <div className="sidebar-heading">전송</div>
        <div className="transfer-heading-actions">
          <button className="transfer-clear-button" type="button" onClick={onOpenViewer}>
            전체 보기
          </button>
          {summary && summary.totalCount > 0 ? (
            <button className="transfer-clear-button" type="button" onClick={() => onDelete()}>
              전체 삭제
            </button>
          ) : null}
        </div>
      </div>
      {summary && summary.totalCount > 0 ? (
        <div className="transfer-summary">
          <div className="transfer-summary-row">
            <strong>{`${summary.completedCount.toLocaleString('ko-KR')} / ${summary.totalCount.toLocaleString('ko-KR')}개`}</strong>
            <span>{`${progress}%`}</span>
          </div>
          <div className="transfer-progress" aria-label={`전체 전송 진행률 ${progress}%`}>
            <span style={{ width: `${progress}%` }} />
          </div>
          <div className="transfer-meta">
            <span>{`남은 항목 ${summary.activeCount.toLocaleString('ko-KR')}개`}</span>
            <span>{`${formatBytes(summary.bytesPerSecond)}/s`}</span>
            <span>{getTransferSummaryEtaLabel(summary)}</span>
          </div>
        </div>
      ) : null}
      <div className="transfer-setting">
        <div className="transfer-setting-row">
          <label htmlFor="transfer-concurrency">동시 전송 슬롯</label>
          <output htmlFor="transfer-concurrency">{`${settings.maxConcurrentTransfers}개`}</output>
        </div>
        <input
          id="transfer-concurrency"
          type="range"
          min={settings.minConcurrentTransfers}
          max={settings.maxAllowedConcurrentTransfers}
          step={1}
          value={settings.maxConcurrentTransfers}
          onChange={(event) => onConcurrencyChange(Number(event.target.value))}
        />
        <div className="transfer-setting-row subtle">
          <span>{`${settings.minConcurrentTransfers}개`}</span>
          <span>{`${settings.maxAllowedConcurrentTransfers}개`}</span>
        </div>
      </div>
      {visibleTasks.length > 0 ? (
        <div className="transfer-list">
          {visibleTasks.map((task) => {
            const progress = getTransferProgress(task)
            const canResume = isTransferTaskResumable(task)

            return (
              <article className="transfer-item" key={task.id} data-state={task.status}>
                <div className="transfer-title">
                  <strong title={task.name}>{task.name}</strong>
                  <span>{getTransferKindLabel(task.kind)}</span>
                </div>
                <div className="transfer-progress" aria-label={`${task.name} 전송률 ${progress}%`}>
                  <span style={{ width: `${progress}%` }} />
                </div>
                <div className="transfer-meta">
                  <span>{getTransferStatusLabel(task)}</span>
                  <span>{`${formatBytes(task.transferredBytes)} / ${task.totalBytes > 0 ? formatBytes(task.totalBytes) : '-'}`}</span>
                </div>
                <div className="transfer-meta">
                  <span>{task.status === 'running' ? `${formatBytes(task.bytesPerSecond ?? 0)}/s` : '속도 -'}</span>
                  <span>{formatDate(task.updatedAt)}</span>
                </div>
                <div className="transfer-meta">
                  <span>{getTransferElapsedLabel(task)}</span>
                  <span>{getTransferEtaLabel(task)}</span>
                </div>
                <div className="transfer-actions">
                  {canResume ? (
                    <button className="transfer-action-button" type="button" disabled={isBusy} onClick={() => onResume(task.id)}>
                      재개
                    </button>
                  ) : null}
                  {task.status === 'running' || task.status === 'queued' || task.status === 'retrying' ? (
                    <button className="transfer-action-button danger" type="button" onClick={() => onStop(task.id)}>
                      중지
                    </button>
                  ) : null}
                  <button className="transfer-action-button danger" type="button" onClick={() => onDelete(task.id)}>
                    기록 삭제
                  </button>
                </div>
              </article>
            )
          })}
        </div>
      ) : (
        <p className="transfer-empty">진행 중인 전송이 없습니다.</p>
      )}
    </div>
  )
}

function TransferViewerDialog({
  state,
  isBusy,
  onClose,
  onLoadPage,
  onResume,
  onStop,
  onDelete
}: {
  state: TransferViewerState
  isBusy: boolean
  onClose: () => void
  onLoadPage: (offset: number) => void
  onResume: (taskId: string) => void
  onStop: (taskId: string) => void
  onDelete: (taskId?: string) => void
}): ReactElement {
  const result = state.result
  const summary = result?.summary ?? null
  const totalTaskCount = result?.totalTaskCount ?? result?.summary.activeCount ?? 0
  const pageStart = result && result.tasks.length > 0 ? result.offset + 1 : 0
  const pageEnd = result ? result.offset + result.tasks.length : 0
  const canGoPrevious = state.offset > 0 && state.status !== 'loading'
  const canGoNext = Boolean(result && result.offset + result.limit < totalTaskCount && state.status !== 'loading')

  return (
    <div className="transfer-dialog-backdrop" role="presentation" onMouseDown={onClose}>
      <section
        className="transfer-dialog"
        role="dialog"
        aria-modal="true"
        aria-label="전체 전송 상황"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="transfer-dialog-header">
          <div>
            <h2>전체 전송 상황</h2>
            <span>{summary ? getTransferSummaryLabel(summary) : '전송 요약을 불러오는 중'}</span>
          </div>
          <button className="transfer-dialog-close" type="button" aria-label="닫기" onClick={onClose}>
            x
          </button>
        </header>

        {summary ? <TransferSummaryBlock summary={summary} /> : null}

        <div className="transfer-dialog-toolbar">
          <span>
            {result
              ? `${pageStart.toLocaleString('ko-KR')}-${pageEnd.toLocaleString('ko-KR')} / 목록 ${totalTaskCount.toLocaleString('ko-KR')}개`
              : '목록 준비 중'}
          </span>
          <div className="transfer-dialog-actions">
            <button type="button" disabled={!canGoPrevious} onClick={() => onLoadPage(Math.max(0, state.offset - state.limit))}>
              이전
            </button>
            <button type="button" disabled={!canGoNext} onClick={() => onLoadPage(state.offset + state.limit)}>
              다음
            </button>
            <button type="button" disabled={state.status === 'loading'} onClick={() => onLoadPage(state.offset)}>
              새로고침
            </button>
            <button className="danger" type="button" disabled={!summary || summary.totalCount === 0} onClick={() => onDelete()}>
              전체 삭제
            </button>
          </div>
        </div>

        {state.status === 'error' ? <p className="transfer-dialog-error">{state.message}</p> : null}

        <div className="transfer-table" role="table" aria-label="전송 작업 목록">
          <div className="transfer-table-header" role="row">
            <span role="columnheader">상태</span>
            <span role="columnheader">이름</span>
            <span role="columnheader">종류</span>
            <span role="columnheader">진행</span>
            <span role="columnheader">속도</span>
            <span role="columnheader">시간</span>
            <span role="columnheader">작업</span>
          </div>
          <div className="transfer-table-body">
            {result && result.tasks.length > 0 ? (
              result.tasks.map((task) => {
                const progress = getTransferProgress(task)
                const canResume = isTransferTaskResumable(task)

                return (
                  <div className="transfer-table-row" role="row" key={task.id} data-state={task.status}>
                    <span role="cell">{getTransferStatusLabel(task)}</span>
                    <strong role="cell" title={task.name}>
                      {task.name}
                    </strong>
                    <span role="cell">{getTransferKindLabel(task.kind)}</span>
                    <span role="cell">{`${progress}% · ${formatBytes(task.transferredBytes)} / ${task.totalBytes > 0 ? formatBytes(task.totalBytes) : '-'}`}</span>
                    <span role="cell">{task.status === 'running' ? `${formatBytes(task.bytesPerSecond ?? 0)}/s` : '-'}</span>
                    <span role="cell">{`${getTransferElapsedLabel(task)} · ${getTransferEtaLabel(task)}`}</span>
                    <span className="transfer-table-actions" role="cell">
                      {canResume ? (
                        <button type="button" disabled={isBusy} onClick={() => onResume(task.id)}>
                          재개
                        </button>
                      ) : null}
                      {task.status === 'running' || task.status === 'queued' || task.status === 'retrying' ? (
                        <button type="button" className="danger" onClick={() => onStop(task.id)}>
                          중지
                        </button>
                      ) : null}
                      <button type="button" className="danger" onClick={() => onDelete(task.id)}>
                        삭제
                      </button>
                    </span>
                  </div>
                )
              })
            ) : (
              <p className="transfer-table-empty">{state.status === 'loading' ? '전송 목록을 불러오는 중입니다.' : '표시할 전송 작업이 없습니다.'}</p>
            )}
          </div>
        </div>
      </section>
    </div>
  )
}

function TransferSummaryBlock({ summary }: { summary: DriveTransferSummary }): ReactElement {
  const progress = getTransferSummaryProgress(summary)

  return (
    <section className="transfer-dialog-summary" aria-label="전체 진행 요약">
      <div className="transfer-summary-row">
        <strong>{`${summary.completedCount.toLocaleString('ko-KR')} / ${summary.totalCount.toLocaleString('ko-KR')}개 완료`}</strong>
        <span>{`${progress}%`}</span>
      </div>
      <div className="transfer-progress" aria-label={`전체 전송 진행률 ${progress}%`}>
        <span style={{ width: `${progress}%` }} />
      </div>
      <div className="transfer-summary-grid">
        <span>{`진행 중 ${summary.runningCount.toLocaleString('ko-KR')}`}</span>
        <span>{`대기 ${summary.queuedCount.toLocaleString('ko-KR')}`}</span>
        <span>{`재시도 ${summary.retryingCount.toLocaleString('ko-KR')}`}</span>
        <span>{`중지 ${summary.pausedCount.toLocaleString('ko-KR')}`}</span>
        <span>{`실패 ${summary.failedCount.toLocaleString('ko-KR')}`}</span>
        <span>{`${formatBytes(summary.transferredBytes)} / ${summary.totalBytes > 0 ? formatBytes(summary.totalBytes) : '-'}`}</span>
        <span>{getTransferSummaryEtaLabel(summary)}</span>
      </div>
    </section>
  )
}

function FolderCompareDialog({
  state,
  isBusy,
  onClose,
  onReconcile
}: {
  state: FolderCompareState
  isBusy: boolean
  onClose: () => void
  onReconcile: (priority: DriveFolderReconcilePriority) => void
}): ReactElement {
  const [pageIndex, setPageIndex] = useState(0)
  const [reconcilePriority, setReconcilePriority] = useState<DriveFolderReconcilePriority>('source')
  const result = state.result
  const differences = result?.differences ?? []
  const totalPages = Math.max(1, Math.ceil(differences.length / FOLDER_COMPARE_PAGE_SIZE))
  const normalizedPageIndex = Math.min(pageIndex, totalPages - 1)
  const pageStartIndex = normalizedPageIndex * FOLDER_COMPARE_PAGE_SIZE
  const pageDifferences = differences.slice(pageStartIndex, pageStartIndex + FOLDER_COMPARE_PAGE_SIZE)
  const pageStart = differences.length > 0 ? pageStartIndex + 1 : 0
  const pageEnd = Math.min(pageStartIndex + pageDifferences.length, differences.length)

  useEffect(() => {
    setPageIndex(0)
  }, [result?.comparedAt, state.status])

  return (
    <div className="transfer-dialog-backdrop" role="presentation" onMouseDown={onClose}>
      <section
        className="folder-compare-dialog"
        role="dialog"
        aria-modal="true"
        aria-label="폴더 비교 결과"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="transfer-dialog-header">
          <div>
            <h2>폴더 비교</h2>
            <span>{`${state.source?.pathLabel ?? '-'} ↔ ${state.target?.pathLabel ?? '-'}`}</span>
          </div>
          <button className="transfer-dialog-close" type="button" aria-label="닫기" onClick={onClose}>
            x
          </button>
        </header>

        <div className="folder-compare-endpoints">
          <div>
            <span>기준</span>
            <strong title={state.source?.pathLabel}>{state.source?.pathLabel ?? '-'}</strong>
            <small>{state.source?.accountLabel ?? '-'}</small>
          </div>
          <div>
            <span>대상</span>
            <strong title={state.target?.pathLabel}>{state.target?.pathLabel ?? '-'}</strong>
            <small>{state.target?.accountLabel ?? '-'}</small>
          </div>
        </div>

        {result ? (
          <div className="folder-compare-summary">
            <span>{`기준 ${result.sourceItemCount.toLocaleString('ko-KR')}개`}</span>
            <span>{`대상 ${result.targetItemCount.toLocaleString('ko-KR')}개`}</span>
            <span>{`기준에만 ${result.onlyInSourceCount.toLocaleString('ko-KR')}개`}</span>
            <span>{`대상에만 ${result.onlyInTargetCount.toLocaleString('ko-KR')}개`}</span>
            <span>{`내용 다름 ${result.changedCount.toLocaleString('ko-KR')}개`}</span>
          </div>
        ) : null}

        {result && differences.length > 0 ? (
          <div className="folder-reconcile-controls">
            <div>
              <strong>맞추기</strong>
              <span>삭제 없이 없는 항목은 추가하고, 다른 파일은 선택한 기준으로 덮어쓰기</span>
            </div>
            <div className="folder-reconcile-priority" aria-label="다른 파일 덮어쓰기 기준">
              <button
                type="button"
                className={reconcilePriority === 'source' ? 'selected' : undefined}
                disabled={isBusy || state.status === 'loading'}
                onClick={() => setReconcilePriority('source')}
              >
                기준 우선
              </button>
              <button
                type="button"
                className={reconcilePriority === 'target' ? 'selected' : undefined}
                disabled={isBusy || state.status === 'loading'}
                onClick={() => setReconcilePriority('target')}
              >
                대상 우선
              </button>
            </div>
            <button
              className="folder-reconcile-run"
              type="button"
              disabled={isBusy || state.status !== 'ready'}
              onClick={() => onReconcile(reconcilePriority)}
            >
              맞추기 실행
            </button>
          </div>
        ) : null}

        <div className="transfer-dialog-toolbar">
          <span>
            {state.status === 'loading'
              ? '비교 중'
              : `${pageStart.toLocaleString('ko-KR')}-${pageEnd.toLocaleString('ko-KR')} / 차이 ${differences.length.toLocaleString('ko-KR')}개`}
          </span>
          <div className="transfer-dialog-actions">
            <button type="button" disabled={normalizedPageIndex <= 0 || state.status === 'loading'} onClick={() => setPageIndex((page) => Math.max(0, page - 1))}>
              이전
            </button>
            <button
              type="button"
              disabled={normalizedPageIndex >= totalPages - 1 || state.status === 'loading'}
              onClick={() => setPageIndex((page) => Math.min(totalPages - 1, page + 1))}
            >
              다음
            </button>
          </div>
        </div>

        {state.status === 'error' ? <p className="transfer-dialog-error">{state.message}</p> : null}

        <div className="folder-compare-table" role="table" aria-label="폴더 비교 차이 목록">
          <div className="folder-compare-table-header" role="row">
            <span role="columnheader">상태</span>
            <span role="columnheader">상대 경로</span>
            <span role="columnheader">기준</span>
            <span role="columnheader">대상</span>
            <span role="columnheader">판정</span>
          </div>
          <div className="folder-compare-table-body">
            {pageDifferences.length > 0 ? (
              pageDifferences.map((difference) => (
                <div className="folder-compare-table-row" role="row" key={`${difference.kind}-${difference.path}`}>
                  <span role="cell">{getFolderCompareKindLabel(difference.kind)}</span>
                  <strong role="cell" title={difference.path}>
                    {difference.path}
                  </strong>
                  <span role="cell">{formatFolderCompareItem(difference.source)}</span>
                  <span role="cell">{formatFolderCompareItem(difference.target)}</span>
                  <span role="cell">{formatFolderCompareReasons(difference)}</span>
                </div>
              ))
            ) : (
              <p className="transfer-table-empty">
                {state.status === 'loading' ? '두 폴더의 항목을 비교하는 중입니다.' : '두 폴더의 차이를 찾지 못했습니다.'}
              </p>
            )}
          </div>
        </div>
      </section>
    </div>
  )
}

function GraphActivityDialog({ events, onClose }: { events: GraphActivityEvent[]; onClose: () => void }): ReactElement {
  return (
    <div className="activity-dialog-backdrop" role="presentation" onMouseDown={onClose}>
      <section
        className="activity-dialog"
        role="dialog"
        aria-modal="true"
        aria-label="OneDrive 통신 상태 로그"
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="activity-dialog-header">
          <div>
            <h2>상태 로그</h2>
            <span>{events.length.toLocaleString('ko-KR')}개 이벤트</span>
          </div>
          <button className="activity-dialog-close" type="button" aria-label="닫기" onClick={onClose}>
            x
          </button>
        </header>
        <div className="activity-log-list">
          {events.length > 0 ? (
            events.map((event) => (
              <article className="activity-log-item" data-level={event.level} key={event.id}>
                <div>
                  <strong>{event.title}</strong>
                  <time>{formatDateTime(event.at)}</time>
                </div>
                <p>{event.message}</p>
                <small>
                  {formatGraphActivityScope(event.scope)}
                  {event.status ? ` · HTTP ${event.status}` : ''}
                  {event.retryAfterMs ? ` · 재시도 ${formatRetryDelayLabel(event.retryAfterMs)}` : ''}
                  {event.progress ? ` · ${formatGraphActivityProgressLabel(event)}` : ''}
                </small>
                {event.progress ? <GraphActivityProgressBar event={event} /> : null}
                {event.detail ? <pre>{event.detail}</pre> : null}
              </article>
            ))
          ) : (
            <div className="activity-log-empty">아직 기록된 OneDrive 통신 이벤트가 없습니다.</div>
          )}
        </div>
      </section>
    </div>
  )
}

function GraphActivityProgressBar({ event }: { event: GraphActivityEvent }): ReactElement {
  const progressPercent = getGraphActivityProgressPercent(event)

  return (
    <div className="activity-progress" data-indeterminate={event.progress?.indeterminate ? 'true' : 'false'} aria-hidden="true">
      <span style={progressPercent === null ? undefined : { width: `${progressPercent}%` }} />
    </div>
  )
}

function AuthSettingsForm({
  clientId,
  tenantId,
  source,
  isSaving,
  isDisabled,
  error,
  message,
  onChange,
  onSave
}: {
  clientId: string
  tenantId: string
  source: MicrosoftAuthSettingsSource
  isSaving: boolean
  isDisabled: boolean
  error: string | null
  message: string | null
  onChange: (nextForm: { clientId: string; tenantId: string }) => void
  onSave: () => void
}): ReactElement {
  return (
    <form
      className="settings-form"
      onSubmit={(event) => {
        event.preventDefault()
        onSave()
      }}
    >
      <div className="settings-form-heading">
        <strong>로그인 설정</strong>
        <span>{getSettingsSourceLabel(source)}</span>
      </div>
      <label>
        <span>Application client ID</span>
        <input
          type="text"
          value={clientId}
          placeholder="00000000-0000-0000-0000-000000000000"
          spellCheck={false}
          disabled={isDisabled || isSaving}
          onChange={(event) => onChange({ clientId: event.target.value, tenantId })}
        />
      </label>
      <label>
        <span>Tenant</span>
        <input
          type="text"
          value={tenantId}
          placeholder="consumers"
          spellCheck={false}
          disabled={isDisabled || isSaving}
          onChange={(event) => onChange({ clientId, tenantId: event.target.value })}
        />
      </label>
      <div className="tenant-presets" aria-label="Tenant 빠른 선택">
        <button
          type="button"
          className={tenantId === 'consumers' ? 'selected' : undefined}
          disabled={isDisabled || isSaving}
          onClick={() => onChange({ clientId, tenantId: 'consumers' })}
        >
          개인
        </button>
        <button
          type="button"
          className={tenantId === 'common' ? 'selected' : undefined}
          disabled={isDisabled || isSaving}
          onClick={() => onChange({ clientId, tenantId: 'common' })}
        >
          공통
        </button>
        <button
          type="button"
          className={tenantId === 'organizations' ? 'selected' : undefined}
          disabled={isDisabled || isSaving}
          onClick={() => onChange({ clientId, tenantId: 'organizations' })}
        >
          조직
        </button>
      </div>
      <div className="redirect-uri-note">
        <span>Redirect URI</span>
        <code>http://localhost</code>
      </div>
      <button className="command-button full" type="submit" disabled={isDisabled || isSaving}>
        {isSaving ? '저장 중' : '설정 저장'}
      </button>
      {message ? <p className="inline-message">{message}</p> : null}
      {error ? <p className="inline-error">{error}</p> : null}
    </form>
  )
}

function DriveTabPane({
  tab,
  account,
  isActive,
  canUseDrive,
  canSignIn,
  canClose,
  isConnecting,
  isBusy,
  selectedItemIds,
  cutItemIds,
  dragOverFolderId,
  currentFolderId,
  onActivate,
  onClose,
  onNewTab,
  onSignIn,
  onOpenParent,
  onOpenBreadcrumb,
  onClearSelection,
  onBoxSelect,
  onSortOptionsChange,
  onColumnWidthsChange,
  onViewModeChange,
  onSelectItem,
  onOpenFolder,
  onPreviewItem,
  onRenameItem,
  onDeleteItem,
  onDragOverFolderChange,
  onMoveItemsToFolder,
  onUploadDroppedFiles,
  onContextMenuItem,
  onContextMenuBackground,
  onLoadMore
}: {
  tab: DriveTab
  account: AuthAccount | null
  isActive: boolean
  canUseDrive: boolean
  canSignIn: boolean
  canClose: boolean
  isConnecting: boolean
  isBusy: boolean
  selectedItemIds: string[]
  cutItemIds: string[]
  dragOverFolderId: string | null
  currentFolderId: string | null
  onActivate: () => void
  onClose: () => void
  onNewTab: () => void
  onSignIn: () => void
  onOpenParent: () => void
  onOpenBreadcrumb: (index: number) => void
  onClearSelection: () => void
  onBoxSelect: (itemIds: string[]) => void
  onSortOptionsChange: (nextSortOptions: DriveSortOptions) => void
  onColumnWidthsChange: (nextColumnWidths: DriveColumnWidths) => void
  onViewModeChange: (nextViewMode: DriveViewMode) => void
  onSelectItem: (item: CloudDriveItem, mode: DriveSelectionMode) => void
  onOpenFolder: (item: CloudDriveItem) => void
  onPreviewItem: (item: CloudDriveItem) => void
  onRenameItem: (item: CloudDriveItem) => void
  onDeleteItem: (item: CloudDriveItem) => void
  onDragOverFolderChange: (folderId: string | null) => void
  onMoveItemsToFolder: (folder: CloudDriveItem, payload: DriveDragPayload) => void
  onUploadDroppedFiles: (parentId: string | null, files: File[]) => void
  onContextMenuItem: (item: CloudDriveItem, point: ContextMenuState) => void
  onContextMenuBackground: (point: ContextMenuState) => void
  onLoadMore: () => void
}): ReactElement {
  const currentFolder = tab.folderPath.at(-1) ?? rootFolder
  const sortedDriveState = getSortedDriveState(tab.driveState, tab.sortOptions)

  return (
    <section className={`drive-pane${isActive ? ' active' : ''}`} aria-label={`${getDriveTabTitle(tab)} 파일 목록`} onMouseDown={onActivate}>
      <div className="drive-pane-tabs">
        <button className="drive-pane-tab" type="button" aria-selected={isActive} onClick={onActivate}>
          <span>{getDriveTabTitle(tab)}</span>
          <small>{account ? getAccountOptionLabel(account) : '계정 없음'}</small>
        </button>
        <button
          className="drive-pane-tool"
          type="button"
          title="새 탭"
          aria-label="새 탭"
          disabled={isBusy || !canUseDrive}
          onMouseDown={(event) => event.stopPropagation()}
          onClick={(event) => {
            event.stopPropagation()
            onNewTab()
          }}
        >
          +
        </button>
        <button
          className="drive-pane-tool"
          type="button"
          title="탭 닫기"
          aria-label="탭 닫기"
          disabled={!canClose || isBusy}
          onMouseDown={(event) => event.stopPropagation()}
          onClick={(event) => {
            event.stopPropagation()
            onClose()
          }}
        >
          x
        </button>
      </div>

      <div className="pane-address-row" aria-label="경로 이동">
        <button
          className="nav-button"
          type="button"
          title="상위 폴더"
          aria-label="상위 폴더"
          disabled={!canUseDrive || tab.folderPath.length <= 1 || tab.driveState.status === 'loading'}
          onClick={onOpenParent}
        >
          ↑
        </button>
        <Breadcrumb path={tab.folderPath} isDisabled={!canUseDrive || tab.driveState.status === 'loading'} onSelect={onOpenBreadcrumb} />
      </div>

      <DriveSortControls
        sortOptions={tab.sortOptions}
        viewMode={tab.viewMode}
        isDisabled={!canUseDrive}
        onChange={onSortOptionsChange}
        onViewModeChange={onViewModeChange}
      />

      <DriveExplorer
        accountId={tab.accountId}
        state={sortedDriveState}
        viewMode={tab.viewMode}
        sortOptions={tab.sortOptions}
        columnWidths={tab.columnWidths}
        selectedItemIds={selectedItemIds}
        cutItemIds={cutItemIds}
        dragOverFolderId={dragOverFolderId}
        currentFolderId={currentFolderId}
        isAuthenticated={canUseDrive}
        canSignIn={canSignIn}
        isConnecting={isConnecting}
        onSignIn={onSignIn}
        onClearSelection={onClearSelection}
        onBoxSelect={onBoxSelect}
        onSortOptionsChange={onSortOptionsChange}
        onColumnWidthsChange={onColumnWidthsChange}
        onSelectItem={onSelectItem}
        onOpenFolder={onOpenFolder}
        onPreviewItem={onPreviewItem}
        onRenameItem={onRenameItem}
        onDeleteItem={onDeleteItem}
        onDragOverFolderChange={onDragOverFolderChange}
        onMoveItemsToFolder={onMoveItemsToFolder}
        onUploadDroppedFiles={onUploadDroppedFiles}
        onContextMenuItem={onContextMenuItem}
        onContextMenuBackground={onContextMenuBackground}
        onLoadMore={onLoadMore}
      />

      <div className="pane-status">
        <span>{getDriveStatusText(sortedDriveState, tab.indexState)}</span>
        <span>{currentFolder.name}</span>
      </div>
    </section>
  )
}

function DriveSortControls({
  sortOptions,
  viewMode,
  isDisabled,
  onChange,
  onViewModeChange
}: {
  sortOptions: DriveSortOptions
  viewMode: DriveViewMode
  isDisabled: boolean
  onChange: (nextSortOptions: DriveSortOptions) => void
  onViewModeChange: (nextViewMode: DriveViewMode) => void
}): ReactElement {
  return (
    <div className="pane-sort-row" aria-label="정렬 옵션">
      <label className="sort-select-label">
        <span>정렬</span>
        <select
          className="sort-select"
          value={sortOptions.field ?? 'none'}
          disabled={isDisabled}
          onChange={(event) =>
            onChange({
              ...sortOptions,
              field: event.target.value === 'none' ? null : (event.target.value as DriveSortField)
            })
          }
        >
          <option value="none">정렬 없음</option>
          {(Object.keys(driveSortFieldLabels) as DriveSortField[]).map((field) => (
            <option key={field} value={field}>
              {driveSortFieldLabels[field]}
            </option>
          ))}
        </select>
      </label>
      <button
        className="sort-direction-button"
        type="button"
        disabled={isDisabled || !sortOptions.field}
        onClick={() => onChange({ ...sortOptions, direction: sortOptions.direction === 'asc' ? 'desc' : 'asc' })}
      >
        {sortOptions.field ? (sortOptions.direction === 'asc' ? '오름차순' : '내림차순') : '방향 없음'}
      </button>
      <label className="sort-toggle">
        <input
          type="checkbox"
          checked={sortOptions.foldersFirst}
          disabled={isDisabled}
          onChange={(event) => onChange({ ...sortOptions, foldersFirst: event.target.checked })}
        />
        <span>폴더 먼저</span>
      </label>
      <div className="view-mode-toggle" aria-label="보기 방식">
        {(Object.keys(driveViewModeLabels) as DriveViewMode[]).map((mode) => (
          <button
            type="button"
            key={mode}
            className="view-mode-button"
            aria-pressed={viewMode === mode}
            aria-label={`${driveViewModeLabels[mode]} 보기`}
            title={`${driveViewModeLabels[mode]} 보기`}
            disabled={isDisabled}
            onClick={() => onViewModeChange(mode)}
          >
            <span className={`view-mode-icon ${mode}`} aria-hidden="true" />
          </button>
        ))}
      </div>
    </div>
  )
}

function Breadcrumb({
  path,
  isDisabled,
  onSelect
}: {
  path: DriveFolderRef[]
  isDisabled: boolean
  onSelect: (index: number) => void
}): ReactElement {
  return (
    <nav className="breadcrumb" aria-label="OneDrive 경로">
      {path.map((folder, index) => (
        <button
          key={`${folder.id ?? 'root'}-${index}`}
          type="button"
          disabled={isDisabled}
          aria-current={index === path.length - 1 ? 'page' : undefined}
          onClick={() => onSelect(index)}
        >
          {folder.name}
        </button>
      ))}
    </nav>
  )
}

function DriveExplorer({
  accountId,
  state,
  viewMode,
  sortOptions,
  columnWidths,
  selectedItemIds,
  cutItemIds,
  dragOverFolderId,
  currentFolderId,
  isAuthenticated,
  canSignIn,
  isConnecting,
  onSignIn,
  onClearSelection,
  onBoxSelect,
  onSortOptionsChange,
  onColumnWidthsChange,
  onSelectItem,
  onOpenFolder,
  onPreviewItem,
  onRenameItem,
  onDeleteItem,
  onDragOverFolderChange,
  onMoveItemsToFolder,
  onUploadDroppedFiles,
  onContextMenuItem,
  onContextMenuBackground,
  onLoadMore
}: {
  accountId: string | null
  state: DriveState
  viewMode: DriveViewMode
  sortOptions: DriveSortOptions
  columnWidths: DriveColumnWidths
  selectedItemIds: string[]
  cutItemIds: string[]
  dragOverFolderId: string | null
  currentFolderId: string | null
  isAuthenticated: boolean
  canSignIn: boolean
  isConnecting: boolean
  onSignIn: () => void
  onClearSelection: () => void
  onBoxSelect: (itemIds: string[]) => void
  onSortOptionsChange: (nextSortOptions: DriveSortOptions) => void
  onColumnWidthsChange: (nextColumnWidths: DriveColumnWidths) => void
  onSelectItem: (item: CloudDriveItem, mode: DriveSelectionMode) => void
  onOpenFolder: (item: CloudDriveItem) => void
  onPreviewItem: (item: CloudDriveItem) => void
  onRenameItem: (item: CloudDriveItem) => void
  onDeleteItem: (item: CloudDriveItem) => void
  onDragOverFolderChange: (folderId: string | null) => void
  onMoveItemsToFolder: (folder: CloudDriveItem, payload: DriveDragPayload) => void
  onUploadDroppedFiles: (parentId: string | null, files: File[]) => void
  onContextMenuItem: (item: CloudDriveItem, point: ContextMenuState) => void
  onContextMenuBackground: (point: ContextMenuState) => void
  onLoadMore: () => void
}): ReactElement {
  const bodyRef = useRef<HTMLDivElement | null>(null)
  const columnResizeSessionRef = useRef<ColumnResizeSession | null>(null)
  const columnHeaderClickTimerRef = useRef<number | null>(null)
  const boxSelectionSessionRef = useRef<BoxSelectionSession | null>(null)
  const suppressNextClickRef = useRef(false)
  const [boxSelection, setBoxSelection] = useState<BoxSelectionState | null>(null)
  const columnGridTemplate = getDriveColumnGridTemplate(columnWidths)
  const columnGridMinWidth = getDriveColumnTotalWidth(columnWidths)
  const emptyMessage =
    state.status === 'loading'
      ? '파일 목록을 불러오는 중입니다.'
      : state.status === 'error'
        ? state.message
        : '비어있는 폴더입니다.'

  useEffect(() => {
    return () => {
      if (columnHeaderClickTimerRef.current) {
        window.clearTimeout(columnHeaderClickTimerRef.current)
      }
    }
  }, [])

  function handleRowDragStart(event: DragEvent<HTMLDivElement>, item: CloudDriveItem): void {
    if (!isAuthenticated) {
      event.preventDefault()
      return
    }

    const draggedItemIds = selectedItemIds.includes(item.id) ? selectedItemIds : [item.id]
    const draggedItemIdSet = new Set(draggedItemIds)
    const draggedItems = state.items.filter((candidate) => draggedItemIdSet.has(candidate.id))
    const payload: DriveDragPayload = {
      itemIds: draggedItemIds,
      sourceAccountId: accountId,
      items: draggedItems.map((candidate) => ({
        itemId: candidate.id,
        name: candidate.name,
        type: candidate.type,
        size: candidate.size
      }))
    }

    if (!selectedItemIds.includes(item.id)) {
      onSelectItem(item, 'replace')
    }

    event.dataTransfer.effectAllowed = 'copyMove'
    event.dataTransfer.setData(DRIVE_ITEM_DRAG_TYPE, JSON.stringify(payload))
    event.dataTransfer.setData('text/plain', draggedItemIds.join('\n'))
  }

  function handleDragOverCurrentFolder(event: DragEvent<HTMLElement>): void {
    if (!isAuthenticated || !hasDroppedFiles(event)) {
      return
    }

    event.preventDefault()
    event.dataTransfer.dropEffect = 'copy'
    onDragOverFolderChange(null)
  }

  function handleDropOnCurrentFolder(event: DragEvent<HTMLElement>): void {
    if (!isAuthenticated || !hasDroppedFiles(event)) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    onDragOverFolderChange(null)
    onUploadDroppedFiles(currentFolderId, Array.from(event.dataTransfer.files))
  }

  function canDropOnFolder(event: DragEvent<HTMLElement>, item: CloudDriveItem): boolean {
    if (!isAuthenticated || item.type !== 'folder') {
      return false
    }

    if (hasDroppedFiles(event)) {
      return true
    }

    return hasDraggedDriveItems(event) && !selectedItemIds.includes(item.id)
  }

  function handleRowDragOver(event: DragEvent<HTMLDivElement>, item: CloudDriveItem): void {
    if (!canDropOnFolder(event, item)) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    event.dataTransfer.dropEffect = hasDroppedFiles(event) ? 'copy' : 'move'
    onDragOverFolderChange(item.id)
  }

  function handleRowDrop(event: DragEvent<HTMLDivElement>, item: CloudDriveItem): void {
    if (!canDropOnFolder(event, item)) {
      return
    }

    event.preventDefault()
    event.stopPropagation()
    onDragOverFolderChange(null)

    if (hasDroppedFiles(event)) {
      onUploadDroppedFiles(item.id, Array.from(event.dataTransfer.files))
      return
    }

    onMoveItemsToFolder(item, getDraggedDrivePayload(event))
  }

  function handleColumnHeaderClick(columnKey: DriveColumnKey): void {
    if (columnHeaderClickTimerRef.current) {
      window.clearTimeout(columnHeaderClickTimerRef.current)
    }

    columnHeaderClickTimerRef.current = window.setTimeout(() => {
      onSortOptionsChange(getNextHeaderSortOptions(sortOptions, columnKey))
      columnHeaderClickTimerRef.current = null
    }, 180)
  }

  function handleColumnHeaderDoubleClick(columnKey: DriveColumnKey): void {
    if (columnHeaderClickTimerRef.current) {
      window.clearTimeout(columnHeaderClickTimerRef.current)
      columnHeaderClickTimerRef.current = null
    }

    onColumnWidthsChange({
      ...columnWidths,
      [columnKey]: getAutoFitColumnWidth(columnKey, state.items)
    })
  }

  function handleColumnResizeMouseDown(event: ReactMouseEvent<HTMLButtonElement>, columnKey: DriveColumnKey): void {
    event.preventDefault()
    event.stopPropagation()
    columnResizeSessionRef.current = {
      columnKey,
      originX: event.clientX,
      originWidth: columnWidths[columnKey],
      initialWidths: columnWidths,
      bodyCursor: document.body.style.cursor,
      bodyUserSelect: document.body.style.userSelect
    }
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    window.addEventListener('mousemove', handleColumnResizeMouseMove)
    window.addEventListener('mouseup', handleColumnResizeMouseUp, { once: true })
  }

  function handleColumnResizeMouseMove(event: MouseEvent): void {
    const session = columnResizeSessionRef.current

    if (!session) {
      return
    }

    const minWidth = minDriveColumnWidths[session.columnKey]
    const maxWidth = maxDriveColumnWidths[session.columnKey]
    const nextWidth = Math.round(clampNumber(session.originWidth + event.clientX - session.originX, minWidth, maxWidth))

    onColumnWidthsChange({
      ...session.initialWidths,
      [session.columnKey]: nextWidth
    })
  }

  function handleColumnResizeMouseUp(): void {
    const session = columnResizeSessionRef.current

    window.removeEventListener('mousemove', handleColumnResizeMouseMove)
    columnResizeSessionRef.current = null

    if (session) {
      document.body.style.cursor = session.bodyCursor
      document.body.style.userSelect = session.bodyUserSelect
    }
  }

  function handleColumnResizeDoubleClick(event: ReactMouseEvent<HTMLButtonElement>, columnKey: DriveColumnKey): void {
    event.preventDefault()
    event.stopPropagation()
    onColumnWidthsChange({
      ...columnWidths,
      [columnKey]: getAutoFitColumnWidth(columnKey, state.items)
    })
  }

  function handleBoxSelectionMouseDown(event: ReactMouseEvent<HTMLDivElement>): void {
    if (!canStartBoxSelection(event)) {
      return
    }

    event.preventDefault()
    boxSelectionSessionRef.current = {
      originX: event.clientX,
      originY: event.clientY,
      baseSelectedItemIds: event.ctrlKey || event.metaKey ? selectedItemIds : [],
      isAdditive: event.ctrlKey || event.metaKey,
      hasMoved: false
    }

    window.addEventListener('mousemove', handleBoxSelectionMouseMove)
    window.addEventListener('mouseup', handleBoxSelectionMouseUp, { once: true })
  }

  function handleBoxSelectionMouseMove(event: MouseEvent): void {
    const session = boxSelectionSessionRef.current

    if (!session) {
      return
    }

    const hasMoved = Math.abs(event.clientX - session.originX) > 4 || Math.abs(event.clientY - session.originY) > 4

    if (!hasMoved) {
      return
    }

    session.hasMoved = true
    const nextSelection = {
      originX: session.originX,
      originY: session.originY,
      currentX: event.clientX,
      currentY: event.clientY
    }
    const intersectingItemIds = getIntersectingDriveItemIds(getBoxSelectionRect(nextSelection))
    const selectedIds = session.isAdditive ? Array.from(new Set([...session.baseSelectedItemIds, ...intersectingItemIds])) : intersectingItemIds

    setBoxSelection(nextSelection)
    onBoxSelect(selectedIds)
  }

  function handleBoxSelectionMouseUp(): void {
    const didMove = boxSelectionSessionRef.current?.hasMoved ?? false

    window.removeEventListener('mousemove', handleBoxSelectionMouseMove)
    boxSelectionSessionRef.current = null
    setBoxSelection(null)

    if (didMove) {
      suppressNextClickRef.current = true
      window.setTimeout(() => {
        suppressNextClickRef.current = false
      }, 0)
    }
  }

  function canStartBoxSelection(event: ReactMouseEvent<HTMLDivElement>): boolean {
    if (!isAuthenticated || event.button !== 0) {
      return false
    }

    const target = event.target

    if (!(target instanceof Element) || target.closest('button,input,select,textarea,a')) {
      return false
    }

    const largeIconItem = target.closest('.large-icon-item')

    if (largeIconItem) {
      return false
    }

    const row = target.closest('.details-row')

    return !row || !target.closest('.name-cell')
  }

  function getIntersectingDriveItemIds(rect: DOMRect): string[] {
    const body = bodyRef.current

    if (!body) {
      return []
    }

    return Array.from(body.querySelectorAll<HTMLElement>('[data-drive-item-id]')).flatMap((row) => {
      const itemId = row.dataset.driveItemId
      const rowRect = row.getBoundingClientRect()
      const intersects = rect.left <= rowRect.right && rect.right >= rowRect.left && rect.top <= rowRect.bottom && rect.bottom >= rowRect.top

      return itemId && intersects ? [itemId] : []
    })
  }

  function suppressClickAfterBoxSelection(event: ReactMouseEvent<HTMLElement>): boolean {
    if (!suppressNextClickRef.current) {
      return false
    }

    event.preventDefault()
    event.stopPropagation()
    return true
  }

  return (
    <div
      className={`drive-explorer ${viewMode === 'details' ? 'details-view' : 'large-icons-view'}`}
      role={viewMode === 'details' ? 'table' : 'listbox'}
      aria-multiselectable={viewMode === 'large-icons' ? true : undefined}
      aria-label="OneDrive 파일 목록"
      onContextMenu={(event) => {
        event.preventDefault()
        onContextMenuBackground({ x: event.clientX, y: event.clientY })
      }}
      onDragOver={handleDragOverCurrentFolder}
      onDrop={handleDropOnCurrentFolder}
      onDragLeave={(event) => {
        if (!(event.relatedTarget instanceof Node) || !event.currentTarget.contains(event.relatedTarget)) {
          onDragOverFolderChange(null)
        }
      }}
    >
      {viewMode === 'details' ? (
        <div className="details-header" role="row" style={{ gridTemplateColumns: columnGridTemplate, minWidth: columnGridMinWidth }}>
          {(Object.keys(driveColumnLabels) as DriveColumnKey[]).map((columnKey) => {
            const sortState = getColumnSortState(sortOptions, columnKey)

            return (
              <span className={`details-column-header${sortState !== 'none' ? ' sorted' : ''}`} role="columnheader" aria-sort={sortState} key={columnKey}>
                <button
                  className="column-sort-button"
                  type="button"
                  title={`${driveColumnLabels[columnKey]} 정렬 전환`}
                  onClick={() => handleColumnHeaderClick(columnKey)}
                  onDoubleClick={() => handleColumnHeaderDoubleClick(columnKey)}
                >
                  <span className="column-title">{driveColumnLabels[columnKey]}</span>
                  <span className="column-sort-indicator" aria-hidden="true">
                    {sortState === 'ascending' ? '↑' : sortState === 'descending' ? '↓' : ''}
                  </span>
                </button>
                <button
                  className="column-resize-handle"
                  type="button"
                  aria-label={`${driveColumnLabels[columnKey]} 칼럼 폭 조절`}
                  onMouseDown={(event) => handleColumnResizeMouseDown(event, columnKey)}
                  onClick={(event) => {
                    event.preventDefault()
                    event.stopPropagation()
                  }}
                  onDoubleClick={(event) => handleColumnResizeDoubleClick(event, columnKey)}
                />
              </span>
            )
          })}
        </div>
      ) : null}

      <div
        ref={bodyRef}
        className={viewMode === 'details' ? 'details-body' : 'large-icons-body'}
        style={viewMode === 'details' ? { minWidth: columnGridMinWidth } : undefined}
        onMouseDown={handleBoxSelectionMouseDown}
        onClickCapture={(event) => {
          suppressClickAfterBoxSelection(event)
        }}
        onClick={(event) => {
          if (event.currentTarget === event.target) {
            onClearSelection()
          }
        }}
      >
        {state.items.length > 0 ? (
          state.items.map((item) => {
            const isSelected = selectedItemIds.includes(item.id)
            const isCut = cutItemIds.includes(item.id)
            const isDropTarget = item.type === 'folder' && dragOverFolderId === item.id

            if (viewMode === 'large-icons') {
              const previewKind = getDriveItemPreviewKind(item)

              return (
                <div
                  className={`large-icon-item${isSelected ? ' selected' : ''}${isCut ? ' cut' : ''}${isDropTarget ? ' drop-target' : ''}`}
                  role="option"
                  tabIndex={0}
                  aria-selected={isSelected}
                  key={item.id}
                  data-drive-item-id={item.id}
                  draggable={isAuthenticated}
                  onDragStart={(event) => handleRowDragStart(event, item)}
                  onDragEnd={() => onDragOverFolderChange(null)}
                  onDragOver={(event) => handleRowDragOver(event, item)}
                  onDragLeave={(event) => {
                    if (!(event.relatedTarget instanceof Node) || !event.currentTarget.contains(event.relatedTarget)) {
                      onDragOverFolderChange(null)
                    }
                  }}
                  onDrop={(event) => handleRowDrop(event, item)}
                  onClick={(event) => {
                    if (suppressClickAfterBoxSelection(event)) {
                      return
                    }

                    const mode = event.shiftKey ? 'range' : event.metaKey || event.ctrlKey ? 'toggle' : 'replace'

                    onSelectItem(item, mode)
                  }}
                  onContextMenu={(event) => {
                    event.preventDefault()
                    event.stopPropagation()
                    onContextMenuItem(item, { x: event.clientX, y: event.clientY })
                  }}
                  onDoubleClick={() => {
                    if (item.type === 'folder') {
                      onOpenFolder(item)
                      return
                    }

                    onPreviewItem(item)
                  }}
                  onKeyDown={(event) => {
                    if (event.key === 'Enter') {
                      event.preventDefault()

                      if (item.type === 'folder') {
                        onOpenFolder(item)
                        return
                      }

                      onPreviewItem(item)
                    }

                    if (event.key === 'F2') {
                      event.preventDefault()
                      onRenameItem(item)
                    }

                    if (event.key === 'Delete') {
                      event.preventDefault()
                      onDeleteItem(item)
                    }
                  }}
                >
                  <DriveThumbnailPreview accountId={accountId} item={item} previewKind={previewKind} />
                  <span className="large-icon-name" title={item.name}>
                    {item.name}
                  </span>
                  <span className="large-icon-meta">{getLargeIconMeta(item)}</span>
                </div>
              )
            }

            return (
              <div
                className={`details-row${isSelected ? ' selected' : ''}${isCut ? ' cut' : ''}${isDropTarget ? ' drop-target' : ''}`}
                role="row"
                tabIndex={0}
                aria-selected={isSelected}
                key={item.id}
                data-drive-item-id={item.id}
                style={{ gridTemplateColumns: columnGridTemplate, minWidth: columnGridMinWidth }}
                draggable={isAuthenticated}
                onDragStart={(event) => handleRowDragStart(event, item)}
                onDragEnd={() => onDragOverFolderChange(null)}
                onDragOver={(event) => handleRowDragOver(event, item)}
                onDragLeave={(event) => {
                  if (!(event.relatedTarget instanceof Node) || !event.currentTarget.contains(event.relatedTarget)) {
                    onDragOverFolderChange(null)
                  }
                }}
                onDrop={(event) => handleRowDrop(event, item)}
                onClick={(event) => {
                  if (suppressClickAfterBoxSelection(event)) {
                    return
                  }

                  const mode = event.shiftKey ? 'range' : event.metaKey || event.ctrlKey ? 'toggle' : 'replace'

                  onSelectItem(item, mode)
                }}
                onContextMenu={(event) => {
                  event.preventDefault()
                  event.stopPropagation()
                  onContextMenuItem(item, { x: event.clientX, y: event.clientY })
                }}
                onDoubleClick={() => {
                  if (item.type === 'folder') {
                    onOpenFolder(item)
                    return
                  }

                  onPreviewItem(item)
                }}
                onKeyDown={(event) => {
                  if (event.key === 'Enter') {
                    event.preventDefault()

                    if (item.type === 'folder') {
                      onOpenFolder(item)
                      return
                    }

                    onPreviewItem(item)
                  }

                  if (event.key === 'F2') {
                    event.preventDefault()
                    onRenameItem(item)
                  }

                  if (event.key === 'Delete') {
                    event.preventDefault()
                    onDeleteItem(item)
                  }
                }}
              >
                <span role="cell" className="name-cell">
                  <DriveItemIcon item={item} />
                  <span className="item-name" title={item.name}>
                    {item.name}
                  </span>
                </span>
                <span role="cell">{formatDate(item.lastModifiedDateTime)}</span>
                <span role="cell">{getTypeLabel(item)}</span>
                <span role="cell">{item.type === 'folder' ? '' : formatBytes(item.size)}</span>
              </div>
            )
          })
        ) : (
          <div
            className="explorer-empty"
            role={viewMode === 'details' ? 'row' : 'presentation'}
            style={viewMode === 'details' ? { minWidth: columnGridMinWidth } : undefined}
          >
            {!isAuthenticated ? (
              <div className="login-empty">
                <strong>OneDrive 파일을 보려면 로그인하세요.</strong>
                <button className="command-button primary" type="button" disabled={!canSignIn || isConnecting} onClick={onSignIn}>
                  {isConnecting ? '로그인 중' : '로그인'}
                </button>
              </div>
            ) : (
              <span>{emptyMessage}</span>
            )}
          </div>
        )}
      </div>

      {boxSelection ? <div className="selection-box" style={getBoxSelectionStyle(boxSelection)} /> : null}

      {state.nextLink ? (
        <div className="explorer-footer" style={viewMode === 'details' ? { minWidth: columnGridMinWidth } : undefined}>
          <button className="command-button" type="button" disabled={state.status === 'loading'} onClick={onLoadMore}>
            더 불러오기
          </button>
        </div>
      ) : null}
    </div>
  )
}

function DriveItemIcon({ item }: { item: CloudDriveItem }): ReactElement {
  const iconKey = getDriveFileIconCacheKey(item)
  const [iconUrl, setIconUrl] = useState<string | null>(() => driveFileIconCache.get(iconKey) ?? null)

  useEffect(() => {
    let isCancelled = false

    setIconUrl(driveFileIconCache.get(iconKey) ?? null)

    if (item.type !== 'file' || driveFileIconCache.has(iconKey)) {
      return () => {
        isCancelled = true
      }
    }

    void loadDriveFileIcon(item, iconKey).then((url) => {
      if (!isCancelled) {
        setIconUrl(url)
      }
    })

    return () => {
      isCancelled = true
    }
  }, [iconKey, item.mimeType, item.name, item.type])

  return (
    <span className={`item-icon ${item.type}${iconUrl ? ' system' : ''}`} aria-hidden="true">
      {iconUrl ? <img className="item-system-icon" src={iconUrl} alt="" draggable={false} /> : null}
    </span>
  )
}

function DriveThumbnailPreview({
  accountId,
  item,
  previewKind
}: {
  accountId: string | null
  item: CloudDriveItem
  previewKind: 'image' | 'video' | 'generic'
}): ReactElement {
  const previewRef = useRef<HTMLSpanElement | null>(null)
  const thumbnailCacheKey = getDriveThumbnailCacheKey(item)
  const canLoadThumbnail = item.type === 'file' && previewKind !== 'generic'
  const [thumbnailState, setThumbnailState] = useState<ThumbnailPreviewState>({ status: canLoadThumbnail ? 'idle' : 'missing' })

  useEffect(() => {
    let isCancelled = false

    setThumbnailState({ status: canLoadThumbnail ? 'idle' : 'missing' })

    if (!canLoadThumbnail) {
      return undefined
    }

    const loadThumbnail = (): void => {
      setThumbnailState({ status: 'loading' })
      void window.oneDriveManager
        .getDriveThumbnail({
          accountId,
          itemId: item.id,
          cacheKey: thumbnailCacheKey,
          priority: 'high',
          size: 'c160x120_crop'
        })
        .then((thumbnail) => {
          if (isCancelled) {
            return
          }

          if (thumbnail.status === 'ready' && thumbnail.url) {
            setThumbnailState({
              status: 'ready',
              url: thumbnail.url,
              width: thumbnail.width,
              height: thumbnail.height
            })
            return
          }

          setThumbnailState({ status: 'missing' })
        })
        .catch(() => {
          if (!isCancelled) {
            setThumbnailState({ status: 'error' })
          }
        })
    }

    const element = previewRef.current

    if (!element || !('IntersectionObserver' in window)) {
      loadThumbnail()
      return () => {
        isCancelled = true
      }
    }

    const observer = new IntersectionObserver(
      (entries) => {
        if (entries.some((entry) => entry.isIntersecting)) {
          observer.disconnect()
          loadThumbnail()
        }
      },
      {
        root: element.closest('.large-icons-body'),
        rootMargin: '280px 120px',
        threshold: 0.01
      }
    )

    observer.observe(element)

    return () => {
      isCancelled = true
      observer.disconnect()
    }
  }, [accountId, canLoadThumbnail, item.id, thumbnailCacheKey])

  return (
    <span ref={previewRef} className="large-icon-preview" data-preview-kind={previewKind} data-thumbnail-state={thumbnailState.status}>
      {thumbnailState.status === 'ready' ? (
        <img
          className="large-icon-preview-image"
          src={thumbnailState.url}
          width={thumbnailState.width}
          height={thumbnailState.height}
          alt=""
          aria-hidden="true"
          decoding="async"
          draggable={false}
          referrerPolicy="no-referrer"
          onError={() => setThumbnailState({ status: 'error' })}
        />
      ) : (
        <span className="large-icon-preview-symbol" aria-hidden="true">
          <DriveItemIcon item={item} />
        </span>
      )}
    </span>
  )
}

function getDriveItemPreviewKind(item: CloudDriveItem): 'image' | 'video' | 'generic' {
  if (item.type !== 'file') {
    return 'generic'
  }

  const mimeType = item.mimeType?.toLocaleLowerCase('en-US') ?? ''

  if (mimeType.startsWith('image/')) {
    return 'image'
  }

  if (mimeType.startsWith('video/')) {
    return 'video'
  }

  const extension = getFileExtension(item.name)

  if (imageFileExtensions.has(extension)) {
    return 'image'
  }

  if (videoFileExtensions.has(extension)) {
    return 'video'
  }

  return 'generic'
}

function getLargeIconMeta(item: CloudDriveItem): string {
  if (item.type === 'folder') {
    return item.childCount === undefined ? getTypeLabel(item) : `${item.childCount.toLocaleString('ko-KR')}개 항목`
  }

  if (item.type === 'package') {
    return getTypeLabel(item)
  }

  return formatBytes(item.size)
}

function getDriveThumbnailCacheKey(item: CloudDriveItem): string {
  return item.cTag ?? item.eTag ?? `${item.size}:${item.lastModifiedDateTime ?? ''}`
}

function getDrivePreviewThumbnailSize(): DriveThumbnailSize {
  const pixelRatio = Math.min(Math.max(window.devicePixelRatio || 1, 1), DRIVE_PREVIEW_THUMBNAIL_MAX_PIXEL_RATIO)
  const width = Math.round(
    clampNumber(
      (window.innerWidth - 64) * pixelRatio,
      DRIVE_PREVIEW_THUMBNAIL_MIN_WIDTH,
      DRIVE_PREVIEW_THUMBNAIL_MAX_WIDTH
    )
  )
  const height = Math.round(
    clampNumber(
      (window.innerHeight - 148) * pixelRatio,
      DRIVE_PREVIEW_THUMBNAIL_MIN_HEIGHT,
      DRIVE_PREVIEW_THUMBNAIL_MAX_HEIGHT
    )
  )

  return `c${width}x${height}`
}

function getFileExtension(name: string): string {
  const extensionIndex = name.lastIndexOf('.')

  if (extensionIndex < 0 || extensionIndex === name.length - 1) {
    return ''
  }

  return name.slice(extensionIndex + 1).toLocaleLowerCase('en-US')
}

function getDriveFileIconCacheKey(item: CloudDriveItem): string {
  if (item.type !== 'file') {
    return item.type
  }

  return ['file', getFileExtension(item.name) || item.mimeType?.toLocaleLowerCase('en-US') || 'generic', 'normal'].join(':')
}

function loadDriveFileIcon(item: CloudDriveItem, iconKey: string): Promise<string | null> {
  if (driveFileIconCache.has(iconKey)) {
    return Promise.resolve(driveFileIconCache.get(iconKey) ?? null)
  }

  let promise = driveFileIconPromises.get(iconKey)

  if (!promise) {
    promise = window.oneDriveManager
      .getDriveFileIcon({
        name: item.name,
        type: item.type,
        mimeType: item.mimeType,
        size: 'normal'
      })
      .then((result) => result.url ?? null)
      .catch(() => null)
      .then((url) => {
        driveFileIconCache.set(iconKey, url)
        return url
      })
      .finally(() => {
        driveFileIconPromises.delete(iconKey)
      })
    driveFileIconPromises.set(iconKey, promise)
  }

  return promise
}

function hasDroppedFiles(event: DragEvent<HTMLElement>): boolean {
  return Array.from(event.dataTransfer.types).includes('Files') && event.dataTransfer.files.length > 0
}

function hasDraggedDriveItems(event: DragEvent<HTMLElement>): boolean {
  return Array.from(event.dataTransfer.types).includes(DRIVE_ITEM_DRAG_TYPE)
}

function getDraggedDrivePayload(event: DragEvent<HTMLElement>): DriveDragPayload {
  try {
    const value = JSON.parse(event.dataTransfer.getData(DRIVE_ITEM_DRAG_TYPE)) as unknown

    if (Array.isArray(value)) {
      return {
        itemIds: value.filter((itemId): itemId is string => typeof itemId === 'string' && itemId.length > 0),
        sourceAccountId: null,
        items: []
      }
    }

    if (value && typeof value === 'object') {
      const payload = value as Partial<DriveDragPayload>
      const itemIds = Array.isArray(payload.itemIds)
        ? payload.itemIds.filter((itemId): itemId is string => typeof itemId === 'string' && itemId.length > 0)
        : []
      const items = Array.isArray(payload.items)
        ? payload.items.filter(
            (item): item is TransferDriveItemRef =>
              Boolean(item) &&
              typeof item === 'object' &&
              typeof (item as TransferDriveItemRef).itemId === 'string' &&
              typeof (item as TransferDriveItemRef).name === 'string'
          )
        : []

      return {
        itemIds,
        sourceAccountId: typeof payload.sourceAccountId === 'string' ? payload.sourceAccountId : null,
        items
      }
    }
  } catch {
    // Invalid drag payloads are ignored.
  }

  return {
    itemIds: [],
    sourceAccountId: null,
    items: []
  }
}

function getBoxSelectionRect(selection: BoxSelectionState): DOMRect {
  const left = Math.min(selection.originX, selection.currentX)
  const top = Math.min(selection.originY, selection.currentY)
  const width = Math.abs(selection.currentX - selection.originX)
  const height = Math.abs(selection.currentY - selection.originY)

  return new DOMRect(left, top, width, height)
}

function getBoxSelectionStyle(selection: BoxSelectionState) {
  const rect = getBoxSelectionRect(selection)

  return {
    left: rect.left,
    top: rect.top,
    width: rect.width,
    height: rect.height
  }
}

function createDriveTab(accountId: string | null): DriveTab {
  return {
    id: createTabId(),
    accountId,
    folderPath: [rootFolder],
    driveState: createIdleDriveState(),
    indexState: { status: 'idle' },
    sortOptions: { ...defaultDriveSortOptions },
    viewMode: 'details',
    columnWidths: { ...defaultDriveColumnWidths },
    paneSize: DEFAULT_DRIVE_PANE_SIZE
  }
}

function createIdleDriveState(): DriveState {
  return { status: 'idle', items: [] }
}

function createTabId(): string {
  return globalThis.crypto?.randomUUID?.() ?? `tab-${Date.now()}-${Math.random().toString(16).slice(2)}`
}

function getDriveTabAccountId(tabs: DriveTab[], tabId: string): string | null {
  return tabs.find((tab) => tab.id === tabId)?.accountId ?? null
}

function setAccountForTab(setTabs: Dispatch<SetStateAction<DriveTab[]>>, tabId: string, accountId: string | null): void {
  setTabs((currentTabs) => currentTabs.map((tab) => (tab.id === tabId ? { ...tab, accountId } : tab)))
}

function setFolderPathForTab(setTabs: Dispatch<SetStateAction<DriveTab[]>>, tabId: string, folderPath: DriveFolderRef[]): void {
  setTabs((currentTabs) => currentTabs.map((tab) => (tab.id === tabId ? { ...tab, folderPath } : tab)))
}

function setDriveStateForTab(setTabs: Dispatch<SetStateAction<DriveTab[]>>, tabId: string, update: StateUpdate<DriveState>): void {
  setTabs((currentTabs) =>
    currentTabs.map((tab) =>
      tab.id === tabId
        ? {
            ...tab,
            driveState: resolveStateUpdate(update, tab.driveState)
          }
        : tab
    )
  )
}

function setIndexStateForTab(setTabs: Dispatch<SetStateAction<DriveTab[]>>, tabId: string, update: StateUpdate<IndexState>): void {
  setTabs((currentTabs) =>
    currentTabs.map((tab) =>
      tab.id === tabId
        ? {
            ...tab,
            indexState: resolveStateUpdate(update, tab.indexState)
          }
        : tab
    )
  )
}

function setSortOptionsForTab(setTabs: Dispatch<SetStateAction<DriveTab[]>>, tabId: string, sortOptions: DriveSortOptions): void {
  setTabs((currentTabs) => currentTabs.map((tab) => (tab.id === tabId ? { ...tab, sortOptions } : tab)))
}

function setViewModeForTab(setTabs: Dispatch<SetStateAction<DriveTab[]>>, tabId: string, viewMode: DriveViewMode): void {
  setTabs((currentTabs) => currentTabs.map((tab) => (tab.id === tabId ? { ...tab, viewMode } : tab)))
}

function setColumnWidthsForTab(setTabs: Dispatch<SetStateAction<DriveTab[]>>, tabId: string, columnWidths: DriveColumnWidths): void {
  setTabs((currentTabs) => currentTabs.map((tab) => (tab.id === tabId ? { ...tab, columnWidths } : tab)))
}

function resolveStateUpdate<T>(update: StateUpdate<T>, currentValue: T): T {
  return typeof update === 'function' ? (update as (current: T) => T)(currentValue) : update
}

function getDriveWorkspaceGridTemplateColumns(tabs: DriveTab[]): string {
  if (tabs.length <= 1) {
    return 'minmax(0, 1fr)'
  }

  return tabs
    .flatMap((tab, index) => {
      const paneColumn = `minmax(${MIN_DRIVE_PANE_WIDTH}px, ${Math.max(tab.paneSize, 0.2)}fr)`

      return index === tabs.length - 1 ? [paneColumn] : [paneColumn, `${PANE_RESIZE_HANDLE_WIDTH}px`]
    })
    .join(' ')
}

function getDriveColumnGridTemplate(widths: DriveColumnWidths): string {
  return `${widths.name}px ${widths.modified}px ${widths.type}px ${widths.size}px`
}

function getDriveColumnTotalWidth(widths: DriveColumnWidths): number {
  return widths.name + widths.modified + widths.type + widths.size
}

function clampNumber(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value))
}

function getNextHeaderSortOptions(sortOptions: DriveSortOptions, columnKey: DriveColumnKey): DriveSortOptions {
  if (sortOptions.field !== columnKey) {
    return {
      ...sortOptions,
      field: columnKey,
      direction: 'asc'
    }
  }

  if (sortOptions.direction === 'asc') {
    return {
      ...sortOptions,
      direction: 'desc'
    }
  }

  return {
    ...sortOptions,
    field: null,
    direction: 'asc'
  }
}

function getColumnSortState(sortOptions: DriveSortOptions, columnKey: DriveColumnKey): 'ascending' | 'descending' | 'none' {
  if (sortOptions.field !== columnKey) {
    return 'none'
  }

  return sortOptions.direction === 'asc' ? 'ascending' : 'descending'
}

function getAutoFitColumnWidth(columnKey: DriveColumnKey, items: CloudDriveItem[]): number {
  const textValues = [driveColumnLabels[columnKey], ...items.map((item) => getDriveColumnCellText(item, columnKey))]
  const context = getCanvasMeasureContext()
  const widestText = textValues.reduce((maxWidth, value) => {
    if (!context) {
      return Math.max(maxWidth, value.length * 8)
    }

    return Math.max(maxWidth, context.measureText(value).width)
  }, 0)
  const extraWidth = columnKey === 'name' ? 64 : 42

  return Math.round(clampNumber(widestText + extraWidth, minDriveColumnWidths[columnKey], maxDriveColumnWidths[columnKey]))
}

function getCanvasMeasureContext(): CanvasRenderingContext2D | null {
  if (typeof document === 'undefined') {
    return null
  }

  const canvas = document.createElement('canvas')
  const context = canvas.getContext('2d')

  if (context) {
    context.font = getComputedStyle(document.body).font
  }

  return context
}

function getDriveColumnCellText(item: CloudDriveItem, columnKey: DriveColumnKey): string {
  if (columnKey === 'modified') {
    return formatDate(item.lastModifiedDateTime)
  }

  if (columnKey === 'type') {
    return getTypeLabel(item)
  }

  if (columnKey === 'size') {
    return item.type === 'folder' ? '' : formatBytes(item.size)
  }

  return item.name
}

function getSortedDriveTabItems(tab: DriveTab): CloudDriveItem[] {
  return sortDriveItemsForView(tab.driveState.items, tab.sortOptions)
}

function getDrivePreviewItems(tab: DriveTab | undefined): CloudDriveItem[] {
  return tab ? getSortedDriveTabItems(tab).filter((item) => item.type !== 'folder') : []
}

function getSortedDriveState(state: DriveState, sortOptions: DriveSortOptions): DriveState {
  return {
    ...state,
    items: sortDriveItemsForView(state.items, sortOptions)
  }
}

function sortDriveItemsForView(items: CloudDriveItem[], sortOptions: DriveSortOptions): CloudDriveItem[] {
  return [...items].sort((left, right) => {
    if (sortOptions.foldersFirst) {
      const folderOrder = getFolderSortOrder(left) - getFolderSortOrder(right)

      if (folderOrder !== 0) {
        return folderOrder
      }
    }

    if (!sortOptions.field) {
      return 0
    }

    const fieldComparison = compareDriveItemsByField(left, right, sortOptions.field)
    const directedComparison = sortOptions.direction === 'asc' ? fieldComparison : -fieldComparison

    if (directedComparison !== 0) {
      return directedComparison
    }

    const typeComparison = driveItemTypeOrder[left.type] - driveItemTypeOrder[right.type]

    if (typeComparison !== 0) {
      return typeComparison
    }

    return driveItemNameCollator.compare(left.name, right.name)
  })
}

function compareDriveItemsByField(left: CloudDriveItem, right: CloudDriveItem, field: DriveSortField): number {
  if (field === 'modified') {
    return getDriveItemModifiedTime(left) - getDriveItemModifiedTime(right)
  }

  if (field === 'type') {
    const typeComparison = driveItemTypeOrder[left.type] - driveItemTypeOrder[right.type]

    if (typeComparison !== 0) {
      return typeComparison
    }

    return driveItemNameCollator.compare(getTypeLabel(left), getTypeLabel(right))
  }

  if (field === 'size') {
    return left.size - right.size
  }

  return driveItemNameCollator.compare(left.name, right.name)
}

function getFolderSortOrder(item: CloudDriveItem): number {
  return item.type === 'folder' ? 0 : 1
}

function getDriveItemModifiedTime(item: CloudDriveItem): number {
  const timestamp = item.lastModifiedDateTime ? new Date(item.lastModifiedDateTime).getTime() : 0

  return Number.isFinite(timestamp) ? timestamp : 0
}

function clearFolderCacheForAccount(cache: Map<string, CachedDriveFolder>, accountId: string | null): void {
  const prefix = `${accountId ?? 'none'}::`

  for (const key of cache.keys()) {
    if (key.startsWith(prefix)) {
      cache.delete(key)
    }
  }
}

function getFolderCacheKey(accountId: string | null, folder: DriveFolderRef): string {
  return `${accountId ?? 'none'}::${folder.id ?? 'root'}`
}

function getAccountForTab(session: AuthSession | null, accountId: string | null): AuthAccount | null {
  if (!accountId) {
    return null
  }

  return session?.accounts.find((account) => account.homeAccountId === accountId) ?? null
}

function getDriveTabTitle(tab: DriveTab): string {
  return tab.folderPath.at(-1)?.name ?? rootFolder.name
}

function getAccountLine(session: AuthSession | null, isConnecting: boolean): string {
  if (isConnecting) {
    return '로그인 중'
  }

  const accountCount = session?.accounts.length ?? 0
  const username = session?.account?.username ?? '로그인 전'

  return accountCount > 1 ? `${username} 외 ${accountCount - 1}개 계정` : username
}

function getAccountOptionLabel(account: AuthAccount): string {
  if (account.name && account.name !== account.username) {
    return `${account.name} (${account.username})`
  }

  return account.username
}

function getAccountUsageLabel(usage: DriveAccountUsage | undefined, isLoading: boolean): string {
  if (isLoading && !usage) {
    return '사용량 확인 중'
  }

  if (!usage) {
    return '사용량 미확인'
  }

  if (usage.isUnavailable) {
    return '사용량 확인 실패'
  }

  if (!usage.total || usage.total <= 0) {
    return `${formatBytes(usage.used)} 사용`
  }

  return `${formatBytes(usage.used)} / ${formatBytes(usage.total)}`
}

function getFolderPathLabel(path: DriveFolderRef[]): string {
  return path.map((folder) => folder.name).join(' / ')
}

function createDriveFolderCompareEndpoint(endpoint: FolderCompareEndpointView): DriveFolderCompareEndpoint {
  return {
    accountId: endpoint.accountId,
    folderId: endpoint.folderId,
    folderName: endpoint.folderName
  }
}

function getFolderCompareKindLabel(kind: DriveFolderCompareDifference['kind']): string {
  if (kind === 'only-in-source') {
    return '기준에만 있음'
  }

  if (kind === 'only-in-target') {
    return '대상에만 있음'
  }

  return '서로 다름'
}

function formatFolderCompareItem(item: DriveFolderCompareDifference['source']): string {
  if (!item) {
    return '-'
  }

  const sizeText = item.type === 'folder' ? '' : ` · ${formatBytes(item.size)}`

  return `${getFolderCompareItemTypeLabel(item)}${sizeText}`
}

function formatFolderCompareReasons(difference: DriveFolderCompareDifference): string {
  if (difference.kind !== 'different') {
    return getFolderCompareKindLabel(difference.kind)
  }

  if (difference.reasons.length === 0) {
    return '정보 차이'
  }

  return difference.reasons.map(getFolderCompareReasonLabel).join(', ')
}

function getFolderCompareReasonLabel(reason: DriveFolderCompareDifference['reasons'][number]): string {
  if (reason === 'type') {
    return '종류'
  }

  if (reason === 'content') {
    return '내용 다름'
  }

  return '정보 차이'
}

function getFolderCompareItemTypeLabel(item: DriveFolderCompareDifference['source']): string {
  if (!item) {
    return '-'
  }

  return itemTypeLabels[item.type]
}

function getQuotaStateLabel(state: string): string {
  if (state === 'normal') {
    return '정상'
  }

  if (state === 'nearing') {
    return '부족'
  }

  if (state === 'critical') {
    return '위험'
  }

  if (state === 'exceeded') {
    return '초과'
  }

  return state
}

function getSettingsSourceLabel(source: MicrosoftAuthSettingsSource): string {
  if (source === 'app') {
    return '앱 설정'
  }

  if (source === 'environment') {
    return '환경 변수'
  }

  return '미설정'
}

function getTransferProgress(task: DriveTransferTask): number {
  if (task.totalBytes <= 0) {
    return task.status === 'completed' ? 100 : 0
  }

  return Math.min(100, Math.max(0, Math.round((task.transferredBytes / task.totalBytes) * 100)))
}

function getTransferSummaryProgress(summary: DriveTransferSummary): number {
  if (summary.totalCount <= 0) {
    return 0
  }

  return Math.min(100, Math.max(0, Math.round((summary.completedCount / summary.totalCount) * 100)))
}

function getTransferSummaryLabel(summary: DriveTransferSummary): string {
  return `${summary.completedCount.toLocaleString('ko-KR')}개 완료, ${summary.activeCount.toLocaleString('ko-KR')}개 남음`
}

function getTransferSummaryEtaLabel(summary: DriveTransferSummary): string {
  if (summary.bytesPerSecond <= 0 || summary.totalBytes <= summary.transferredBytes) {
    return '완료 예상 -'
  }

  return `완료 예상 ${formatEtaTime(summary.totalBytes - summary.transferredBytes, summary.bytesPerSecond)}`
}

function getTransferElapsedLabel(task: DriveTransferTask): string {
  const startedAt = new Date(task.createdAt).getTime()

  if (!Number.isFinite(startedAt)) {
    return '경과 -'
  }

  return `경과 ${formatDuration(Date.now() - startedAt)}`
}

function getTransferEtaLabel(task: DriveTransferTask): string {
  if (task.status !== 'running' || !task.bytesPerSecond || task.bytesPerSecond <= 0 || task.totalBytes <= task.transferredBytes) {
    return '완료 예상 -'
  }

  return `완료 예상 ${formatEtaTime(task.totalBytes - task.transferredBytes, task.bytesPerSecond)}`
}

function isTransferTaskResumable(task: DriveTransferTask): boolean {
  return task.status === 'failed' || task.status === 'queued' || task.status === 'paused' || task.status === 'retrying'
}

function getTransferKindLabel(kind: DriveTransferTask['kind']): string {
  if (kind === 'upload') {
    return '업로드'
  }

  if (kind === 'download') {
    return '다운로드'
  }

  return '계정 간 전송'
}

function getTransferStatusLabel(task: DriveTransferTask): string {
  if (task.status === 'completed') {
    return '완료'
  }

  if (task.status === 'running') {
    return `${getTransferProgress(task)}%`
  }

  if (task.status === 'queued') {
    return '대기 중'
  }

  if (task.status === 'retrying') {
    if (task.nextRetryAt) {
      return `재시도 대기 ${formatDate(task.nextRetryAt)}`
    }

    return task.message ?? '자동 재시도 대기'
  }

  if (task.status === 'paused') {
    return '중지됨'
  }

  return task.message ?? '중단됨'
}

function formatGraphActivityStatus(event: GraphActivityEvent): string {
  const prefix = event.level === 'error' ? '오류' : event.level === 'warning' ? '주의' : event.level === 'success' ? '완료' : '상태'

  return `${prefix}: ${event.message}`
}

function formatGraphActivityScope(scope: GraphActivityEvent['scope']): string {
  if (scope === 'index') {
    return '인덱스'
  }

  if (scope === 'thumbnail') {
    return '썸네일'
  }

  if (scope === 'transfer') {
    return '전송'
  }

  return 'OneDrive'
}

function getGraphActivityProgressPercent(event: GraphActivityEvent): number | null {
  const progress = event.progress

  if (!progress || progress.indeterminate || !progress.total || progress.total <= 0 || typeof progress.current !== 'number') {
    return null
  }

  return clampNumber(Math.round((progress.current / progress.total) * 100), 0, 100)
}

function formatGraphActivityProgressLabel(event: GraphActivityEvent): string {
  const progress = event.progress

  if (!progress) {
    return ''
  }

  if (progress.total && progress.total > 0 && typeof progress.current === 'number') {
    return `${progress.current.toLocaleString('ko-KR')} / ${progress.total.toLocaleString('ko-KR')}`
  }

  if (typeof progress.current === 'number') {
    return `${progress.current.toLocaleString('ko-KR')}개 처리`
  }

  return '진행 중'
}

function formatDateTime(value: string): string {
  const date = new Date(value)

  if (Number.isNaN(date.getTime())) {
    return value
  }

  return date.toLocaleString('ko-KR')
}

function formatRetryDelayLabel(delayMs: number): string {
  const seconds = Math.max(1, Math.round(delayMs / 1000))

  return seconds < 60 ? `${seconds}초` : `${Math.round(seconds / 60)}분`
}

function getDriveStatusText(state: DriveState, indexState: IndexState): string {
  if (indexState.status === 'syncing') {
    const indexedItemCount = indexState.index?.itemCount ?? 0

    return indexedItemCount > 0
      ? `${state.items.length.toLocaleString('ko-KR')}개 항목 · 인덱싱 중 ${indexedItemCount.toLocaleString('ko-KR')}개`
      : '탐색 인덱스 구성 중'
  }

  if (indexState.status === 'error') {
    return '탐색 인덱스 오류'
  }

  if (state.status === 'loading') {
    return '목록 불러오는 중'
  }

  if (state.status === 'error') {
    return '목록 오류'
  }

  if (indexState.status === 'ready' && indexState.index.isReady) {
    return `${state.items.length.toLocaleString('ko-KR')}개 항목 · 인덱스 ${indexState.index.itemCount.toLocaleString('ko-KR')}개`
  }

  return `${state.items.length.toLocaleString('ko-KR')}개 항목`
}

function getEnvironmentStatus(state: EnvironmentState): string {
  if (state.status === 'loading') {
    return '환경 확인 중'
  }

  if (state.status === 'error') {
    return '환경 확인 실패'
  }

  return platformLabels[state.environment.platform.name] ?? state.environment.platform.name
}

function validateAuthSettings(clientId: string): string | null {
  if (!clientId || clientId === PLACEHOLDER_CLIENT_ID) {
    return 'Microsoft Entra 앱의 Application client ID를 입력하세요.'
  }

  if (!CLIENT_ID_PATTERN.test(clientId)) {
    return 'Application client ID는 GUID 형식이어야 합니다.'
  }

  return null
}

function validateDriveItemName(name: string): string | null {
  if (!name) {
    return '파일 이름을 입력하세요.'
  }

  if (/[<>:"/\\|?*\u0000-\u001f]/.test(name)) {
    return '파일 이름에 사용할 수 없는 문자가 있습니다.'
  }

  return null
}

function getAvailableNewFolderName(items: CloudDriveItem[]): string {
  const baseName = '새 폴더'
  const occupiedNames = new Set(items.map((item) => normalizeDriveItemNameForConflict(item.name)))

  if (!occupiedNames.has(normalizeDriveItemNameForConflict(baseName))) {
    return baseName
  }

  let index = 2
  let candidate = `${baseName} (${index})`

  while (occupiedNames.has(normalizeDriveItemNameForConflict(candidate))) {
    index += 1
    candidate = `${baseName} (${index})`
  }

  return candidate
}

function normalizeDriveItemNameForConflict(name: string): string {
  return name.trim().toLocaleLowerCase('ko-KR')
}

function getTypeLabel(item: CloudDriveItem): string {
  if (item.type === 'folder') {
    return item.childCount === undefined ? itemTypeLabels.folder : `${itemTypeLabels.folder} (${item.childCount.toLocaleString('ko-KR')})`
  }

  return item.mimeType ?? itemTypeLabels[item.type]
}

function formatBytes(size: number): string {
  if (size <= 0) {
    return '0 B'
  }

  const units = ['B', 'KB', 'MB', 'GB', 'TB']
  const exponent = Math.min(Math.floor(Math.log(size) / Math.log(1024)), units.length - 1)
  const value = size / 1024 ** exponent

  return `${value.toFixed(value >= 10 || exponent === 0 ? 0 : 1)} ${units[exponent]}`
}

function formatDate(value: string | undefined): string {
  if (!value) {
    return '-'
  }

  return new Intl.DateTimeFormat('ko-KR', {
    dateStyle: 'medium',
    timeStyle: 'short'
  }).format(new Date(value))
}

function formatEtaTime(remainingBytes: number, bytesPerSecond: number): string {
  const remainingMs = (Math.max(0, remainingBytes) / Math.max(1, bytesPerSecond)) * 1000
  const targetDate = new Date(Date.now() + remainingMs)
  const targetTime = new Intl.DateTimeFormat('ko-KR', {
    timeStyle: 'short'
  }).format(targetDate)

  return `${targetTime} (${formatDuration(remainingMs)} 남음)`
}

function formatDuration(durationMs: number): string {
  const totalSeconds = Math.max(0, Math.floor(durationMs / 1000))
  const hours = Math.floor(totalSeconds / 3600)
  const minutes = Math.floor((totalSeconds % 3600) / 60)
  const seconds = totalSeconds % 60

  if (hours > 0) {
    return `${hours}시간 ${minutes}분`
  }

  if (minutes > 0) {
    return `${minutes}분 ${seconds}초`
  }

  return `${seconds}초`
}
