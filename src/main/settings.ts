import { app } from 'electron'
import { mkdir, readFile, unlink, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import type {
  MicrosoftAuthSettings,
  MicrosoftAuthSettingsInput,
  MicrosoftAuthSettingsSource,
  TransferSettings,
  TransferSettingsInput
} from '../shared/types'

const SETTINGS_FILE_NAME = 'settings.json'
const DEFAULT_TENANT_ID = 'consumers'
const DEFAULT_MAX_CONCURRENT_TRANSFERS = 4
const MIN_CONCURRENT_TRANSFERS = 1
const MAX_ALLOWED_CONCURRENT_TRANSFERS = 64
const PLACEHOLDER_CLIENT_ID = '00000000-0000-0000-0000-000000000000'
const CLIENT_ID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i
const TENANT_ID_PATTERN = /^[a-z0-9][a-z0-9.-]{0,252}[a-z0-9]$|^(common|organizations|consumers)$/i

type PersistedSettings = {
  version: 1
  microsoft?: {
    clientId?: string
    tenantId?: string
  }
  transfer?: {
    maxConcurrentTransfers?: number
  }
}

export async function getMicrosoftAuthSettings(): Promise<MicrosoftAuthSettings> {
  const persistedSettings = await readPersistedSettings()
  const persistedClientId = normalizeClientId(persistedSettings.microsoft?.clientId ?? '')
  const persistedTenantId = normalizeTenantId(persistedSettings.microsoft?.tenantId ?? '')
  const environmentSettings = getEnvironmentAuthSettings()

  if (environmentSettings.clientId) {
    return createSettings(environmentSettings.clientId, environmentSettings.tenantId, 'environment')
  }

  if (persistedClientId) {
    return createSettings(persistedClientId, persistedTenantId || DEFAULT_TENANT_ID, 'app')
  }

  return createSettings('', persistedTenantId || environmentSettings.tenantId || DEFAULT_TENANT_ID, 'missing')
}

export async function updateMicrosoftAuthSettings(input: MicrosoftAuthSettingsInput): Promise<MicrosoftAuthSettings> {
  const clientId = normalizeClientId(input.clientId)
  const tenantId = normalizeTenantId(input.tenantId ?? DEFAULT_TENANT_ID) || DEFAULT_TENANT_ID

  if (!clientId) {
    throw new Error('Microsoft Entra 앱의 client ID를 입력하세요.')
  }

  if (!CLIENT_ID_PATTERN.test(clientId)) {
    throw new Error('client ID는 GUID 형식이어야 합니다.')
  }

  if (!TENANT_ID_PATTERN.test(tenantId)) {
    throw new Error('tenant 값은 common, organizations, consumers, 테넌트 GUID 또는 도메인 형식이어야 합니다.')
  }

  await writePersistedSettings({
    ...((await readPersistedSettings()) ?? { version: 1 }),
    version: 1,
    microsoft: {
      clientId,
      tenantId
    }
  })

  return getMicrosoftAuthSettings()
}

export async function resetMicrosoftAuthSettings(): Promise<MicrosoftAuthSettings> {
  try {
    await unlink(getSettingsPath())
  } catch {
    // Settings file may not exist on a fresh install.
  }

  return getMicrosoftAuthSettings()
}

export async function getTransferSettings(): Promise<TransferSettings> {
  const persistedSettings = await readPersistedSettings()

  return createTransferSettings(persistedSettings.transfer?.maxConcurrentTransfers)
}

export async function updateTransferSettings(input: TransferSettingsInput): Promise<TransferSettings> {
  const persistedSettings = await readPersistedSettings()
  const maxConcurrentTransfers = normalizeConcurrentTransfers(input.maxConcurrentTransfers)

  await writePersistedSettings({
    ...persistedSettings,
    version: 1,
    transfer: {
      maxConcurrentTransfers
    }
  })

  return getTransferSettings()
}

function getEnvironmentAuthSettings(): { clientId: string; tenantId: string } {
  const clientId = normalizeClientId(
    import.meta.env.MAIN_VITE_MICROSOFT_CLIENT_ID ||
      process.env.MAIN_VITE_MICROSOFT_CLIENT_ID ||
      process.env.ONEDRIVE_MANAGER_CLIENT_ID ||
      ''
  )
  const tenantId =
    normalizeTenantId(
      import.meta.env.MAIN_VITE_MICROSOFT_TENANT_ID ||
        process.env.MAIN_VITE_MICROSOFT_TENANT_ID ||
        process.env.ONEDRIVE_MANAGER_TENANT_ID ||
        DEFAULT_TENANT_ID
    ) || DEFAULT_TENANT_ID

  return {
    clientId,
    tenantId
  }
}

function createSettings(clientId: string, tenantId: string, source: MicrosoftAuthSettingsSource): MicrosoftAuthSettings {
  return {
    clientId,
    tenantId,
    isConfigured: Boolean(clientId),
    source
  }
}

function normalizeClientId(value: string): string {
  const clientId = value.trim()

  return clientId === PLACEHOLDER_CLIENT_ID ? '' : clientId
}

function normalizeTenantId(value: string): string {
  return value.trim() || DEFAULT_TENANT_ID
}

function createTransferSettings(value: number | undefined): TransferSettings {
  return {
    maxConcurrentTransfers: normalizeConcurrentTransfers(value),
    minConcurrentTransfers: MIN_CONCURRENT_TRANSFERS,
    maxAllowedConcurrentTransfers: MAX_ALLOWED_CONCURRENT_TRANSFERS
  }
}

function normalizeConcurrentTransfers(value: number | undefined): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    return DEFAULT_MAX_CONCURRENT_TRANSFERS
  }

  return Math.min(MAX_ALLOWED_CONCURRENT_TRANSFERS, Math.max(MIN_CONCURRENT_TRANSFERS, Math.floor(value)))
}

async function readPersistedSettings(): Promise<PersistedSettings> {
  try {
    return JSON.parse(await readFile(getSettingsPath(), 'utf8')) as PersistedSettings
  } catch {
    return { version: 1 }
  }
}

async function writePersistedSettings(settings: PersistedSettings): Promise<void> {
  const settingsPath = getSettingsPath()

  await mkdir(dirname(settingsPath), { recursive: true })
  await writeFile(settingsPath, JSON.stringify(settings, null, 2), 'utf8')
}

function getSettingsPath(): string {
  return join(app.getPath('userData'), SETTINGS_FILE_NAME)
}
