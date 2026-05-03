import { app, safeStorage, shell } from 'electron'
import * as http from 'node:http'
import { mkdir, readFile, unlink, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import {
  type AccountInfo,
  type AuthenticationResult,
  type AuthorizeResponse,
  type Configuration,
  type ICachePlugin,
  type ILoopbackClient,
  type TokenCacheContext,
  PublicClientApplication
} from '@azure/msal-node'
import { getMicrosoftAuthSettings } from '../settings'
import type { AuthAccount, AuthSession } from '../../shared/types'

const GRAPH_SCOPES = ['User.Read', 'Files.ReadWrite', 'offline_access']
const CACHE_FILE_NAME = 'msal-token-cache.json'
const AUTH_STATE_FILE_NAME = 'auth-state.json'
const INTERACTIVE_AUTH_TIMEOUT_MS = 5 * 60_000
const INTERACTIVE_AUTH_TIMEOUT_MESSAGE = 'Microsoft 로그인이 5분 안에 완료되지 않아 취소되었습니다. 다시 로그인하세요.'

let publicClientApplication: PublicClientApplication | null = null

type StoredCache = {
  version: 1
  encrypted: boolean
  data: string
}

type StoredAuthState = {
  version: 1
  activeAccountId?: string
}

class TimedLoopbackClient implements ILoopbackClient {
  private server: http.Server | undefined
  private timeout: ReturnType<typeof setTimeout> | undefined

  constructor(private readonly timeoutMs: number) {}

  async listenForAuthCode(successTemplate?: string, errorTemplate?: string): Promise<AuthorizeResponse> {
    if (this.server) {
      throw new Error('로그인 콜백 서버가 이미 실행 중입니다.')
    }

    return new Promise<AuthorizeResponse>((resolve, reject) => {
      let didSettle = false

      const settle = (callback: () => void): void => {
        if (didSettle) {
          return
        }

        didSettle = true
        callback()
      }

      const rejectWith = (error: Error): void => {
        settle(() => {
          this.closeServer()
          reject(error)
        })
      }

      this.server = http.createServer((request, response) => {
        const requestUrl = request.url

        if (!requestUrl) {
          response.end(errorTemplate ?? '로그인 콜백 주소를 확인하지 못했습니다.')
          rejectWith(new Error('로그인 콜백 주소를 확인하지 못했습니다.'))
          return
        }

        if (requestUrl === '/') {
          response.end(successTemplate ?? '로그인이 완료되었습니다. 이 창을 닫고 앱으로 돌아가세요.')
          return
        }

        const redirectUri = this.getRedirectUri()
        const parsedUrl = new URL(requestUrl, redirectUri)
        const authCodeResponse = Object.fromEntries(parsedUrl.searchParams.entries()) as AuthorizeResponse

        if (authCodeResponse.code) {
          response.writeHead(302, { location: redirectUri })
          response.end()
          settle(() => resolve(authCodeResponse))
          return
        }

        if (authCodeResponse.error) {
          response.end(errorTemplate ?? `로그인 오류: ${authCodeResponse.error}`)
          settle(() => resolve(authCodeResponse))
          return
        }

        response.end(errorTemplate ?? '로그인 응답에 인증 코드가 없습니다.')
        rejectWith(new Error('로그인 응답에 인증 코드가 없습니다.'))
      })

      this.server.once('error', (error) => {
        rejectWith(error instanceof Error ? error : new Error(String(error)))
      })
      this.server.listen(0, '127.0.0.1')

      this.timeout = setTimeout(() => {
        rejectWith(new Error(INTERACTIVE_AUTH_TIMEOUT_MESSAGE))
      }, this.timeoutMs)
    })
  }

  getRedirectUri(): string {
    if (!this.server || !this.server.listening) {
      throw new Error('로그인 콜백 서버가 아직 준비되지 않았습니다.')
    }

    const address = this.server.address()

    if (!address || typeof address === 'string' || !address.port) {
      this.closeServer()
      throw new Error('로그인 콜백 서버 주소를 확인하지 못했습니다.')
    }

    return `http://localhost:${address.port}`
  }

  closeServer(): void {
    if (this.timeout) {
      clearTimeout(this.timeout)
      this.timeout = undefined
    }

    if (!this.server) {
      return
    }

    this.server.close()

    if (typeof this.server.closeAllConnections === 'function') {
      this.server.closeAllConnections()
    }

    this.server.unref()
    this.server = undefined
  }
}

export function getGraphScopes(): string[] {
  return [...GRAPH_SCOPES]
}

export async function getAuthSession(): Promise<AuthSession> {
  const settings = await getMicrosoftAuthSettings()

  if (!settings.isConfigured) {
    return {
      isConfigured: false,
      isAuthenticated: false,
      accounts: [],
      settings,
      scopes: getGraphScopes(),
      message: 'Microsoft Entra 앱의 client ID를 입력하면 로그인할 수 있습니다.'
    }
  }

  const accounts = await getSortedAccounts()
  const account = await getActiveAccount(accounts)
  const mappedAccounts = accounts.map(mapAccount)

  return {
    isConfigured: true,
    isAuthenticated: Boolean(account),
    account: account ? mapAccount(account) : undefined,
    accounts: mappedAccounts,
    activeAccountId: account?.homeAccountId,
    settings,
    scopes: getGraphScopes()
  }
}

export async function connectAccount(): Promise<AuthSession> {
  const pca = await getPublicClientApplication()

  let result: AuthenticationResult

  try {
    result = await pca.acquireTokenInteractive({
      scopes: GRAPH_SCOPES,
      prompt: 'select_account',
      loopbackClient: new TimedLoopbackClient(INTERACTIVE_AUTH_TIMEOUT_MS),
      openBrowser: async (url) => {
        await shell.openExternal(url)
      },
      successTemplate: createBrowserTemplate('OneDrive 관리자', '로그인이 완료되었습니다. 이 창을 닫고 앱으로 돌아가세요.'),
      errorTemplate: createBrowserTemplate('OneDrive 관리자', '로그인을 완료하지 못했습니다. 앱으로 돌아가 다시 시도하세요.')
    })
  } catch (error) {
    throw new Error(formatInteractiveAuthError(error))
  }

  if (!result?.account) {
    throw new Error('로그인이 완료되지 않았습니다.')
  }

  await writeAuthState({ version: 1, activeAccountId: result.account.homeAccountId })
  return getAuthSession()
}

function formatInteractiveAuthError(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error)

  if (message.includes('redirect_uri') || message.includes('AADSTS50011')) {
    return [
      'Microsoft Entra 앱 등록의 Redirect URI 설정이 맞지 않습니다.',
      'Azure Portal에서 App registrations > Authentication으로 이동한 뒤',
      'Mobile and desktop applications 플랫폼에 http://localhost 를 추가하세요.',
      'Web 플랫폼이 아니라 Mobile and desktop applications에 등록해야 합니다.'
    ].join(' ')
  }

  if (message.includes('userAudience') || message.includes('/common/')) {
    return [
      'Microsoft Entra 앱 등록의 지원 계정 유형과 앱의 Tenant 설정이 맞지 않습니다.',
      '앱 등록이 Personal Microsoft accounts only라면 로그인 설정의 Tenant를 consumers로 저장하세요.',
      'Tenant를 common으로 쓰려면 앱 등록의 Supported account types를 개인 Microsoft 계정과 조직 계정을 모두 허용하도록 바꿔야 합니다.'
    ].join(' ')
  }

  if (message.includes('access_denied') || message.includes('cancel')) {
    return 'Microsoft 로그인이 취소되었습니다.'
  }

  if (message.includes(INTERACTIVE_AUTH_TIMEOUT_MESSAGE) || message.includes('timed out')) {
    return INTERACTIVE_AUTH_TIMEOUT_MESSAGE
  }

  return message || '로그인을 완료하지 못했습니다.'
}

export async function disconnectAccount(): Promise<AuthSession> {
  const pca = await getPublicClientApplication()
  const accounts = await getSortedAccounts()
  const activeAccount = await getActiveAccount(accounts)

  if (activeAccount) {
    await pca.getTokenCache().removeAccount(activeAccount)
  }

  const remainingAccounts = (await getSortedAccounts()).filter((account) => account.homeAccountId !== activeAccount?.homeAccountId)

  if (remainingAccounts.length > 0) {
    await writeAuthState({ version: 1, activeAccountId: remainingAccounts[0]?.homeAccountId })
  } else {
    await clearAuthState()

    try {
      await unlink(getTokenCachePath())
    } catch {
      // Cache file may not exist on first sign out.
    }
  }

  return getAuthSession()
}

export async function switchActiveAccount(accountId: string): Promise<AuthSession> {
  const normalizedAccountId = accountId.trim()

  if (!normalizedAccountId) {
    throw new Error('전환할 계정을 선택하세요.')
  }

  const account = (await getSortedAccounts()).find((candidate) => candidate.homeAccountId === normalizedAccountId)

  if (!account) {
    throw new Error('선택한 Microsoft 계정을 찾지 못했습니다.')
  }

  await writeAuthState({ version: 1, activeAccountId: account.homeAccountId })
  return getAuthSession()
}

export async function getGraphAccessToken(accountId?: string | null): Promise<string> {
  const pca = await getPublicClientApplication()
  const account = accountId ? await getAccountById(accountId) : await getActiveAccount()

  if (!account) {
    throw new Error('Microsoft 계정 로그인이 필요합니다.')
  }

  let result: AuthenticationResult

  try {
    result = await pca.acquireTokenSilent({
      account,
      scopes: GRAPH_SCOPES
    })
  } catch {
    throw new Error('인증이 만료되었거나 OneDrive 쓰기 권한 동의가 필요합니다. 다시 로그인하세요.')
  }

  return result.accessToken
}

export async function getActiveAccountId(): Promise<string | null> {
  try {
    return (await getActiveAccount())?.homeAccountId ?? null
  } catch {
    return null
  }
}

async function getSortedAccounts(): Promise<AccountInfo[]> {
  const pca = await getPublicClientApplication()
  const accounts = await pca.getAllAccounts()

  return [...accounts].sort((left, right) => left.username.localeCompare(right.username, 'ko-KR', { sensitivity: 'base' }))
}

async function getActiveAccount(accounts?: AccountInfo[]): Promise<AccountInfo | null> {
  const availableAccounts = accounts ?? (await getSortedAccounts())

  if (availableAccounts.length === 0) {
    await clearAuthState()
    return null
  }

  const state = await readAuthState()
  const activeAccount = state.activeAccountId
    ? availableAccounts.find((account) => account.homeAccountId === state.activeAccountId)
    : undefined

  if (activeAccount) {
    return activeAccount
  }

  const [fallbackAccount] = availableAccounts

  if (fallbackAccount) {
    await writeAuthState({ version: 1, activeAccountId: fallbackAccount.homeAccountId })
  }

  return fallbackAccount ?? null
}

async function getAccountById(accountId: string): Promise<AccountInfo | null> {
  const normalizedAccountId = accountId.trim()

  if (!normalizedAccountId) {
    return null
  }

  return (await getSortedAccounts()).find((account) => account.homeAccountId === normalizedAccountId) ?? null
}

export function resetAuthClient(): void {
  publicClientApplication = null
}

export async function resetAuthCache(): Promise<void> {
  resetAuthClient()
  await clearAuthState()

  try {
    await unlink(getTokenCachePath())
  } catch {
    // Token cache file may not exist on a fresh install.
  }
}

async function getPublicClientApplication(): Promise<PublicClientApplication> {
  if (publicClientApplication) {
    return publicClientApplication
  }

  const settings = await getMicrosoftAuthSettings()

  if (!settings.isConfigured) {
    throw new Error('Microsoft Entra 앱의 client ID를 입력한 뒤 로그인하세요.')
  }

  const configuration: Configuration = {
    auth: {
      clientId: settings.clientId,
      authority: `https://login.microsoftonline.com/${settings.tenantId}`
    },
    cache: {
      cachePlugin: createEncryptedCachePlugin(getTokenCachePath())
    }
  }

  publicClientApplication = new PublicClientApplication(configuration)
  return publicClientApplication
}

function mapAccount(account: AccountInfo): AuthAccount {
  return {
    homeAccountId: account.homeAccountId,
    username: account.username,
    name: account.name ?? account.username
  }
}

function getTokenCachePath(): string {
  return join(app.getPath('userData'), CACHE_FILE_NAME)
}

function getAuthStatePath(): string {
  return join(app.getPath('userData'), AUTH_STATE_FILE_NAME)
}

function createBrowserTemplate(title: string, message: string): string {
  return `<!doctype html>
<html lang="ko">
  <head>
    <meta charset="utf-8">
    <title>${escapeHtml(title)}</title>
    <style>
      body {
        margin: 0;
        min-height: 100vh;
        display: grid;
        place-items: center;
        background: #f4f7f6;
        color: #202124;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Apple SD Gothic Neo", "Malgun Gothic", sans-serif;
      }
      main {
        width: min(520px, calc(100vw - 48px));
        border: 1px solid #d7dee2;
        border-radius: 8px;
        padding: 28px;
        background: #ffffff;
        box-shadow: 0 12px 28px rgba(32, 45, 53, 0.08);
      }
      h1 {
        margin: 0 0 10px;
        font-size: 22px;
      }
      p {
        margin: 0;
        color: #536068;
        line-height: 1.6;
      }
    </style>
  </head>
  <body>
    <main>
      <h1>${escapeHtml(title)}</h1>
      <p>${escapeHtml(message)}</p>
    </main>
  </body>
</html>`
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;')
}

function createEncryptedCachePlugin(cachePath: string): ICachePlugin {
  return {
    beforeCacheAccess: async (cacheContext: TokenCacheContext) => {
      const serializedCache = await readPersistedCache(cachePath)

      if (serializedCache) {
        cacheContext.tokenCache.deserialize(serializedCache)
      }
    },
    afterCacheAccess: async (cacheContext: TokenCacheContext) => {
      if (cacheContext.cacheHasChanged) {
        await writePersistedCache(cachePath, cacheContext.tokenCache.serialize())
      }
    }
  }
}

async function readPersistedCache(cachePath: string): Promise<string> {
  let rawCache: string

  try {
    rawCache = await readFile(cachePath, 'utf8')
  } catch {
    return ''
  }

  try {
    const storedCache = JSON.parse(rawCache) as StoredCache

    if (storedCache.encrypted) {
      if (!safeStorage.isEncryptionAvailable()) {
        return ''
      }

      return safeStorage.decryptString(Buffer.from(storedCache.data, 'base64'))
    }

    return storedCache.data
  } catch {
    return ''
  }
}

async function writePersistedCache(cachePath: string, serializedCache: string): Promise<void> {
  await mkdir(dirname(cachePath), { recursive: true })

  const storedCache: StoredCache = safeStorage.isEncryptionAvailable()
    ? {
        version: 1,
        encrypted: true,
        data: safeStorage.encryptString(serializedCache).toString('base64')
      }
    : {
        version: 1,
        encrypted: false,
        data: serializedCache
      }

  await writeFile(cachePath, JSON.stringify(storedCache), 'utf8')
}

async function readAuthState(): Promise<StoredAuthState> {
  try {
    const state = JSON.parse(await readFile(getAuthStatePath(), 'utf8')) as StoredAuthState

    if (state.version === 1) {
      return {
        version: 1,
        activeAccountId: typeof state.activeAccountId === 'string' ? state.activeAccountId : undefined
      }
    }
  } catch {
    // Missing or invalid auth state falls back to the first cached account.
  }

  return { version: 1 }
}

async function writeAuthState(state: StoredAuthState): Promise<void> {
  const authStatePath = getAuthStatePath()

  await mkdir(dirname(authStatePath), { recursive: true })
  await writeFile(authStatePath, JSON.stringify(state, null, 2), 'utf8')
}

async function clearAuthState(): Promise<void> {
  try {
    await unlink(getAuthStatePath())
  } catch {
    // Auth state may not exist on a fresh install.
  }
}
