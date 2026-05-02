export type PlatformName = 'aix' | 'darwin' | 'freebsd' | 'linux' | 'openbsd' | 'sunos' | 'win32'

export type PlatformInfo = {
  name: PlatformName
  arch: string
  release: string
  homeDirectory: string
  appDataPath: string
}

export type OneDriveLocation = {
  label: string
  path: string
  source: 'cloud-storage' | 'environment' | 'home' | 'user-profile'
  exists: boolean
}

export type AppEnvironment = {
  platform: PlatformInfo
  oneDriveLocations: OneDriveLocation[]
}

export type RevealPathResult = {
  ok: boolean
  message?: string
}

