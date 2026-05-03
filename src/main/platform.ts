import { app } from 'electron'
import { arch, homedir, platform, release } from 'node:os'
import type { AppEnvironment, PlatformName } from '../shared/types'

export function getAppEnvironment(): AppEnvironment {
  const homeDirectory = homedir()

  return {
    platform: {
      name: platform() as PlatformName,
      arch: arch(),
      release: release(),
      homeDirectory,
      appDataPath: app.getPath('userData')
    }
  }
}
