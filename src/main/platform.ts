import { app, shell } from 'electron'
import { existsSync } from 'node:fs'
import { readdir, stat } from 'node:fs/promises'
import { arch, homedir, platform, release } from 'node:os'
import { isAbsolute, join, resolve } from 'node:path'
import type { AppEnvironment, OneDriveLocation, PlatformName, RevealPathResult } from '../shared/types'

type CandidateMap = Map<string, Omit<OneDriveLocation, 'exists'>>

const ONEDRIVE_ENV_KEYS = ['OneDrive', 'OneDriveConsumer', 'OneDriveCommercial'] as const

export async function getAppEnvironment(): Promise<AppEnvironment> {
  const homeDirectory = homedir()

  return {
    platform: {
      name: platform() as PlatformName,
      arch: arch(),
      release: release(),
      homeDirectory,
      appDataPath: app.getPath('userData')
    },
    oneDriveLocations: await findOneDriveLocations(homeDirectory)
  }
}

export async function revealPathInFileManager(targetPath: unknown): Promise<RevealPathResult> {
  if (typeof targetPath !== 'string' || targetPath.trim().length === 0) {
    return { ok: false, message: 'Invalid path.' }
  }

  if (!isAbsolute(targetPath)) {
    return { ok: false, message: 'Only absolute paths can be opened.' }
  }

  const normalizedPath = resolve(targetPath)

  if (!(await directoryExists(normalizedPath))) {
    return { ok: false, message: 'Path does not exist.' }
  }

  shell.showItemInFolder(normalizedPath)
  return { ok: true }
}

async function findOneDriveLocations(homeDirectory: string): Promise<OneDriveLocation[]> {
  const candidates: CandidateMap = new Map()

  addEnvironmentCandidates(candidates)
  addCandidate(candidates, join(homeDirectory, 'OneDrive'), 'OneDrive', 'home')
  await addPrefixedDirectories(candidates, homeDirectory, 'OneDrive', 'home')

  if (process.platform === 'darwin') {
    const cloudStoragePath = join(homeDirectory, 'Library', 'CloudStorage')

    addCandidate(candidates, join(cloudStoragePath, 'OneDrive-Personal'), 'OneDrive Personal', 'cloud-storage')
    await addPrefixedDirectories(candidates, cloudStoragePath, 'OneDrive', 'cloud-storage')
  }

  if (process.platform === 'win32') {
    const userProfile = process.env.USERPROFILE

    if (userProfile) {
      addCandidate(candidates, join(userProfile, 'OneDrive'), 'OneDrive', 'user-profile')
      await addPrefixedDirectories(candidates, userProfile, 'OneDrive', 'user-profile')
    }
  }

  const locations = await Promise.all(
    [...candidates.values()].map(async (candidate) => ({
      ...candidate,
      exists: await directoryExists(candidate.path)
    }))
  )

  return locations.sort((left, right) => {
    if (left.exists !== right.exists) {
      return left.exists ? -1 : 1
    }

    return left.label.localeCompare(right.label)
  })
}

function addEnvironmentCandidates(candidates: CandidateMap): void {
  for (const key of ONEDRIVE_ENV_KEYS) {
    const candidatePath = process.env[key]

    if (candidatePath) {
      addCandidate(candidates, candidatePath, key, 'environment')
    }
  }
}

async function addPrefixedDirectories(
  candidates: CandidateMap,
  rootPath: string,
  prefix: string,
  source: OneDriveLocation['source']
): Promise<void> {
  if (!existsSync(rootPath)) {
    return
  }

  let entries: string[]

  try {
    entries = await readdir(rootPath)
  } catch {
    return
  }

  for (const entry of entries) {
    if (entry === prefix || entry.startsWith(`${prefix} - `) || entry.startsWith(`${prefix}-`)) {
      addCandidate(candidates, join(rootPath, entry), entry, source)
    }
  }
}

function addCandidate(
  candidates: CandidateMap,
  candidatePath: string,
  label: string,
  source: OneDriveLocation['source']
): void {
  const normalizedPath = resolve(candidatePath)
  const key = process.platform === 'win32' ? normalizedPath.toLowerCase() : normalizedPath

  if (!candidates.has(key)) {
    candidates.set(key, {
      label,
      path: normalizedPath,
      source
    })
  }
}

async function directoryExists(targetPath: string): Promise<boolean> {
  try {
    const stats = await stat(targetPath)
    return stats.isDirectory()
  } catch {
    return false
  }
}

