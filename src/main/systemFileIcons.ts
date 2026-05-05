import { app } from 'electron'
import { mkdir, open } from 'node:fs/promises'
import { join } from 'node:path'
import type { DriveFileIconRequest, DriveFileIconResult, DriveFileIconSize } from '../shared/types'

const SYSTEM_FILE_ICON_SAMPLE_DIR_NAME = 'system-file-icons'
const systemFileIconCache = new Map<string, Promise<DriveFileIconResult>>()

const mimeExtensionHints = new Map<string, string>([
  ['application/json', 'json'],
  ['application/msword', 'doc'],
  ['application/pdf', 'pdf'],
  ['application/rtf', 'rtf'],
  ['application/vnd.ms-excel', 'xls'],
  ['application/vnd.ms-powerpoint', 'ppt'],
  ['application/vnd.openxmlformats-officedocument.presentationml.presentation', 'pptx'],
  ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'xlsx'],
  ['application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'docx'],
  ['application/zip', 'zip'],
  ['audio/mpeg', 'mp3'],
  ['image/jpeg', 'jpg'],
  ['image/png', 'png'],
  ['text/csv', 'csv'],
  ['text/html', 'html'],
  ['text/markdown', 'md'],
  ['text/plain', 'txt'],
  ['video/mp4', 'mp4']
])

type SystemFileIconLookup = {
  cacheKey: string
  extension: string
  size: DriveFileIconSize
}

export function getSystemFileIcon(request: DriveFileIconRequest): Promise<DriveFileIconResult> {
  const lookup = normalizeSystemFileIconLookup(request)
  let promise = systemFileIconCache.get(lookup.cacheKey)

  if (!promise) {
    promise = readSystemFileIcon(lookup)
    systemFileIconCache.set(lookup.cacheKey, promise)
  }

  return promise
}

async function readSystemFileIcon(lookup: SystemFileIconLookup): Promise<DriveFileIconResult> {
  try {
    const samplePath = await ensureSystemFileIconSamplePath(lookup.extension)
    const icon = await app.getFileIcon(samplePath, { size: lookup.size })

    return {
      cacheKey: lookup.cacheKey,
      url: icon.isEmpty() ? undefined : icon.toDataURL()
    }
  } catch {
    return {
      cacheKey: lookup.cacheKey
    }
  }
}

async function ensureSystemFileIconSamplePath(extension: string): Promise<string> {
  const directoryPath = join(app.getPath('userData'), SYSTEM_FILE_ICON_SAMPLE_DIR_NAME)
  const samplePath = join(directoryPath, extension ? `sample.${extension}` : 'sample')

  await mkdir(directoryPath, { recursive: true })

  const handle = await open(samplePath, 'a')
  await handle.close()

  return samplePath
}

function normalizeSystemFileIconLookup(request: DriveFileIconRequest): SystemFileIconLookup {
  const size = normalizeSystemFileIconSize(request.size)
  const extension = request.type === 'file' ? getFileIconExtension(request) : ''
  const cacheKey = [request.type, extension || 'generic', size].join(':')

  return {
    cacheKey,
    extension,
    size
  }
}

function normalizeSystemFileIconSize(size: DriveFileIconSize | undefined): DriveFileIconSize {
  return size === 'small' || size === 'large' ? size : 'normal'
}

function getFileIconExtension(request: DriveFileIconRequest): string {
  const extension = normalizeFileIconExtension(getFileExtension(request.name))

  if (extension) {
    return extension
  }

  const mimeExtension = request.mimeType ? mimeExtensionHints.get(request.mimeType.toLocaleLowerCase('en-US')) : undefined

  return normalizeFileIconExtension(mimeExtension ?? '')
}

function getFileExtension(name: string): string {
  const extensionIndex = name.lastIndexOf('.')

  if (extensionIndex < 0 || extensionIndex === name.length - 1) {
    return ''
  }

  return name.slice(extensionIndex + 1)
}

function normalizeFileIconExtension(extension: string): string {
  const normalizedExtension = extension.trim().toLocaleLowerCase('en-US')

  return /^[a-z0-9][a-z0-9_+-]{0,31}$/.test(normalizedExtension) ? normalizedExtension : ''
}
