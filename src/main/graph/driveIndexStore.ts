import { app } from 'electron'
import { Buffer } from 'node:buffer'
import { mkdir } from 'node:fs/promises'
import { join } from 'node:path'
import { DatabaseSync } from 'node:sqlite'
import type { CloudDriveItem, CloudDriveItemType } from '../../shared/types'

const DRIVE_INDEXES_DIR_NAME = 'drive-indexes'
const DRIVE_INDEX_DB_FILE_NAME = 'drive-index.sqlite'
const DRIVE_INDEX_SCHEMA_VERSION = 1
const SQLITE_BUSY_TIMEOUT_MS = 5000
const ROOT_TRIE_PARENT_ID = '__root__'

export type DriveIndex = {
  version: 1
  rootItemId?: string
  deltaLink?: string
  syncedAt?: string
  expandedFolderIds: Record<string, true>
  items: Record<string, CloudDriveItem>
}

export type DriveDeltaStagePayload = {
  itemId: string
  sequence: number
  payload: string
  isTombstone?: boolean
}

type DriveIndexDatabaseEntry = {
  db: DatabaseSync
  path: string
}

type DriveItemRow = {
  id: string
  name: string
  type: CloudDriveItemType
  size: number
  last_modified_date_time: string | null
  web_url: string | null
  parent_id: string | null
  child_count: number | null
  mime_type: string | null
  quick_xor_hash: string | null
  c_tag: string | null
  e_tag: string | null
}

type MetaRow = {
  key: string
  value: string
}

type ExpandedFolderRow = {
  folder_id: string
}

type DeltaStageRow = {
  item_id: string
  sequence: number
  payload: string
}

const driveIndexDatabases = new Map<string, DriveIndexDatabaseEntry>()

export function createEmptyDriveIndex(rootItemId?: string): DriveIndex {
  return {
    version: 1,
    rootItemId,
    expandedFolderIds: {},
    items: {}
  }
}

export async function readDriveIndexFromStore(accountId: string | null): Promise<DriveIndex | null> {
  if (!accountId) {
    return createEmptyDriveIndex()
  }

  const db = await getDriveIndexDatabase(accountId)
  const metaRows = db.prepare('SELECT key, value FROM drive_index_meta').all() as MetaRow[]
  const itemRows = db.prepare('SELECT * FROM drive_items').all() as DriveItemRow[]
  const expandedFolderRows = db.prepare('SELECT folder_id FROM expanded_folders').all() as ExpandedFolderRow[]

  if (metaRows.length === 0 && itemRows.length === 0 && expandedFolderRows.length === 0) {
    return null
  }

  const meta = new Map(metaRows.map((row) => [row.key, row.value]))
  const items: Record<string, CloudDriveItem> = {}
  const expandedFolderIds: Record<string, true> = {}

  for (const row of itemRows) {
    items[row.id] = rowToDriveItem(row)
  }

  for (const row of expandedFolderRows) {
    expandedFolderIds[row.folder_id] = true
  }

  return {
    version: 1,
    rootItemId: meta.get('root_item_id') || undefined,
    deltaLink: meta.get('delta_link') || undefined,
    syncedAt: meta.get('synced_at') || undefined,
    expandedFolderIds,
    items
  }
}

export async function writeDriveIndexToStore(index: DriveIndex, accountId: string | null): Promise<void> {
  if (!accountId) {
    throw new Error('Microsoft 계정 로그인이 필요합니다.')
  }

  const db = await getDriveIndexDatabase(accountId)

  withTransaction(db, () => {
    db.exec(`
      DELETE FROM drive_item_fts;
      DELETE FROM drive_path_trie;
      DELETE FROM drive_items;
      DELETE FROM expanded_folders;
      DELETE FROM drive_index_meta;
    `)

    const insertMeta = db.prepare('INSERT INTO drive_index_meta (key, value) VALUES (?, ?)')
    insertMeta.run('schema_version', String(DRIVE_INDEX_SCHEMA_VERSION))

    if (index.rootItemId) {
      insertMeta.run('root_item_id', index.rootItemId)
    }

    if (index.deltaLink) {
      insertMeta.run('delta_link', index.deltaLink)
    }

    if (index.syncedAt) {
      insertMeta.run('synced_at', index.syncedAt)
    }

    const insertItem = db.prepare(`
      INSERT INTO drive_items (
        id,
        name,
        normalized_name,
        type,
        size,
        last_modified_date_time,
        web_url,
        parent_id,
        child_count,
        mime_type,
        quick_xor_hash,
        c_tag,
        e_tag,
        metadata_text
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `)
    const insertTrieNode = db.prepare(`
      INSERT OR REPLACE INTO drive_path_trie (parent_id, normalized_name, item_id, item_type)
      VALUES (?, ?, ?, ?)
    `)
    const insertFts = db.prepare('INSERT INTO drive_item_fts (item_id, name, metadata) VALUES (?, ?, ?)')
    const insertExpandedFolder = db.prepare('INSERT OR REPLACE INTO expanded_folders (folder_id) VALUES (?)')

    for (const item of Object.values(index.items)) {
      const normalizedName = normalizeDriveItemNameForIndex(item.name)
      const metadataText = createDriveItemMetadataText(item)

      insertItem.run(
        item.id,
        item.name,
        normalizedName,
        item.type,
        Math.max(0, Math.floor(item.size)),
        item.lastModifiedDateTime ?? null,
        item.webUrl ?? null,
        item.parentId ?? null,
        item.childCount ?? null,
        item.mimeType ?? null,
        item.quickXorHash ?? null,
        item.cTag ?? null,
        item.eTag ?? null,
        metadataText
      )
      insertFts.run(item.id, item.name, metadataText)

      if (item.parentId) {
        insertTrieNode.run(item.parentId, normalizedName, item.id, item.type)
      } else if (index.rootItemId && item.id !== index.rootItemId) {
        insertTrieNode.run(ROOT_TRIE_PARENT_ID, normalizedName, item.id, item.type)
      }
    }

    for (const folderId of Object.keys(index.expandedFolderIds)) {
      insertExpandedFolder.run(folderId)
    }
  })
}

export async function listDriveIndexChildren(accountId: string | null, parentId: string): Promise<CloudDriveItem[]> {
  if (!accountId) {
    return []
  }

  const db = await getDriveIndexDatabase(accountId)
  const rows = db
    .prepare('SELECT * FROM drive_items WHERE parent_id = ? ORDER BY type, normalized_name')
    .all(parentId) as DriveItemRow[]

  return rows.map(rowToDriveItem)
}

export async function findDriveIndexChildByName(
  accountId: string | null,
  parentId: string,
  name: string
): Promise<CloudDriveItem | null> {
  if (!accountId) {
    return null
  }

  const db = await getDriveIndexDatabase(accountId)
  const row = db
    .prepare(`
      SELECT drive_items.*
      FROM drive_path_trie
      JOIN drive_items ON drive_items.id = drive_path_trie.item_id
      WHERE drive_path_trie.parent_id = ? AND drive_path_trie.normalized_name = ?
      LIMIT 1
    `)
    .get(parentId, normalizeDriveItemNameForIndex(name)) as DriveItemRow | undefined

  return row ? rowToDriveItem(row) : null
}

export async function getDriveIndexItemById(accountId: string | null, itemId: string): Promise<CloudDriveItem | null> {
  if (!accountId) {
    return null
  }

  const db = await getDriveIndexDatabase(accountId)
  const row = db.prepare('SELECT * FROM drive_items WHERE id = ? LIMIT 1').get(itemId) as DriveItemRow | undefined

  return row ? rowToDriveItem(row) : null
}

export async function searchDriveIndexItems(accountId: string | null, query: string, limit: number): Promise<CloudDriveItem[]> {
  if (!accountId) {
    return []
  }

  const trimmedQuery = query.normalize('NFC').trim()

  if (!trimmedQuery) {
    return []
  }

  const db = await getDriveIndexDatabase(accountId)
  const normalizedLimit = Math.min(500, Math.max(1, Math.floor(limit)))

  if ([...trimmedQuery].length < 3) {
    const pattern = `%${escapeLikePattern(normalizeDriveItemNameForIndex(trimmedQuery))}%`
    const rows = db
      .prepare(`
        SELECT *
        FROM drive_items
        WHERE normalized_name LIKE ? ESCAPE '\\'
          OR lower(metadata_text) LIKE ? ESCAPE '\\'
        ORDER BY type, normalized_name
        LIMIT ?
      `)
      .all(pattern, pattern, normalizedLimit) as DriveItemRow[]

    return rows.map(rowToDriveItem)
  }

  const rows = db
    .prepare(`
      SELECT drive_items.*
      FROM drive_item_fts
      JOIN drive_items ON drive_items.id = drive_item_fts.item_id
      WHERE drive_item_fts MATCH ?
      ORDER BY rank
      LIMIT ?
    `)
    .all(quoteFts5Match(trimmedQuery), normalizedLimit) as DriveItemRow[]

  return rows.map(rowToDriveItem)
}

export async function resetDriveDeltaStaging(accountId: string | null): Promise<void> {
  if (!accountId) {
    return
  }

  const db = await getDriveIndexDatabase(accountId)
  db.exec('DELETE FROM drive_delta_staging')
}

export async function appendDriveDeltaStaging(
  accountId: string | null,
  items: DriveDeltaStagePayload[]
): Promise<void> {
  if (!accountId || items.length === 0) {
    return
  }

  const db = await getDriveIndexDatabase(accountId)
  const insert = db.prepare(`
    INSERT INTO drive_delta_staging (item_id, sequence, payload, is_tombstone)
    VALUES (?, ?, ?, ?)
  `)

  withTransaction(db, () => {
    for (const item of items) {
      insert.run(item.itemId, item.sequence, item.payload, item.isTombstone ? 1 : 0)
    }
  })
}

export async function readLastOccurrenceDriveDeltaStaging(accountId: string | null): Promise<DriveDeltaStagePayload[]> {
  if (!accountId) {
    return []
  }

  const db = await getDriveIndexDatabase(accountId)
  const rows = db
    .prepare(`
      SELECT item_id, sequence, payload
      FROM (
        SELECT
          item_id,
          sequence,
          payload,
          row_number() OVER (PARTITION BY item_id ORDER BY sequence DESC) AS occurrence_rank
        FROM drive_delta_staging
      )
      WHERE occurrence_rank = 1
      ORDER BY sequence
    `)
    .all() as DeltaStageRow[]

  return rows.map((row) => ({
    itemId: row.item_id,
    sequence: row.sequence,
    payload: row.payload
  }))
}

export async function clearDriveDeltaStaging(accountId: string | null): Promise<void> {
  await resetDriveDeltaStaging(accountId)
}

export function closeDriveIndexStores(): void {
  for (const entry of driveIndexDatabases.values()) {
    entry.db.close()
  }

  driveIndexDatabases.clear()
}

export function getDriveIndexStoreDirectory(): string {
  return join(app.getPath('userData'), DRIVE_INDEXES_DIR_NAME)
}

export function getLegacyAccountDriveIndexPath(accountId: string): string {
  return join(getDriveIndexAccountDirectory(accountId), 'drive-index.json')
}

async function getDriveIndexDatabase(accountId: string): Promise<DatabaseSync> {
  const databasePath = getDriveIndexDatabasePath(accountId)
  const existingEntry = driveIndexDatabases.get(accountId)

  if (existingEntry?.path === databasePath) {
    return existingEntry.db
  }

  await mkdir(getDriveIndexAccountDirectory(accountId), { recursive: true })
  const db = new DatabaseSync(databasePath, { timeout: SQLITE_BUSY_TIMEOUT_MS })
  configureDriveIndexDatabase(db)
  driveIndexDatabases.set(accountId, { db, path: databasePath })

  return db
}

function configureDriveIndexDatabase(db: DatabaseSync): void {
  db.exec(`
    PRAGMA journal_mode = WAL;
    PRAGMA synchronous = NORMAL;
    PRAGMA foreign_keys = ON;
    PRAGMA temp_store = MEMORY;

    CREATE TABLE IF NOT EXISTS drive_index_meta (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS drive_items (
      id TEXT PRIMARY KEY,
      name TEXT NOT NULL,
      normalized_name TEXT NOT NULL,
      type TEXT NOT NULL CHECK (type IN ('file', 'folder', 'package')),
      size INTEGER NOT NULL DEFAULT 0,
      last_modified_date_time TEXT,
      web_url TEXT,
      parent_id TEXT,
      child_count INTEGER,
      mime_type TEXT,
      quick_xor_hash TEXT,
      c_tag TEXT,
      e_tag TEXT,
      metadata_text TEXT NOT NULL DEFAULT ''
    );

    CREATE INDEX IF NOT EXISTS idx_drive_items_parent_name
      ON drive_items(parent_id, normalized_name);
    CREATE INDEX IF NOT EXISTS idx_drive_items_parent
      ON drive_items(parent_id);
    CREATE INDEX IF NOT EXISTS idx_drive_items_type
      ON drive_items(type);

    CREATE TABLE IF NOT EXISTS drive_path_trie (
      parent_id TEXT NOT NULL,
      normalized_name TEXT NOT NULL,
      item_id TEXT NOT NULL,
      item_type TEXT NOT NULL CHECK (item_type IN ('file', 'folder', 'package')),
      PRIMARY KEY (parent_id, normalized_name),
      FOREIGN KEY (item_id) REFERENCES drive_items(id) ON DELETE CASCADE
    );

    CREATE INDEX IF NOT EXISTS idx_drive_path_trie_item
      ON drive_path_trie(item_id);

    CREATE TABLE IF NOT EXISTS expanded_folders (
      folder_id TEXT PRIMARY KEY
    );

    CREATE TABLE IF NOT EXISTS drive_delta_staging (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      item_id TEXT NOT NULL,
      sequence INTEGER NOT NULL,
      payload TEXT NOT NULL,
      is_tombstone INTEGER NOT NULL DEFAULT 0
    );

    CREATE INDEX IF NOT EXISTS idx_drive_delta_staging_item_sequence
      ON drive_delta_staging(item_id, sequence DESC);

    CREATE TABLE IF NOT EXISTS local_merkle_nodes (
      cache_scope TEXT NOT NULL,
      node_path TEXT NOT NULL,
      parent_path TEXT,
      hash_algorithm TEXT NOT NULL DEFAULT 'BLAKE3',
      hash_hex TEXT NOT NULL,
      byte_size INTEGER NOT NULL DEFAULT 0,
      modified_at TEXT,
      updated_at TEXT NOT NULL,
      PRIMARY KEY (cache_scope, node_path)
    );

    CREATE INDEX IF NOT EXISTS idx_local_merkle_nodes_parent
      ON local_merkle_nodes(cache_scope, parent_path);

    CREATE VIRTUAL TABLE IF NOT EXISTS drive_item_fts
      USING fts5(item_id UNINDEXED, name, metadata, tokenize='trigram');
  `)
}

function rowToDriveItem(row: DriveItemRow): CloudDriveItem {
  return {
    id: row.id,
    name: row.name,
    type: row.type,
    size: row.size,
    lastModifiedDateTime: row.last_modified_date_time ?? undefined,
    webUrl: row.web_url ?? undefined,
    parentId: row.parent_id ?? undefined,
    childCount: row.child_count ?? undefined,
    mimeType: row.mime_type ?? undefined,
    quickXorHash: row.quick_xor_hash ?? undefined,
    cTag: row.c_tag ?? undefined,
    eTag: row.e_tag ?? undefined
  }
}

function createDriveItemMetadataText(item: CloudDriveItem): string {
  return [
    item.type,
    item.mimeType,
    item.webUrl,
    item.lastModifiedDateTime,
    item.quickXorHash,
    item.cTag,
    item.eTag
  ]
    .filter((value): value is string => Boolean(value))
    .join(' ')
}

function normalizeDriveItemNameForIndex(name: string): string {
  return name.normalize('NFC').trim().toLocaleLowerCase('ko-KR')
}

function quoteFts5Match(value: string): string {
  return `"${value.replace(/"/g, '""')}"`
}

function escapeLikePattern(value: string): string {
  return value.replace(/[\\%_]/g, (character) => `\\${character}`)
}

function withTransaction(db: DatabaseSync, callback: () => void): void {
  db.exec('BEGIN IMMEDIATE')

  try {
    callback()
    db.exec('COMMIT')
  } catch (error) {
    db.exec('ROLLBACK')
    throw error
  }
}

function getDriveIndexDatabasePath(accountId: string): string {
  return join(getDriveIndexAccountDirectory(accountId), DRIVE_INDEX_DB_FILE_NAME)
}

function getDriveIndexAccountDirectory(accountId: string): string {
  return join(getDriveIndexStoreDirectory(), encodeAccountIdForPath(accountId))
}

function encodeAccountIdForPath(accountId: string): string {
  return Buffer.from(accountId, 'utf8').toString('base64url')
}
