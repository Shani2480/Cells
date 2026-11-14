/**
 * AutoRecover, as Excel's: while a workbook has unsaved changes, a copy is kept every few minutes in the app's own
 * storage (IndexedDB). Saving or closing the workbook removes its copy, so copies found at start-up are from a session
 * that ended unexpectedly; Document Recovery offers them.
 */
export type RecoveredFile = {
  id: string
  /** The workbook's name as shown in the title ("Budget", "Book1"). */
  title: string
  fileName: string
  path?: string
  saved: number
  bytes: Uint8Array
}

export const AUTORECOVER_MS = 2 * 60 * 1000

const DB = "cells"
const STORE = "recovery"

function openDb(): Promise<IDBDatabase | null> {
  if (typeof indexedDB === "undefined") return Promise.resolve(null)
  return new Promise((resolve) => {
    try {
      const request = indexedDB.open(DB, 1)
      request.onupgradeneeded = () => request.result.createObjectStore(STORE, { keyPath: "id" })
      request.onsuccess = () => resolve(request.result)
      request.onerror = () => resolve(null)
    } catch {
      resolve(null)
    }
  })
}

async function withStore<T>(
  mode: IDBTransactionMode,
  run: (store: IDBObjectStore) => IDBRequest<T>,
): Promise<T | null> {
  const db = await openDb()
  if (!db) return null
  return new Promise((resolve) => {
    try {
      const request = run(db.transaction(STORE, mode).objectStore(STORE))
      request.onsuccess = () => resolve(request.result)
      request.onerror = () => resolve(null)
    } catch {
      resolve(null)
    } finally {
      db.close()
    }
  })
}

export async function keepRecovery(file: RecoveredFile): Promise<void> {
  await withStore("readwrite", (s) => s.put(file))
}

export async function dropRecovery(id: string): Promise<void> {
  await withStore("readwrite", (s) => s.delete(id))
}

export async function recoveredFiles(): Promise<RecoveredFile[]> {
  const all = (await withStore<RecoveredFile[]>("readonly", (s) => s.getAll() as IDBRequest<RecoveredFile[]>)) ?? []
  return all.sort((a, b) => b.saved - a.saved)
}
