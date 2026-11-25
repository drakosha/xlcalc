/** CFB sector allocation and the root's Workbook stream ([MS-CFB], sections 2.2–2.6). */
import { InputError } from '../input-error.js'
import { hasCompoundSignature } from '../xlsx/read/book.js'
import type { XlsxLimits } from '../xlsx/types.js'
import { ENTRY_BYTES, TOTAL_BYTES } from '../xlsx/zip/read.js'

const FREE = 0xffffffff
const END = 0xfffffffe
const FAT = 0xfffffffd
const DIFAT = 0xfffffffc
const MINI_SIZE = 64
const MINI_CUTOFF = 4096

export class CompoundLayoutError extends Error {}

function bad(reason: string): never {
  throw new CompoundLayoutError(`xls: invalid compound file: ${reason}`)
}

const viewOf = (bytes: Uint8Array) => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
const word = (view: DataView, at: number) => view.getUint32(at, true)

function chain(first: number, next: (id: number) => number, available: number, count?: number, stopAtCount = false): number[] {
  const ids: number[] = []
  const seen = new Set<number>()
  for (let id = first; id !== END; id = next(id)) {
    if (stopAtCount && ids.length === count) break
    if (id >= available || ids.length === count) bad(`sector chain points outside its bounds at ${id}`)
    if (seen.has(id)) bad(`sector chain repeats sector ${id}`)
    seen.add(id)
    ids.push(id)
  }
  if (count !== undefined && ids.length !== count) bad('sector chain ends before its declared size')
  return ids
}

interface DirectoryEntry {
  name: string
  type: number
  left: number
  right: number
  child: number
  start: number
  size: number
}

export interface Compound {
  workbook: Uint8Array
  stream(storage: string, name: string): Uint8Array | undefined
}

export function openCompound(bytes: Uint8Array, limits: XlsxLimits): Compound {
  if (!hasCompoundSignature(bytes)) throw new InputError('data', 'expected an Excel 97–2003 .xls compound file')
  if (bytes.length < 512) bad('truncated header')
  const header = viewOf(bytes)
  const version = header.getUint16(26, true)
  const shift = header.getUint16(30, true)
  if ((version !== 3 && version !== 4) || shift !== (version === 3 ? 9 : 12) ||
      header.getUint16(28, true) !== 0xfffe || header.getUint16(32, true) !== 6 ||
      word(header, 56) !== MINI_CUTOFF) bad('unsupported sector layout')
  const sectorSize = 2 ** shift
  const sectorCount = Math.ceil(bytes.length / sectorSize) - 1
  if (sectorCount < 1) bad('truncated sector')
  const sector = (id: number): Uint8Array => {
    if (id >= sectorCount) bad(`sector ${id} is outside the file`)
    const start = (id + 1) * sectorSize
    return bytes.subarray(start, start + sectorSize)
  }

  const fatCount = word(header, 44)
  const difatCount = word(header, 72)
  if (fatCount < 1 || fatCount + difatCount > sectorCount) bad('invalid FAT or DIFAT sector count')
  const fullSector = (id: number): DataView => {
    const data = sector(id)
    if (data.length !== sectorSize) bad(`allocation sector ${id} is truncated`)
    return viewOf(data)
  }
  const fat: DataView[] = []
  const reserved = new Set<number>()
  const addFat = (id: number) => {
    if (id === FREE) return
    if (fat.length === fatCount || reserved.has(id)) bad('invalid or repeated FAT sector')
    fat.push(fullSector(id))
    reserved.add(id)
  }
  for (let at = 76; at < 512; at += 4) addFat(word(header, at))
  const firstDifat = word(header, 68)
  const difat = chain(difatCount === 0 && firstDifat === FREE ? END : firstDifat, (id) => {
    if (reserved.has(id)) bad('DIFAT sector overlaps the FAT')
    reserved.add(id)
    const data = fullSector(id)
    for (let at = 0; at < sectorSize - 4; at += 4) addFat(word(data, at))
    return word(data, sectorSize - 4)
  }, sectorCount, difatCount)
  const perSector = sectorSize / 4
  if (fat.length !== fatCount || fatCount * perSector < sectorCount) bad('incomplete FAT')
  const next = (id: number) => word(fat[Math.floor(id / perSector)], (id % perSector) * 4)
  for (const id of difat) if (next(id) !== DIFAT) bad('DIFAT sector is not marked in the FAT')
  // Some writers link the FAT sectors themselves in the FAT. The DIFAT already reserves each of them.
  for (const id of reserved) {
    const at = next(id)
    if (at !== FAT && at !== DIFAT && at !== END && !reserved.has(at)) bad('allocation sector is not marked in the FAT')
  }

  function disjoint(ids: number[], used: Set<number>, kind: string): void {
    for (const id of ids) if (used.has(id)) bad(`overlapping ${kind} ${id}`)
  }

  function regular(first: number, count?: number): number[] {
    const ids = chain(first, next, sectorCount, count)
    disjoint(ids, reserved, 'sector')
    for (const id of ids) reserved.add(id)
    return ids
  }

  function read(ids: number[], blockSize: number, get: (id: number) => Uint8Array, size: number, used: Set<number>): Uint8Array {
    disjoint(ids, used, blockSize === MINI_SIZE ? 'mini-sector' : 'sector')
    const block = (id: number, length: number) => {
      const data = get(id)
      if (data.length < length) bad(`stream data in sector ${id} is truncated`)
      return data.subarray(0, length)
    }
    let result: Uint8Array
    if (ids.length === 1) result = block(ids[0], size)
    else {
      result = new Uint8Array(size)
      for (let i = 0; i < ids.length; i++) {
        const at = i * blockSize
        result.set(block(ids[i], Math.min(blockSize, size - at)), at)
      }
    }
    // An unreadable optional stream must not take sectors away from later reads.
    for (const id of ids) used.add(id)
    return result
  }

  const directory = regular(word(header, 48), version === 4 ? word(header, 40) : undefined)
  const entriesPerSector = sectorSize / 128
  function entry(id: number): DirectoryEntry {
    if (id >= directory.length * entriesPerSector) bad(`directory entry ${id} is outside the directory`)
    const data = fullSector(directory[Math.floor(id / entriesPerSector)])
    const at = (id % entriesPerSector) * 128
    const length = data.getUint16(at + 64, true)
    const type = data.getUint8(at + 66)
    let name = ''
    if (type === 1 || type === 2) {
      if (length < 2 || length > 64 || length % 2 || data.getUint16(at + length - 2, true) !== 0) bad('invalid directory name')
      for (let i = 0; i < length - 2; i += 2) name += String.fromCharCode(data.getUint16(at + i, true))
    }
    const size = word(data, at + 120) + (version === 4 ? word(data, at + 124) * 2 ** 32 : 0)
    return {
      name: name.toUpperCase(),
      type,
      left: word(data, at + 68),
      right: word(data, at + 72),
      child: word(data, at + 76),
      start: word(data, at + 116),
      size,
    }
  }

  const root = entry(0)
  if (root.type !== 5) bad('invalid root entry')
  const visited = new Set<number>([0])
  function children(storage: DirectoryEntry): Map<string, DirectoryEntry> {
    const pending = [storage.child]
    const entries = new Map<string, DirectoryEntry>()
    const found = new Set<number>()
    while (pending.length) {
      const id = pending.pop()!
      if (id === FREE) continue
      if (visited.has(id) || found.has(id)) bad(`directory tree repeats entry ${id}`)
      found.add(id)
      const item = entry(id)
      pending.push(item.left, item.right)
      if (item.type !== 1 && item.type !== 2) continue
      if (entries.has(item.name)) bad(`duplicate directory name ${item.name}`)
      entries.set(item.name, item)
    }
    for (const id of found) visited.add(id)
    return entries
  }

  const streams = children(root)
  const workbook = streams.get('WORKBOOK') ?? streams.get('BOOK')
  if (!workbook) {
    if (streams.has('ENCRYPTEDPACKAGE')) throw new InputError('data', 'password-protected Excel workbook')
    throw new InputError('data', 'the compound file has no Excel Workbook stream')
  }
  const entryBytes = limits.entryBytes ?? ENTRY_BYTES
  const totalBytes = limits.totalBytes ?? TOTAL_BYTES
  let spent = 0
  let miniFat: DataView[] | undefined
  let rootSectors: number[] | undefined
  const loaded = new Map<DirectoryEntry, Uint8Array>()
  function load(item: DirectoryEntry): Uint8Array {
    const held = loaded.get(item)
    if (held) return held
    const size = item.size
    const name = item.name === 'WORKBOOK' ? 'Workbook' : item.name
    if (size > entryBytes) throw new Error(`xls: ${name} stream exceeds ${entryBytes} bytes (limits.entryBytes)`)
    if (spent + size > totalBytes) throw new Error(`xls: ${name} stream exceeds the ${totalBytes} byte total (limits.totalBytes)`)
    let data: Uint8Array
    if (!size) data = new Uint8Array()
    else if (size >= MINI_CUTOFF) {
      // Excel stops at a stream's directory size even if its allocation chain continues.
      const ids = chain(item.start, next, sectorCount, Math.ceil(size / sectorSize), true)
      data = read(ids, sectorSize, sector, size, reserved)
    } else {
      // Mini-sector IDs address the root stream. The mini-FAT and root use regular sectors.
      const first = word(header, 60)
      const count = word(header, 64)
      const table = miniFat ??= regular(count === 0 && first === FREE ? END : first, count).map(fullSector)
      const roots = rootSectors ??= regular(root.start, Math.ceil(root.size / sectorSize))
      const nextMini = (id: number) => word(table[Math.floor(id / perSector)], (id % perSector) * 4)
      const available = Math.min(Math.floor(root.size / MINI_SIZE), table.length * perSector)
      const ids = chain(item.start, nextMini, available, Math.ceil(size / MINI_SIZE), true)
      data = read(ids, MINI_SIZE, (id) => {
        const at = id * MINI_SIZE
        const offset = at % sectorSize
        return sector(roots[Math.floor(at / sectorSize)]).subarray(offset, offset + MINI_SIZE)
      }, size, miniUsed)
    }
    spent += size
    loaded.set(item, data)
    return data
  }

  const miniUsed = new Set<number>()
  const folders = new Map<string, Map<string, DirectoryEntry>>()
  if (!workbook.size) bad('empty Workbook stream')
  return {
    workbook: load(workbook),
    stream(storage, name) {
      const itemName = name.toUpperCase()
      if (!storage) {
        const item = streams.get(itemName)
        return item?.type === 2 ? load(item) : undefined
      }
      const key = storage.toUpperCase()
      const folder = streams.get(key)
      if (folder?.type !== 1) return undefined
      let entries = folders.get(key)
      if (!entries) folders.set(key, (entries = children(folder)))
      const item = entries.get(itemName)
      return item?.type === 2 ? load(item) : undefined
    },
  }
}
