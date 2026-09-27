<script setup lang="ts">
import { ref, onMounted, onUnmounted } from 'vue'
import JSZip from 'jszip'
import * as XLSX from 'xlsx'

const fileInput = ref<HTMLInputElement | null>(null)
const fileName = ref('')
const headers = ref<string[]>([])
const tableData = ref<any[]>([])
const editingCell = ref<{ row: number; col: string } | null>(null)
const editValue = ref('')
const error = ref('')
const loading = ref(false)
let originalZip: JSZip | null = null
let originalFileType = '' // 'xlsx' | 'xls'

interface CellImage {
  cell: string
  imageUrl: string
  imgId: number
}
interface SheetCellImages { [cellRef: string]: CellImage[] }
const cellImages = ref<SheetCellImages>({})

interface HistoryEntry { data: any[]; headers: string[] }
const history = ref<HistoryEntry[]>([])
const historyIndex = ref(-1)
const MAX_HISTORY = 50

interface RowData { [key: string]: string }

const pushHistory = () => {
  const snapshot: HistoryEntry = {
    data: JSON.parse(JSON.stringify(tableData.value)),
    headers: [...headers.value]
  }
  history.value = history.value.slice(0, historyIndex.value + 1)
  history.value.push(snapshot)
  if (history.value.length > MAX_HISTORY) history.value.shift()
  historyIndex.value = history.value.length - 1
}

const undo = () => {
  if (historyIndex.value <= 0) return
  historyIndex.value--
  const prev = history.value[historyIndex.value]
  tableData.value = JSON.parse(JSON.stringify(prev.data))
  headers.value = [...prev.headers]
}

const redo = () => {
  if (historyIndex.value >= history.value.length - 1) return
  historyIndex.value++
  const next = history.value[historyIndex.value]
  tableData.value = JSON.parse(JSON.stringify(next.data))
  headers.value = [...next.headers]
}

const handleKeydown = (e: KeyboardEvent) => {
  if (editingCell.value) return
  const mod = e.ctrlKey || e.metaKey
  if (mod && e.key === 'z' && !e.shiftKey) { e.preventDefault(); undo() }
  if ((mod && e.key === 'y') || (mod && e.shiftKey && e.key === 'Z')) { e.preventDefault(); redo() }
}

onMounted(() => window.addEventListener('keydown', handleKeydown))
onUnmounted(() => window.removeEventListener('keydown', handleKeydown))

const triggerFileInput = () => fileInput.value?.click()

// 列号转字母（0-based → "A", "B", "AA" 等）
const colToLetter = (col: number): string => {
  let result = ''
  let c = col
  while (c >= 0) {
    result = String.fromCharCode(65 + (c % 26)) + result
    c = Math.floor(c / 26) - 1
  }
  return result
}

const parseDrawingRels = async (relText: string): Promise<Map<string, string>> => {
  const map = new Map<string, string>()
  // 用正则解析，避免 DOMParser 命名空间问题
  const relMatches = relText.match(/<Relationship\s+Id="([^"]*)"[^>]*Target="([^"]*)"/g) || []
  relMatches.forEach(m => {
    const idMatch = m.match(/Id="([^"]*)"/)
    const targetMatch = m.match(/Target="([^"]*)"/)
    if (idMatch && targetMatch) map.set(idMatch[1], targetMatch[1])
  })
  return map
}

const parseDrawings = async (xml: string, relMap: Map<string, string>): Promise<{ cell: string; mediaId: number; rId: string }[]> => {
  const images: { cell: string; mediaId: number; rId: string }[] = []
  // 用正则解析 drawings XML，避免命名空间问题
  const anchorMatches = xml.match(/<\w*:(?:twoCellAnchor|oneCellAnchor)[^>]*>([\s\S]*?)<\/\w*:(?:twoCellAnchor|oneCellAnchor)>/g) || []
  let mediaId = 1
  for (const anchorXml of anchorMatches) {
    const pos = parseAnchorRegex(anchorXml)
    if (!pos) continue
    const blipMatch = anchorXml.match(/r:embed="([^"]*)"/) || anchorXml.match(/embed="([^"]*)"/) || anchorXml.match(/r:link="([^"]*)"/) || anchorXml.match(/link="([^"]*)"/)
    if (!blipMatch) continue
    const rId = blipMatch[1]
    const mediaFile = relMap.get(rId)
    if (!mediaFile) continue
    const cellRef = `${colToLetter(pos.startCol)}${pos.startRow + 1}`
    images.push({ cell: cellRef, mediaId: mediaId++, rId })
  }
  return images
}

const parseAnchorRegex = (anchorXml: string): { startCol: number; startRow: number } | null => {
  const fromMatch = anchorXml.match(/<\w*:from>([\s\S]*?)<\/\w*:from>/) || anchorXml.match(/<from>([\s\S]*?)<\/from>/)
  if (!fromMatch) return null
  const fromContent = fromMatch[1]
  const colMatch = fromContent.match(/<\w*:col>(-?\d+)<\/\w*:col>/) || fromContent.match(/<col>(-?\d+)<\/col>/)
  const rowMatch = fromContent.match(/<\w*:row>(-?\d+)<\/\w*:row>/) || fromContent.match(/<row>(-?\d+)<\/row>/)
  if (!colMatch || !rowMatch) return null
  return { startCol: parseInt(colMatch[1]), startRow: parseInt(rowMatch[1]) }
}

const rIdToImageUrl = async (zip: JSZip, rId: string, relMap: Map<string, string>): Promise<string | null> => {
  const mediaFile = relMap.get(rId)
  if (!mediaFile) return null
  const resolvedPath = mediaFile.startsWith('../') ? 'xl/' + mediaFile.slice(3) : mediaFile
  const media = zip.file(resolvedPath) || zip.file(mediaFile)
  if (!media) return null
  const buffer = await media.async('uint8array')
  const mimeMatch = resolvedPath.match(/\.([^.\/]+)$/)
  const mime = mimeMatch ? `image/${mimeMatch[1]}` : 'image/png'
  const b64 = btoa(String.fromCharCode(...Array.from(new Uint8Array(buffer))))
  return `data:${mime};base64,${b64}`
}

const imgStyle = (imgId: number) => {
  const sizes: Record<number, string> = { 1: '40px', 2: '50px', 3: '60px', 4: '45px' }
  const size = sizes[imgId] || '40px'
  return { width: size, height: size, objectFit: 'cover' } as const
}

const loadImages = async (zip: JSZip) => {
  cellImages.value = {}

  // 方式1：传统 drawings 方式
  const sheetRels = zip.file('xl/worksheets/_rels/sheet1.xml.rels')
  if (sheetRels) {
    const sheetRelsText = await sheetRels.async('string')
    const sheetRelMap = new Map<string, string>()
    // 用正则解析，避免 DOMParser 命名空间问题
    const relMatches = sheetRelsText.match(/<Relationship\s+Id="([^"]*)"[^>]*Target="([^"]*)"/g) || []
    relMatches.forEach(m => {
      const idMatch = m.match(/Id="([^"]*)"/)
      const targetMatch = m.match(/Target="([^"]*)"/)
      if (idMatch && targetMatch) sheetRelMap.set(idMatch[1], targetMatch[1])
    })
    const drawingRel = sheetRelMap.get('rId1')
    if (drawingRel) {
      const drawingPath = drawingRel.replace('../', '')
      const drawingDir = drawingPath.substring(0, drawingPath.lastIndexOf('/'))
      const drawingFilename = drawingPath.substring(drawingPath.lastIndexOf('/') + 1)
      const drawingRelsFile = zip.file(`xl/${drawingDir}/_rels/${drawingFilename}.rels`)
      if (drawingRelsFile) {
        const drawingRelsText = await drawingRelsFile.async('string')
        const relMap = await parseDrawingRels(drawingRelsText)
        const drawingFile = zip.file(`xl/${drawingPath}`)
        if (drawingFile) {
          const drawingXml = await drawingFile.async('string')
          const drawingImages = await parseDrawings(drawingXml, relMap)
          for (const img of drawingImages) {
            const url = await rIdToImageUrl(zip, img.rId, relMap)
            if (!url) continue
            if (!cellImages.value[img.cell]) cellImages.value[img.cell] = []
            cellImages.value[img.cell].push({ cell: img.cell, imageUrl: url, imgId: img.mediaId })
          }
          return
        }
      }
    }
  }

  // 方式2：Rich Data 方式（Excel 2022+ 嵌入图片）
  await loadRichDataImages(zip)
}

const loadRichDataImages = async (zip: JSZip) => {
  const rvrRelsFile = zip.file('xl/richData/_rels/richValueRel.xml.rels')
  if (!rvrRelsFile) return
  const rvrRelsText = await rvrRelsFile.async('string')
  const rvrRelMap = new Map<string, string>()
  const relMatches = rvrRelsText.match(/<Relationship\s+Id="([^"]*)"[^>]*Target="([^"]*)"/g) || []
  relMatches.forEach(m => {
    const idMatch = m.match(/Id="([^"]*)"/)
    const targetMatch = m.match(/Target="([^"]*)"/)
    if (idMatch && targetMatch) rvrRelMap.set(idMatch[1], targetMatch[1])
  })

  const sheetFile = zip.file('xl/worksheets/sheet1.xml')
  if (!sheetFile) return
  const sheetText = await sheetFile.async('string')
  const cellMatches = sheetText.match(/<c\s+r\s*="([^"]*)"[^>]*vm\s*=\s*"([^"]*)"/g) || []
  let imgId = 1
  for (const match of cellMatches) {
    const cellRefMatch = match.match(/r="([^"]*)"/)
    const vmMatch = match.match(/vm="([^"]*)"/)
    if (!cellRefMatch || !vmMatch) continue
    const cellRef = cellRefMatch[1]
    const vm = vmMatch[1]
    if (vm === '0') continue
    const relId = `rId${parseInt(vm)}`
    const mediaPath = rvrRelMap.get(relId)
    if (!mediaPath) continue
    const resolvedPath = mediaPath.startsWith('../') ? 'xl/' + mediaPath.slice(3) : mediaPath
    const mediaFile = zip.file(resolvedPath) || zip.file(mediaPath)
    if (!mediaFile) continue
    const buffer = await mediaFile.async('uint8array')
    const mimeMatch = resolvedPath.match(/\.([^.\/]+)$/)
    const mime = mimeMatch ? `image/${mimeMatch[1]}` : 'image/png'
    const b64 = btoa(String.fromCharCode(...Array.from(new Uint8Array(buffer))))
    const url = `data:${mime};base64,${b64}`
    if (!cellImages.value[cellRef]) cellImages.value[cellRef] = []
    cellImages.value[cellRef].push({ cell: cellRef, imageUrl: url, imgId: imgId++ })
  }
}

// 从 xlsx zip 解析数据（正则版本，避免命名空间问题）
const parseXlsxFromZip = async (zip: JSZip): Promise<{ data: RowData[]; headers: string[] }> => {
  const sheetFile = zip.file('xl/worksheets/sheet1.xml')
  if (!sheetFile) throw new Error('未找到工作表')
  const sheetText = await sheetFile.async('string')

  const sstFile = zip.file('xl/sharedStrings.xml')
  const sharedStrings: string[] = []
  if (sstFile) {
    const sstText = await sstFile.async('string')
    // 用正则提取共享字符串
    const siMatches = sstText.match(/<si>([\s\S]*?)<\/si>/g) || []
    for (const si of siMatches) {
      const tMatches = si.match(/<t[^>]*>([^<]*)<\/t>/g) || []
      let str = ''
      for (const tMatch of tMatches) {
        const textMatch = tMatch.match(/>([^<]*)</)
        if (textMatch) str += textMatch[1]
      }
      sharedStrings.push(str)
    }
  }

  // 用正则提取行和单元格
  const rowMatches = sheetText.match(/<row[^>]*>([\s\S]*?)<\/row>/g) || []
  const data: RowData[] = []
  for (const rowXml of rowMatches) {
    const rowData: RowData = {}
    const cellMatches = rowXml.match(/<c\s+r\s*="([^"]*)"[^>]*>([\s\S]*?)<\/c>/g) || []
    for (const cellMatch of cellMatches) {
      const cellRef = cellMatch.match(/r="([^"]*)"/)?.[1] || ''
      const col = cellRef.replace(/\d+/, '')
      const typeMatch = cellMatch.match(/t="([^"]*)"/)
      const type = typeMatch ? typeMatch[1] : ''
      const vMatch = cellMatch.match(/<v[^>]*>([^<]*)<\/v>/)
      const value = (type === 's' && vMatch) ? (sharedStrings[parseInt(vMatch[1])] ?? '') : (vMatch?.[1] ?? '')
      rowData[col] = value as string
    }
    if (Object.keys(rowData).length > 0) data.push(rowData)
  }
  const headers = data.length > 0
    ? Object.keys(data[0]).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
    : []
  return { data, headers }
}

// 从 xls 二进制解析数据（使用 xlsx 库）
const parseXls = async (arrayBuffer: ArrayBuffer): Promise<{ data: RowData[]; headers: string[]; images: Map<string, CellImage[]> }> => {
  const workbook = XLSX.read(arrayBuffer, { type: 'array' })
  const sheetName = workbook.SheetNames[0]
  const sheet = workbook.Sheets[sheetName]
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) as string[][]

  if (json.length === 0) return { data: [], headers: [], images: new Map() }
  const headers = json[0] as string[]
  const data: RowData[] = json.slice(1).filter(row => row.some(v => v !== '')).map(row => {
    const rowData: RowData = {}
    headers.forEach((h, i) => { rowData[h] = String(row[i] ?? '') })
    return rowData
  })

  // 使用 OLE 解析提取图片
  const result = new Map<string, CellImage[]>()
  const data8 = new Uint8Array(arrayBuffer)
  try {
    const { tree, streams } = parseOleDocument(data8)
    let imgIdx = 0
    for (const entry of tree) {
      const streamData = streams.get(entry.parent + '/' + entry.name)
      if (!streamData || streamData.length < 100) continue
      for (const { sig, mime } of IMAGE_SIGS) {
        let pos = 0
        while (true) {
          pos = findSubarray(streamData, sig, pos)
          if (pos === -1) break
          const extracted = extractImage(sig, mime, streamData, pos)
          if (!extracted || extracted.length < 50) { pos++; continue }
          const url = toDataURL(extracted, mime)
          imgIdx++
          const cellRef = `${String.fromCharCode(65 + (imgIdx % 26))}${imgIdx}`
          if (!result.has(cellRef)) result.set(cellRef, [])
          result.get(cellRef)!.push({ cell: cellRef, imageUrl: url, imgId: imgIdx })
          pos += extracted.length
        }
      }
    }
  } catch (e) {
    console.warn('[parseXls] OLE parsing failed:', e)
  }

  return { data, headers, images: result }
}

// ============ .xls 图片提取（OLE 复合文档解析）============
const OLE_SIG = new Uint8Array([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1])

type OleEntry = {
  name: string
  type: number   // 0x01=storage, 0x02=stream, 0x05=root
  startSector: number
  size: number
  classId: string
  parent: string
  childSid: number
  siblingSid: number
}

// OLE 小端读取工具
const oleRead16 = (b: Uint8Array, o: number): number => b[o] | (b[o+1] << 8)
const oleRead32 = (b: Uint8Array, o: number): number =>
  b[o] | (b[o+1] << 8) | (b[o+2] << 16) | (b[o+3] << 24)
const oleReadGuid = (b: Uint8Array, o: number): string => {
  const d = oleRead32(b, o)
  const w1 = oleRead16(b, o + 4)
  const w2 = oleRead16(b, o + 6)
  const b1 = b[o + 8]
  const b2 = b[o + 9]
  const hexBytes = Array.from(b.slice(o + 10, o + 16)).map(x => x.toString(16).padStart(2, '0')).join('')
  return `${d.toString(16).padStart(8,'0')}-${w1.toString(16).padStart(4,'0')}-${w2.toString(16).padStart(4,'0')}-${b1.toString(16).padStart(2,'0')}${b2.toString(16).padStart(2,'0')}-${hexBytes}`
}

// 解析目录项（128 字节）
const parseOleEntry = (buf: Uint8Array, offset: number, parent: string): OleEntry | null => {
  const nameLen = oleRead16(buf, offset + 0x40)
  if (nameLen === 0) return null
  let name = ''
  for (let i = 0; i < nameLen; i++) {
    const c = buf[offset + i * 2]
    if (c === 0) break
    name += String.fromCharCode(c)
  }
  // 处理大条目名称（UTF-16LE，以 null 结尾）
  if (!name) {
    const ns = oleRead32(buf, offset + 0x40)
    const so = oleRead32(buf, offset + 0x44)
    const raw = buf.slice(so, so + ns)
    name = String.fromCharCode(...Array.from(raw))
      .split('\0')[0]
  }
  const type = buf[offset + 0x42]
  if (type !== 0x01 && type !== 0x02 && type !== 0x05) return null
  const childSid = oleRead32(buf, offset + 0x44)
  const siblingSid = oleRead32(buf, offset + 0x48)
  const classId = type === 0x02 ? oleReadGuid(buf, offset + 0x50) : ''
  const startSector = oleRead32(buf, offset + 0x70)
  const size = oleRead32(buf, offset + 0x74)
  return { name, type, startSector, size, classId, parent, childSid, siblingSid }
}

// 递归构建目录树（带深度限制）
const buildOleTree = (buf: Uint8Array, entry: OleEntry, all: OleEntry[] = [], depth: number = 0) => {
  if (depth > 100) return all // 防止无限递归
  all.push(entry)
  let sid = entry.childSid
  while (sid !== 0xFFFFFFFE && sid < buf.length / 128) {
    const child = parseOleEntry(buf, sid * 128, entry.parent + '/' + entry.name)
    if (!child) break
    buildOleTree(buf, child, all, depth + 1)
    sid = child.siblingSid
  }
  return all
}

// 读取 FAT 链
const readFatChain = (fat: number[], start: number): number[] => {
  const chain: number[] = []
  let cur = start
  const END = 0xFFFFFFFE
  while (cur !== END && cur < fat.length) {
    chain.push(cur)
    cur = fat[cur]
    if (cur >= fat.length) break
  }
  return chain
}

// 从 OLE 文件读取流内容
const readOleStream = (data: Uint8Array, fat: number[], start: number, size: number, sectorSize: number): Uint8Array => {
  if (size === 0 || start === 0xFFFFFFFE) return new Uint8Array(0)
  const sectors = readFatChain(fat, start)
  const result = new Uint8Array(size)
  let offset = 0
  for (const sec of sectors) {
    if (sec >= fat.length) break
    const chunk = data.slice(sec * sectorSize, (sec + 1) * sectorSize)
    const copy = Math.min(chunk.length, size - offset)
    result.set(chunk.subarray(0, copy), offset)
    offset += copy
    if (offset >= size) break
  }
  return result
}

// 解析 OLE 复合文档，返回 { entries, streams }
const parseOleDocument = (data: Uint8Array) => {
  // 验证签名
  for (let i = 0; i < 8; i++) if (data[i] !== OLE_SIG[i]) throw new Error('不是有效的 OLE 文件')

  const sectorSize = Math.pow(2, oleRead16(data, 0x0E))
  // const miniSectorSize = Math.pow(2, oleRead16(data, 0x0F))
  const numFATsectors = oleRead32(data, 0x2C)
  const firstDirSector = oleRead32(data, 0x30)
  const firstDifatSector = oleRead32(data, 0x44)
  const fatSizePerSector = Math.floor((sectorSize - 512) / 4)

  // 构建 FAT
  const allFatSectors: number[] = []
  for (let i = 0; i < Math.min(numFATsectors, fatSizePerSector); i++) {
    allFatSectors.push(oleRead32(data, 0x40 + i * 4))
  }
  let difatCur = firstDifatSector
  while (allFatSectors.length < numFATsectors && difatCur !== 0xFFFFFFFE) {
    const difatSector = data.slice(difatCur * sectorSize, (difatCur + 1) * sectorSize)
    for (let i = 0; i < fatSizePerSector; i++) {
      if (allFatSectors.length >= numFATsectors) break
      allFatSectors.push(oleRead32(difatSector, i * 4))
    }
    difatCur = oleRead32(difatSector, fatSizePerSector * 4 - 4)
  }

  const fat: number[] = []
  for (const fs of allFatSectors) {
    const sector = data.slice(fs * sectorSize, (fs + 1) * sectorSize)
    for (let i = 0; i < sector.length / 4; i++) fat.push(oleRead32(sector, i * 4))
  }

  // 读取目录扇区
  const dirSectors: number[] = []
  let dirSec = firstDirSector
  while (dirSec !== 0xFFFFFFFE) {
    dirSectors.push(dirSec)
    dirSec = fat[dirSec] ?? 0xFFFFFFFE
  }

  const entries: OleEntry[] = []
  for (const ds of dirSectors) {
    const sector = data.slice(ds * sectorSize, (ds + 1) * sectorSize)
    for (let i = 0; i < sector.length / 128; i++) {
      const entry = parseOleEntry(sector, i * 128, '')
      if (entry) entries.push(entry)
    }
  }

  const rootEntry = entries.find(e => e.type === 0x05) || entries[0]
  const tree = buildOleTree(data, rootEntry)

  // 读取所有流
  const streams: Map<string, Uint8Array> = new Map()
  for (const entry of tree) {
    if (entry.type === 0x02 && entry.size > 0) {
      streams.set(entry.parent + '/' + entry.name, readOleStream(data, fat, entry.startSector, entry.size, sectorSize))
    }
  }

  return { tree, streams, sectorSize }
}

// 在 buf 中查找 pattern（返回起始索引，未找到返回 -1）
const findSubarray = (buf: Uint8Array, pattern: Uint8Array, fromIndex: number): number => {
  if (pattern.length === 0 || pattern.length > buf.length) return -1
  for (let i = fromIndex; i <= buf.length - pattern.length; i++) {
    let found = true
    for (let j = 0; j < pattern.length; j++) {
      if (buf[i + j] !== pattern[j]) { found = false; break }
    }
    if (found) return i
  }
  return -1
}

// 扫描流中的图片数据
const IMAGE_SIGS = [
  { sig: new Uint8Array([0x89, 0x50, 0x4E, 0x47]), mime: 'image/png', marker: 'PNG' },
  { sig: new Uint8Array([0xFF, 0xD8, 0xFF]),      mime: 'image/jpeg', marker: 'JPG' },
  { sig: new Uint8Array([0x47, 0x49, 0x46, 0x38]), mime: 'image/gif',  marker: 'GIF' },
]

// 提取完整 PNG
const extractPng = (buf: Uint8Array, start: number): Uint8Array | null => {
  let pos = start + 8
  const iend = new Uint8Array([0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82])
  while (pos + 8 <= buf.length) {
    const chunk = buf.subarray(pos, pos + 8)
    if (chunk[0] === iend[0] && chunk[1] === iend[1] && chunk[2] === iend[2] && chunk[3] === iend[3]) {
      return buf.subarray(start, pos + 8)
    }
    pos++
  }
  return null
}

// 提取完整 JPG
const extractJpg = (buf: Uint8Array, start: number): Uint8Array | null => {
  let pos = start
  while (pos + 1 < buf.length) {
    if (buf[pos] === 0xFF && (buf[pos+1] === 0xD9 || buf[pos+1] === 0xDA)) {
      return buf.subarray(start, pos + 2)
    }
    pos++
  }
  return null
}

// 提取完整 GIF
const extractGif = (buf: Uint8Array, start: number): Uint8Array | null => {
  const header = new Uint8Array([0x47, 0x49, 0x46, 0x38])
  if (buf[start] !== header[0] || buf[start+1] !== header[1] || buf[start+2] !== header[2] || buf[start+3] !== header[3]) return null
  // GIF 总长度不超过 64KB
  const end = Math.min(start + 65536, buf.length)
  return buf.subarray(start, end)
}

const extractImage = (_sig: Uint8Array, mime: string, buf: Uint8Array, start: number): Uint8Array | null => {
  if (mime === 'image/png') return extractPng(buf, start)
  if (mime === 'image/jpeg') return extractJpg(buf, start)
  if (mime === 'image/gif')  return extractGif(buf, start)
  return null
}

// 将 Uint8Array 转为 base64 DataURL
const toDataURL = (data: Uint8Array, mime: string): string => {
  const chunk = 8192
  let binary = ''
  for (let i = 0; i < data.length; i += chunk) {
    binary += String.fromCharCode(...data.subarray(i, Math.min(i + chunk, data.length)))
  }
  return `data:${mime};base64,${btoa(binary)}`
}

const handleFileChange = async (event: Event) => {
  const target = event.target as HTMLInputElement
  const file = target.files?.[0]
  if (!file) return
  const lowerName = file.name.toLowerCase()
  if (!lowerName.endsWith('.xlsx') && !lowerName.endsWith('.xls')) {
    error.value = '请选择 .xlsx 或 .xls 格式的文件'
    return
  }

  fileName.value = file.name
  error.value = ''
  loading.value = true
  headers.value = []
  tableData.value = []
  cellImages.value = {}
  history.value = []
  historyIndex.value = -1
  originalZip = null
  originalFileType = lowerName.endsWith('.xlsx') ? 'xlsx' : 'xls'

  try {
    const arrayBuffer = await file.arrayBuffer()
    if (originalFileType === 'xlsx') {
      originalZip = await JSZip.loadAsync(arrayBuffer)
      const result = await parseXlsxFromZip(originalZip)
      tableData.value = result.data.slice(0, 200)
      headers.value = result.headers
      await loadImages(originalZip)
    } else {
      const result = await parseXls(arrayBuffer)
      tableData.value = result.data.slice(0, 200)
      headers.value = result.headers
      cellImages.value = {}
      for (const [cellRef, imgs] of result.images) {
        cellImages.value[cellRef] = imgs
      }
    }
    pushHistory()
  } catch (e) {
    error.value = `解析失败: ${e instanceof Error ? e.message : '未知错误'}`
  } finally {
    loading.value = false
  }
}

const startEdit = (rowIndex: number, col: string) => {
  editingCell.value = { row: rowIndex, col }
  editValue.value = tableData.value[rowIndex]?.[col] ?? ''
}

const saveEdit = () => {
  if (!editingCell.value) return
  const { row, col } = editingCell.value
  if (tableData.value[row]) { tableData.value[row][col] = editValue.value; pushHistory() }
  editingCell.value = null
}

const cancelEdit = () => { editingCell.value = null }
const canUndo = () => historyIndex.value > 0
const canRedo = () => historyIndex.value < history.value.length - 1

const exportFile = async () => {
  if (!fileName.value) return
  const baseName = fileName.value.replace(/\.(xls|xlsx)$/, '')

  if (originalFileType === 'xlsx' && originalZip) {
    // 纯字符串替换：只修改 <v> 标签内容，完全保留原始 XML 格式
    const xmlFile = originalZip.file('xl/worksheets/sheet1.xml')
    if (!xmlFile) return
    let xmlText = await xmlFile.async('string')

    tableData.value.forEach((rowData, i) => {
      headers.value.forEach((h) => {
        const colLetter = colToLetter(headers.value.indexOf(h))
        const rowNum = i + 1
        const cellRef = `${colLetter}${rowNum}`
        // 精确匹配 <c r="XN" ...><v>...</v></c>，只替换 <v> 内的值
        const escapedVal = escapeXml(rowData[h] ?? '')
        const regex = new RegExp(
          `(<c\\s+r\\s*=\\s*"${cellRef}"[^>]*>)\\s*<v[^>]*>[^<]*</v>`,
          'g'
        )
        xmlText = xmlText.replace(regex, `$1<v>${escapedVal}</v>`)
      })
    })

    originalZip.file('xl/worksheets/sheet1.xml', xmlText)
    const blob = await originalZip.generateAsync({ type: 'blob' })
    downloadBlob(blob, baseName + '_edited.xlsx')
  } else {
    // xls / 其他：用 xlsx 库重新生成
    const ws = XLSX.utils.aoa_to_sheet(tableData.value.map(row => headers.value.map(h => row[h] ?? '')))
    if (headers.value.length > 0) {
      ws['!cols'] = headers.value.map(() => ({ wch: 20 }))
    }
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1')
    const ext = originalFileType === 'xls' ? 'xls' : 'xlsx'
    const blob = XLSX.write(wb, { bookType: ext as 'xls' | 'xlsx', type: 'blob' as any })
    downloadBlob(blob, baseName + '_edited.' + ext)
  }
}

const downloadBlob = (blob: Blob, filename: string) => {
  const a = document.createElement('a')
  a.href = URL.createObjectURL(blob)
  a.download = filename
  a.click()
  URL.revokeObjectURL(a.href)
}

function escapeXml(str: string): string {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}
</script>

<template>
  <div class="card">
    <div class="card-header">
      <h2 class="card-title">Excel 编辑器</h2>
      <p class="card-desc">上传并预览 Excel 内容，支持 .xls / .xlsx 格式，双击编辑、撤销重做、图片显示（Ctrl+Z 撤销 · Ctrl+Y 重做）</p>
    </div>

    <div class="upload-area" @click="triggerFileInput">
      <input ref="fileInput" type="file" accept=".xlsx,.xls" style="display:none" @change="handleFileChange" />
      <div class="upload-icon">⊞</div>
      <p class="upload-text">{{ fileName || '点击上传 .xlsx 或 .xls 文件' }}</p>
      <p class="upload-hint">预览前 200 行数据，支持图片显示</p>
    </div>

    <div v-if="error" class="error-msg">{{ error }}</div>

    <div v-if="tableData.length > 0" class="toolbar">
      <div class="toolbar-left">
        <span class="info">{{ tableData.length }} 行 × {{ headers.length }} 列</span>
        <span class="format-badge" :class="originalFileType">{{ originalFileType.toUpperCase() }}</span>
        <span v-if="Object.keys(cellImages).length > 0" class="image-badge">🖼 {{ Object.keys(cellImages).length }} 张图片</span>
        <div class="undo-redo">
          <button class="btn-history" :disabled="!canUndo()" @click="undo" title="撤销 (Ctrl+Z)">↩ 撤销</button>
          <button class="btn-history" :disabled="!canRedo()" @click="redo" title="重做 (Ctrl+Y)">↪ 重做</button>
        </div>
      </div>
      <button class="btn-export" @click="exportFile">↓ 导出修改</button>
    </div>

    <div v-loading="loading" class="table-wrap" v-if="!loading && tableData.length > 0">
      <div class="table-container">
        <table class="table">
          <thead>
            <tr>
              <th class="row-num">#</th>
              <th v-for="h in headers" :key="h" class="col-header">{{ h }}</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="(row, i) in tableData" :key="i">
              <td class="row-num">{{ i + 1 }}</td>
              <td v-for="col in headers" :key="col" class="cell" @dblclick="startEdit(i, col)">
                <span v-if="editingCell?.row === i && editingCell?.col === col">
                  <input v-model="editValue" @blur="saveEdit" @keydown.enter="saveEdit" @keydown.escape="cancelEdit" class="edit-input" />
                </span>
                <span v-else class="cell-value">{{ row[col] ?? '' }}</span>
                <div v-if="cellImages[`${col}${i + 1}`]" class="cell-images">
                  <img
                    v-for="(img, idx) in cellImages[`${col}${i + 1}`]"
                    :key="idx"
                    :src="img.imageUrl"
                    class="cell-img"
                    :style="imgStyle(img.imgId)"
                    :title="img.cell"
                  />
                </div>
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>

    <div v-if="!loading && !error && tableData.length === 0 && fileName" class="empty">暂无数据</div>
  </div>
</template>

<style scoped>
.card {
  background: rgba(255, 255, 255, 0.8);
  backdrop-filter: blur(20px);
  -webkit-backdrop-filter: blur(20px);
  border-radius: 20px;
  border: 1px solid rgba(255, 255, 255, 0.6);
  box-shadow: 0 4px 24px rgba(0, 0, 0, 0.06), 0 1px 2px rgba(0, 0, 0, 0.04);
  padding: 32px;
}

.card-header { margin-bottom: 24px; }

.card-title {
  font-size: 22px;
  font-weight: 600;
  color: #1d1d1f;
  letter-spacing: -0.02em;
  margin: 0 0 6px;
}

.card-desc {
  font-size: 14px;
  color: #6e6e73;
  margin: 0;
  line-height: 1.5;
}

.upload-area {
  border: 2px dashed rgba(0, 0, 0, 0.12);
  border-radius: 14px;
  padding: 40px 24px;
  text-align: center;
  cursor: pointer;
  transition: all 0.2s ease;
  background: rgba(249, 249, 249, 0.5);
}

.upload-area:hover {
  border-color: #0071e3;
  background: rgba(0, 113, 227, 0.03);
}

.upload-icon {
  font-size: 28px;
  color: #0071e3;
  margin-bottom: 10px;
}

.upload-text {
  font-size: 15px;
  color: #1d1d1f;
  font-weight: 500;
  margin: 0 0 4px;
}

.upload-hint {
  font-size: 13px;
  color: #86868b;
  margin: 0;
}

.error-msg {
  margin-top: 12px;
  padding: 10px 14px;
  border-radius: 10px;
  background: rgba(255, 59, 48, 0.08);
  color: #ff3b30;
  font-size: 14px;
}

.toolbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-top: 20px;
  padding: 12px 16px;
  background: rgba(0, 0, 0, 0.03);
  border-radius: 10px;
}

.toolbar-left {
  display: flex;
  align-items: center;
  gap: 12px;
  flex-wrap: wrap;
}

.info { font-size: 13px; color: #6e6e73; font-weight: 500; }

.format-badge {
  font-size: 11px;
  font-weight: 600;
  padding: 2px 8px;
  border-radius: 6px;
  letter-spacing: 0.05em;
  text-transform: uppercase;
}
.format-badge.xlsx { background: rgba(16, 185, 129, 0.1); color: #059669; }
.format-badge.xls  { background: rgba(245, 158, 11, 0.1);  color: #d97706; }

.image-badge {
  font-size: 12px;
  color: #0071e3;
  background: rgba(0, 113, 227, 0.08);
  padding: 3px 10px;
  border-radius: 12px;
  font-weight: 500;
}

.undo-redo { display: flex; gap: 6px; }

.btn-history {
  padding: 5px 12px;
  font-size: 12px;
  font-weight: 500;
  color: #6e6e73;
  background: rgba(0, 0, 0, 0.04);
  border: 1px solid rgba(0, 0, 0, 0.08);
  border-radius: 7px;
  cursor: pointer;
  transition: all 0.15s;
  font-family: inherit;
}

.btn-history:hover:not(:disabled) {
  background: rgba(0, 113, 227, 0.08);
  color: #0071e3;
  border-color: rgba(0, 113, 227, 0.2);
}

.btn-history:disabled { opacity: 0.4; cursor: not-allowed; }

.btn-export {
  padding: 7px 16px;
  font-size: 13px;
  font-weight: 500;
  color: #fff;
  background: #0071e3;
  border: none;
  border-radius: 8px;
  cursor: pointer;
  transition: background 0.15s;
  font-family: inherit;
}

.btn-export:hover { background: #0077ed; }

.table-wrap {
  margin-top: 16px;
  overflow-x: auto;
  border-radius: 12px;
  border: 1px solid rgba(0, 0, 0, 0.06);
  max-height: 560px;
  overflow-y: auto;
}

.table-container { position: relative; min-width: 400px; }

.table { width: 100%; border-collapse: collapse; font-size: 13px; }
.table thead { position: sticky; top: 0; z-index: 1; }

.row-num {
  background: #f5f5f7;
  color: #86868b;
  font-weight: 500;
  padding: 8px 12px;
  text-align: center;
  border-bottom: 1px solid rgba(0, 0, 0, 0.06);
  width: 40px;
  font-variant-numeric: tabular-nums;
}

.col-header {
  background: #f5f5f7;
  color: #1d1d1f;
  font-weight: 600;
  padding: 8px 12px;
  text-align: left;
  border-bottom: 1px solid rgba(0, 0, 0, 0.06);
  white-space: nowrap;
  font-size: 12px;
  letter-spacing: 0.02em;
  text-transform: uppercase;
}

.cell {
  padding: 6px 12px;
  border-bottom: 1px solid rgba(0, 0, 0, 0.04);
  min-width: 80px;
  max-width: 240px;
  color: #1d1d1f;
  cursor: default;
  position: relative;
  vertical-align: middle;
}

.cell:hover { background: rgba(0, 113, 227, 0.04); }

.cell-value {
  display: block;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.edit-input {
  width: 100%;
  padding: 4px 8px;
  font-size: 13px;
  border: 2px solid #0071e3;
  border-radius: 6px;
  outline: none;
  background: #fff;
  color: #1d1d1f;
  font-family: inherit;
}

.cell-images {
  position: absolute;
  bottom: 2px;
  right: 2px;
  display: flex;
  flex-direction: column;
  gap: 2px;
  pointer-events: none;
  z-index: 2;
}

.cell-img {
  border-radius: 4px;
  box-shadow: 0 1px 4px rgba(0, 0, 0, 0.15);
  display: block;
}

.empty {
  margin-top: 24px;
  text-align: center;
  color: #86868b;
  font-size: 14px;
}
</style>
