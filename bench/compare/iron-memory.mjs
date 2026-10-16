// node --expose-gc iron-memory.mjs <workload>: IronCalc's retained memory after opening a book, in a fresh process: the
// JavaScript heap and the WebAssembly memory, which grows and is never given back, so its growth is an upper bound.
import { ironMemory } from './iron.mjs'
import { workload } from '../hyperformula/workloads.mjs'
import { open, prepareInput } from '../hyperformula/shared.mjs'

const [id] = process.argv.slice(2)
const { book } = workload(id)
const prepared = prepareInput('ironcalc', book)
global.gc()
const heap = process.memoryUsage().heapUsed, wasm = ironMemory.buffer.byteLength
const held = open('ironcalc', book, prepared)
global.gc()
process.send?.({ type: 'result', result: { status: 'ok', heapMb: (process.memoryUsage().heapUsed - heap) / 2 ** 20, wasmMb: (ironMemory.buffer.byteLength - wasm) / 2 ** 20, held: typeof held } })
process.disconnect?.()
