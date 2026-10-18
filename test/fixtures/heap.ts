/** The heap used, in MB, after two turns of the event loop and two collections (`--expose-gc`). */
export async function heap(): Promise<number> {
  const gc = (globalThis as { gc: () => void }).gc
  const turn = () => new Promise((done) => setImmediate(done))
  await turn()
  await turn()
  gc()
  gc()
  return process.memoryUsage().heapUsed / 2 ** 20
}
