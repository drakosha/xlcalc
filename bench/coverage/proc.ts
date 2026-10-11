/** Waiting and processes, for the harnesses of both platforms. */

export const sleep = (ms: number) => new Promise<void>((r) => setTimeout(r, ms))

/** Whether a process of this machine still runs. */
export function alive(pid: number): boolean {
  try {
    process.kill(pid, 0)
    return true
  } catch {
    return false
  }
}
