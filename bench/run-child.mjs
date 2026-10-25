/**
 * A forked child that messages `{type: 'phase', name}` as it goes and `{type: 'result', result}` at the end, run under
 * a watchdog that SIGKILLs it when its current phase takes longer than `phaseLimitMs`. Resolves with the raw facts —
 * the result message, whether it timed out, the phase it reached, when it began and how long it ran, the tail of its
 * stderr, its exit code and signal, and the fork's error if it never ran — for the caller to classify
 * (`hyperformula/run.mjs`, `compare/run.mjs`).
 */
import { fork } from 'node:child_process'

export function runChild(url, args, { cwd, execArgv, env, phaseLimitMs, stderrTail }) {
  return new Promise((resolve) => {
    const started = performance.now()
    const child = fork(url, args, { cwd, silent: true, execArgv, env })
    let timer, phase = 'startup', phaseStarted = started, received, timedOut = false, stderr = '', error
    const arm = () => {
      clearTimeout(timer)
      timer = setTimeout(() => { timedOut = true; child.kill('SIGKILL') }, phaseLimitMs)
    }
    arm()
    child.stderr.on('data', (chunk) => { stderr = (stderr + chunk).slice(-stderrTail) })
    child.stdout.resume()
    child.on('message', (m) => {
      if (m.type === 'phase') { phase = m.name; phaseStarted = performance.now(); arm() }
      if (m.type === 'result') { received = m.result; clearTimeout(timer) }
    })
    child.on('error', (e) => {
      clearTimeout(timer)
      resolve({ received, timedOut, phase, phaseElapsedMs: performance.now() - phaseStarted, started, stderr, error: e })
    })
    child.on('exit', (code, signal) => {
      clearTimeout(timer)
      resolve({ received, timedOut, phase, phaseElapsedMs: performance.now() - phaseStarted, started, code, signal, stderr })
    })
  })
}
