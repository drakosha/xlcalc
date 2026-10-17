// node test/suites.mjs [vitest options]: the tests in two runs, one after the other: first all but those that time
// themselves (vitest.config.ts, project `main`), then those alone, one file at a time (project `timed`), so that the
// load of the first run does not fail them. Both run whatever the first gives; the exit code is the worse.
import { spawnSync } from 'node:child_process'

const vitest = (project) => spawnSync(process.execPath, ['node_modules/vitest/vitest.mjs', 'run', '--project', project, ...process.argv.slice(2)], { stdio: 'inherit' }).status ?? 1
const codes = ['main', 'timed'].map(vitest)
process.exit(Math.max(...codes))
