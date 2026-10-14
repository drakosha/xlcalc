/** Where the harness's files are: its own folder (`HERE`, beside its AppleScripts) and the repository's root. */
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const HERE = fileURLToPath(new URL('.', import.meta.url))
export const ROOT = join(HERE, '../..')
