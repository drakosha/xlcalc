/** Paste Special against the answers of Excel for Mac (`bench/coverage/paste.ts`), as `paste-sets.ts` checks them. */
import { asRu, INTO_MERGED, PASTE_KNOWN, PASTE_RU_SETS, PASTE_SETS, pasteSets } from './paste-sets.js'

/** Where the engine pastes otherwise than Excel for Mac: the known differences of both Excels, and All Merging Conditional Formats into a merged area, which the Mac pasted through its window as Windows pastes it. */
const MAC_KNOWN: Record<string, string> = { ...PASTE_KNOWN, 'merging-mac/into-merged': INTO_MERGED }

pasteSets('Excel for Mac', 'paste', PASTE_SETS, MAC_KNOWN)
pasteSets('Excel for Mac ru-RU', 'paste/ru-RU', PASTE_RU_SETS, {}, asRu)
