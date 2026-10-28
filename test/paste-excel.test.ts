/** Paste Special against the answers of Excel for Mac (`bench/coverage/paste.ts`), as `paste-sets.ts` checks them. */
import { asRu, PASTE_KNOWN, PASTE_RU_SETS, PASTE_SETS, pasteSets } from './paste-sets.js'

pasteSets('Excel for Mac', 'paste', PASTE_SETS, PASTE_KNOWN)
pasteSets('Excel for Mac ru-RU', 'paste/ru-RU', PASTE_RU_SETS, {}, asRu)
