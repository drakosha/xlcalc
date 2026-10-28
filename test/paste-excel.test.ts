/** Paste Special against the answers of Excel for Mac (`bench/coverage/paste.ts`), as `paste-sets.ts` checks them. */
import { PASTE_KNOWN, PASTE_SETS, pasteSets } from './paste-sets.js'

pasteSets('Excel for Mac', 'paste', PASTE_SETS, PASTE_KNOWN)
