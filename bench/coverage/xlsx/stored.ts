/**
 * Books whose formulas call functions the engine computes no value for, as Excel computes and saves them: OFFSET and
 * a function no add-in gives (`MYADDIN`), which the engine does not know, a cube function without its connection,
 * WEBSERVICE at an address nobody answers; IMAGE and STOCKHISTORY, whose services Excel does not reach. `npx tsx bench/windows/run.ts xlsx stored` has Excel for Windows compute
 * and save each into `test/excel/win/xlsx/stored/`; `test/stored-held-excel.test.ts` reads what Excel saved and checks
 * that the engine shows the values the file stores until it calculates the cells (`unknownFunctions: 'held'`).
 */
import type { BookInput } from '../../../src/contract.js'
import { writeXlsx } from '../../../src/xlsx.js'

export interface StoredFile {
  id: string
  about: string
  book: BookInput
}

export const STORED_FILES: StoredFile[] = [
  {
    id: 'outside',
    about: 'OFFSET over A1 and K1, an add-in function that does not exist, IFERROR over it, a cube value, a web service, and cells reading them',
    book: {
      sheets: [
        {
          name: 'S',
          cells: {
            A1: { value: 2 },
            K1: { value: 4 },
            B1: { formula: 'OFFSET(K1,0,0)*5' },
            C1: { formula: 'B1+1' },
            D1: { formula: 'MYADDIN(A1)*2' },
            E1: { formula: 'IFERROR(MYADDIN(A1),"none")' },
            F1: { formula: 'CUBEVALUE("Sales","[Measures].[Amount]")' },
            G1: { formula: 'WEBSERVICE("http://192.0.2.1/a")' },
            H1: { formula: 'IFERROR(F1,1)+IFERROR(G1,2)' },
          },
        },
      ],
    },
  },
  {
    id: 'blocked',
    about: "IMAGE and STOCKHISTORY, #BLOCKED! where Microsoft's services are not reached, ERROR.TYPE over them: how Excel stores an error of connections",
    book: {
      sheets: [
        {
          name: 'S',
          cells: {
            B1: { formula: 'IMAGE("https://192.0.2.1/a.png")' },
            C1: { formula: 'STOCKHISTORY("MSFT",45000)' },
            D1: { formula: 'ERROR.TYPE(B1)' },
            E1: { formula: 'ERROR.TYPE(C1)' },
          },
        },
      ],
    },
  },
]

/** A file's bytes as `writeXlsx` writes its book. */
export const storedFile = async (f: StoredFile): Promise<Uint8Array> => (await writeXlsx({ book: f.book })).data
