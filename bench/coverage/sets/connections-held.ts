/**
 * The functions that reach outside the workbook, computed where they reach nothing: cube functions over a connection
 * the workbook does not have, RTD without its server, WEBSERVICE and IMAGE at an address nobody answers (TEST-NET-1,
 * 192.0.2.1), STOCKHISTORY, TRANSLATE and DETECTLANGUAGE without a sign-in, CALL and REGISTER.ID of a library that does
 * not exist, EUROCONVERT without its add-in, and whether Excel knows ECMA.CEILING. Each also with an error and a wrong
 * kind of value in its arguments, which decide before the connection does. A set of Excel for Windows alone: Excel for
 * Mac asks for sign-ins and web access in windows of its own. `PREDICTED`, from Microsoft's descriptions of the
 * functions, was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

const NOWHERE = '"http://192.0.2.1/a"'

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  // cube functions: a connection the workbook does not have is #NAME?
  ['CUBEMEMBER', 'CUBEMEMBER("Sales","[Time].[2025]")', e('NAME')],
  ['CUBEMEMBER', 'CUBEMEMBER("Sales","[Time].[2025]","Year")', e('NAME')],
  ['CUBEMEMBER', 'CUBEMEMBER(NA(),"[Time].[2025]")', e('NA')],
  ['CUBEMEMBER', 'CUBEMEMBER("Sales",5)', e('NAME')],
  ['CUBEVALUE', 'CUBEVALUE("Sales","[Measures].[Amount]")', e('NAME')],
  ['CUBEVALUE', 'CUBEVALUE("Sales")', e('NAME')],
  ['CUBEVALUE', 'CUBEVALUE("Sales",NA())', e('NA')],
  ['CUBEVALUE', 'CUBEVALUE("","[Measures].[Amount]")', e('NAME')],
  ['CUBESET', 'CUBESET("Sales","[Time].Children")', e('NAME')],
  ['CUBESET', 'CUBESET("Sales","[Time].Children","Years",1,"[Measures].[Amount]")', e('NAME')],
  ['CUBESET', 'CUBESET("Sales","[Time].Children","Years",9)', e('NAME')],
  ['CUBESETCOUNT', 'CUBESETCOUNT(CUBESET("Sales","[Time].Children"))', e('NAME')],
  ['CUBESETCOUNT', 'CUBESETCOUNT("[Time].Children")', e('VALUE')],
  ['CUBESETCOUNT', 'CUBESETCOUNT(5)', e('VALUE')],
  ['CUBESETCOUNT', 'CUBESETCOUNT(NA())', e('NA')],
  ['CUBERANKEDMEMBER', 'CUBERANKEDMEMBER("Sales","[Time].Children",1)', e('NAME')],
  ['CUBERANKEDMEMBER', 'CUBERANKEDMEMBER("Sales","[Time].Children",0)', e('NAME')],
  ['CUBEMEMBERPROPERTY', 'CUBEMEMBERPROPERTY("Sales","[Time].[2025]","Key")', e('NAME')],
  ['CUBEKPIMEMBER', 'CUBEKPIMEMBER("Sales","Margin",1)', e('NAME')],
  ['CUBEKPIMEMBER', 'CUBEKPIMEMBER("Sales","Margin",9)', e('NAME')],
  // RTD without a server registered
  ['RTD', 'RTD("NoSuch.Server","","Topic")', e('NA')],
  ['RTD', 'RTD("NoSuch.Server","","Topic","Two")', e('NA')],
  ['RTD', 'RTD(NA(),"","Topic")', e('NA')],
  // the web at an address nobody answers, and text that is no address
  ['WEBSERVICE', `WEBSERVICE(${NOWHERE})`, e('VALUE')],
  ['WEBSERVICE', 'WEBSERVICE("abc")', e('VALUE')],
  ['WEBSERVICE', 'WEBSERVICE(REPT("a",2049))', e('VALUE')],
  ['WEBSERVICE', 'WEBSERVICE(NA())', e('NA')],
  ['WEBSERVICE', 'WEBSERVICE(5)', e('VALUE')],
  ['IMAGE', `IMAGE(${NOWHERE})`, e('CONNECT')],
  ['IMAGE', 'IMAGE("abc")', e('VALUE')],
  ['IMAGE', `IMAGE(${NOWHERE},"alt",5)`, e('VALUE')],
  ['IMAGE', 'IMAGE(NA())', e('NA')],
  // services that need a sign-in
  ['STOCKHISTORY', 'STOCKHISTORY("MSFT",45658)', e('BLOCKED')],
  ['STOCKHISTORY', 'STOCKHISTORY("MSFT",45658,45688,0,1,0,1)', e('BLOCKED')],
  ['STOCKHISTORY', 'STOCKHISTORY("MSFT","abc")', e('VALUE')],
  ['STOCKHISTORY', 'STOCKHISTORY(NA(),45658)', e('NA')],
  ['TRANSLATE', 'TRANSLATE("hello","en","fr")', e('BLOCKED')],
  ['TRANSLATE', 'TRANSLATE("hello")', e('BLOCKED')],
  ['DETECTLANGUAGE', 'DETECTLANGUAGE("hello")', e('BLOCKED')],
  ['DETECTLANGUAGE', 'DETECTLANGUAGE(NA())', e('NA')],
  // add-ins: a library that does not exist, the Euro Currency Tools not loaded
  ['CALL', 'CALL("NoSuchLibrary.dll","NoSuchFunction","J")', e('VALUE')],
  ['REGISTER.ID', 'REGISTER.ID("NoSuchLibrary.dll","NoSuchFunction","J")', e('VALUE')],
  ['REGISTER.ID', 'REGISTER.ID("NoSuchLibrary.dll","NoSuchFunction")', e('VALUE')],
  ['EUROCONVERT', 'EUROCONVERT(1,"EUR","DEM")', e('NAME')],
  ['EUROCONVERT', 'EUROCONVERT(1,"EUR","DEM",TRUE,3)', e('NAME')],
  // ECMA.CEILING, the name of CEILING in the file format's standard
  ['ECMA.CEILING', 'ECMA.CEILING(4.3,2)', n(6)],
  ['ECMA.CEILING', 'ECMA.CEILING(-4.3,2)', n(-4)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
