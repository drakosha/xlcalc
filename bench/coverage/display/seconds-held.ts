/**
 * Held out: the values of the case set `seconds-held` (`sets/seconds-held.ts`), moments at the edge of a second, shown by
 * the format `d h:mm:ss` in a column wide enough: the day and the time of day a cell shows take the same rounding to the
 * second as TEXT. Each answer was predicted by the engine and committed before Excel showed the case.
 */
import type { DisplayCase } from './cases.js'

export function buildDisplayCases(): DisplayCase[] {
  return [
    {"id":"day#1","cell":{"value":45688.430907407404,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 10:20:30"}},
    {"id":"day#2","cell":{"value":45688.430908553244,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 10:20:30"}},
    {"id":"day#3","cell":{"value":45688.43090855787,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 10:20:30"}},
    {"id":"day#4","cell":{"value":45688.430908560185,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 10:20:30"}},
    {"id":"day#5","cell":{"value":45688.430908563656,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 10:20:30"}},
    {"id":"day#6","cell":{"value":45688.4309085647,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 10:20:30"}},
    {"id":"day#7","cell":{"value":45688.4309085648,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 10:20:30"}},
    {"id":"day#8","cell":{"value":45688.43090856481,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 10:20:30"}},
    {"id":"day#9","cell":{"value":45688.43090856482,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 10:20:30"}},
    {"id":"day#10","cell":{"value":45688.43090856597,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 10:20:31"}},
    {"id":"day#11","cell":{"value":444615.13495949074,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"24 3:14:21"}},
    {"id":"day#12","cell":{"value":2504553.1110706017,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"25 2:39:56"}},
    {"id":"day#13","cell":{"value":2577729.1113368054,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"30 2:40:20"}},
    {"id":"day#14","cell":{"value":1963420.4880381944,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"28 11:42:46"}},
    {"id":"day#15","cell":{"value":2693171.0251446757,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"24 0:36:13"}},
    {"id":"day#16","cell":{"value":2072685.826105324,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"25 19:49:35"}},
    {"id":"day#17","cell":{"value":45688.00390625,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 0:05:38"}},
    {"id":"day#18","cell":{"value":45688.01171875,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 0:16:53"}},
    {"id":"day#19","cell":{"value":45688.99609375,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 23:54:23"}},
    {"id":"day#20","cell":{"value":45688.99999420602,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 23:59:59"}},
    {"id":"day#21","cell":{"value":45688.99999420833,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 23:59:59"}},
    {"id":"day#22","cell":{"value":45688.99999421295,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"1 0:00:00"}},
    {"id":"day#23","cell":{"value":45688.999994212965,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"1 0:00:00"}},
    {"id":"day#24","cell":{"value":46022.999994210055,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 23:59:59"}},
    {"id":"day#25","cell":{"value":46022.99999421295,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"1 0:00:00"}},
    {"id":"day#26","cell":{"value":59.99999421296295,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"29 0:00:00"}},
    {"id":"day#27","cell":{"value":59.99999421296296,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"29 0:00:00"}},
    {"id":"day#28","cell":{"value":0.9999942129629625,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"1 0:00:00"}},
    {"id":"day#29","cell":{"value":999999.9999942129,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"26 0:00:00"}},
    {"id":"day#30","cell":{"value":1020183.9999942129,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"28 23:59:59"}},
    {"id":"day#31","cell":{"value":1000000.999994213,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"26 23:59:59"}},
    {"id":"day#32","cell":{"value":2958464.9999942128,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"30 23:59:59"}},
    {"id":"day#33","cell":{"value":2958464.999994213,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"30 23:59:59"}},
    {"id":"day#34","cell":{"value":2958465.999994213,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 23:59:59"}},
    {"id":"day#35","cell":{"value":2958465.9999942128,"format":"d h:mm:ss"},"width":20,"predicted":{"text":"31 23:59:59"}},
  ]
}
