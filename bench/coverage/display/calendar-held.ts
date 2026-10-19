/**
 * Held-out display cases of the calendar prefixes (Excel en-US, `display calendar-held`), each answer predicted and committed before
 * Excel showed the case. `B1` before a date code is the Gregorian calendar and `B2` the Hijri one (ECMA-376 numFmts;
 * Excel's help: Hijri 9/25/1420 is 1/1/2000, the tabular calendar). A `B1` date is predicted as the engine shows the code
 * after the prefix, a `B2` one by the Hijri day, month and year in the numeric fields of its code, its time as it is.
 * Codes with the names of Hijri months wait for the next set.
 */
import type { DisplayCase } from './cases.js'

export function buildDisplayCases(): DisplayCase[] {
  return [
    {"id":"b1#1","cell":{"value":36526,"format":"B1dd/mm/yyyy"},"width":30,"predicted":{"text":"01/01/2000"}},
    {"id":"b1#2","cell":{"value":36526,"format":"B1yyyy"},"width":30,"predicted":{"text":"2000"}},
    {"id":"b1#3","cell":{"value":36526,"format":"B1d mmmm yyyy"},"width":30,"predicted":{"text":"1 January 2000"}},
    {"id":"b1#4","cell":{"value":36526,"format":"b1m/d/yy"},"width":30,"predicted":{"text":"1/1/00"}},
    {"id":"b1#5","cell":{"value":45688.5,"format":"B1dd/mm/yyyy"},"width":30,"predicted":{"text":"31/01/2025"}},
    {"id":"b1#6","cell":{"value":45688.5,"format":"B1yyyy"},"width":30,"predicted":{"text":"2025"}},
    {"id":"b1#7","cell":{"value":45688.5,"format":"B1d mmmm yyyy"},"width":30,"predicted":{"text":"31 January 2025"}},
    {"id":"b1#8","cell":{"value":45688.5,"format":"b1m/d/yy"},"width":30,"predicted":{"text":"1/31/25"}},
    {"id":"b1#9","cell":{"value":45001.25,"format":"B1dd/mm/yyyy"},"width":30,"predicted":{"text":"16/03/2023"}},
    {"id":"b1#10","cell":{"value":45001.25,"format":"B1yyyy"},"width":30,"predicted":{"text":"2023"}},
    {"id":"b1#11","cell":{"value":45001.25,"format":"B1d mmmm yyyy"},"width":30,"predicted":{"text":"16 March 2023"}},
    {"id":"b1#12","cell":{"value":45001.25,"format":"b1m/d/yy"},"width":30,"predicted":{"text":"3/16/23"}},
    {"id":"b1#13","cell":{"value":44562.75,"format":"B1dd/mm/yyyy"},"width":30,"predicted":{"text":"01/01/2022"}},
    {"id":"b1#14","cell":{"value":44562.75,"format":"B1yyyy"},"width":30,"predicted":{"text":"2022"}},
    {"id":"b1#15","cell":{"value":44562.75,"format":"B1d mmmm yyyy"},"width":30,"predicted":{"text":"1 January 2022"}},
    {"id":"b1#16","cell":{"value":44562.75,"format":"b1m/d/yy"},"width":30,"predicted":{"text":"1/1/22"}},
    {"id":"b2#1","cell":{"value":36526,"format":"B2dd/mm/yyyy"},"width":30,"predicted":{"text":"25/09/1420"}},
    {"id":"b2#2","cell":{"value":36526,"format":"B2yyyy"},"width":30,"predicted":{"text":"1420"}},
    {"id":"b2#3","cell":{"value":36526,"format":"B2d/m/yy"},"width":30,"predicted":{"text":"25/9/20"}},
    {"id":"b2#4","cell":{"value":36526,"format":"b2dd/mm/yyyy"},"width":30,"predicted":{"text":"25/09/1420"}},
    {"id":"b2#5","cell":{"value":36526,"format":"B2dd/mm/yyyy h:mm"},"width":30,"predicted":{"text":"25/09/1420 0:00"}},
    {"id":"b2#6","cell":{"value":45688.5,"format":"B2dd/mm/yyyy"},"width":30,"predicted":{"text":"02/08/1446"}},
    {"id":"b2#7","cell":{"value":45688.5,"format":"B2yyyy"},"width":30,"predicted":{"text":"1446"}},
    {"id":"b2#8","cell":{"value":45688.5,"format":"B2d/m/yy"},"width":30,"predicted":{"text":"2/8/46"}},
    {"id":"b2#9","cell":{"value":45688.5,"format":"b2dd/mm/yyyy"},"width":30,"predicted":{"text":"02/08/1446"}},
    {"id":"b2#10","cell":{"value":45688.5,"format":"B2dd/mm/yyyy h:mm"},"width":30,"predicted":{"text":"02/08/1446 12:00"}},
    {"id":"b2#11","cell":{"value":45001.25,"format":"B2dd/mm/yyyy"},"width":30,"predicted":{"text":"24/08/1444"}},
    {"id":"b2#12","cell":{"value":45001.25,"format":"B2yyyy"},"width":30,"predicted":{"text":"1444"}},
    {"id":"b2#13","cell":{"value":45001.25,"format":"B2d/m/yy"},"width":30,"predicted":{"text":"24/8/44"}},
    {"id":"b2#14","cell":{"value":45001.25,"format":"b2dd/mm/yyyy"},"width":30,"predicted":{"text":"24/08/1444"}},
    {"id":"b2#15","cell":{"value":45001.25,"format":"B2dd/mm/yyyy h:mm"},"width":30,"predicted":{"text":"24/08/1444 6:00"}},
    {"id":"b2#16","cell":{"value":44562.75,"format":"B2dd/mm/yyyy"},"width":30,"predicted":{"text":"28/05/1443"}},
    {"id":"b2#17","cell":{"value":44562.75,"format":"B2yyyy"},"width":30,"predicted":{"text":"1443"}},
    {"id":"b2#18","cell":{"value":44562.75,"format":"B2d/m/yy"},"width":30,"predicted":{"text":"28/5/43"}},
    {"id":"b2#19","cell":{"value":44562.75,"format":"b2dd/mm/yyyy"},"width":30,"predicted":{"text":"28/05/1443"}},
    {"id":"b2#20","cell":{"value":44562.75,"format":"B2dd/mm/yyyy h:mm"},"width":30,"predicted":{"text":"28/05/1443 18:00"}},
  ]
}
