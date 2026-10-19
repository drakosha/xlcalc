/**
 * Held-out display cases of names under the calendar prefixes (Excel ru-RU, `display calendar-names-held-ru ru`), each answer predicted and committed
 * before Excel showed the case, after the set `calendar-held`. `B1` showed English month names in ru-RU there: its
 * month and day names are predicted as `[$-409]` shows them. `B2`'s months are predicted by the Unicode CLDR names of the
 * tabular Hijri calendar in the locale, `mmm` as `mmmm` (Hijri months have no abbreviations, VBA's reference to
 * `Format`); a weekday as the engine names it.
 */
import type { DisplayCase } from './cases.js'

export function buildDisplayCases(): DisplayCase[] {
  return [
    {"id":"b1n#1","cell":{"value":36526,"format":"B1d mmmm yyyy"},"width":40,"predicted":{"text":"1 January 2000"}},
    {"id":"b1n#2","cell":{"value":36526,"format":"B1dddd, d mmm yyyy"},"width":40,"predicted":{"text":"Saturday, 1 Jan 2000"}},
    {"id":"b1n#3","cell":{"value":36526,"format":"B1mmm"},"width":40,"predicted":{"text":"Jan"}},
    {"id":"b1n#4","cell":{"value":36526,"format":"B1ddd"},"width":40,"predicted":{"text":"Sat"}},
    {"id":"b1n#5","cell":{"value":45688.5,"format":"B1d mmmm yyyy"},"width":40,"predicted":{"text":"31 January 2025"}},
    {"id":"b1n#6","cell":{"value":45688.5,"format":"B1dddd, d mmm yyyy"},"width":40,"predicted":{"text":"Friday, 31 Jan 2025"}},
    {"id":"b1n#7","cell":{"value":45688.5,"format":"B1mmm"},"width":40,"predicted":{"text":"Jan"}},
    {"id":"b1n#8","cell":{"value":45688.5,"format":"B1ddd"},"width":40,"predicted":{"text":"Fri"}},
    {"id":"b1n#9","cell":{"value":45001.25,"format":"B1d mmmm yyyy"},"width":40,"predicted":{"text":"16 March 2023"}},
    {"id":"b1n#10","cell":{"value":45001.25,"format":"B1dddd, d mmm yyyy"},"width":40,"predicted":{"text":"Thursday, 16 Mar 2023"}},
    {"id":"b1n#11","cell":{"value":45001.25,"format":"B1mmm"},"width":40,"predicted":{"text":"Mar"}},
    {"id":"b1n#12","cell":{"value":45001.25,"format":"B1ddd"},"width":40,"predicted":{"text":"Thu"}},
    {"id":"b1n#13","cell":{"value":44562.75,"format":"B1d mmmm yyyy"},"width":40,"predicted":{"text":"1 January 2022"}},
    {"id":"b1n#14","cell":{"value":44562.75,"format":"B1dddd, d mmm yyyy"},"width":40,"predicted":{"text":"Saturday, 1 Jan 2022"}},
    {"id":"b1n#15","cell":{"value":44562.75,"format":"B1mmm"},"width":40,"predicted":{"text":"Jan"}},
    {"id":"b1n#16","cell":{"value":44562.75,"format":"B1ddd"},"width":40,"predicted":{"text":"Sat"}},
    {"id":"b2n#1","cell":{"value":36526,"format":"B2mmmm"},"width":40,"predicted":{"text":"рамадан"}},
    {"id":"b2n#2","cell":{"value":36526,"format":"B2mmm"},"width":40,"predicted":{"text":"рамадан"}},
    {"id":"b2n#3","cell":{"value":36526,"format":"B2d mmmm yyyy"},"width":40,"predicted":{"text":"25 рамадан 1420"}},
    {"id":"b2n#4","cell":{"value":36526,"format":"B2dddd"},"width":40,"predicted":{"text":"суббота"}},
    {"id":"b2n#5","cell":{"value":45688.5,"format":"B2mmmm"},"width":40,"predicted":{"text":"шаабан"}},
    {"id":"b2n#6","cell":{"value":45688.5,"format":"B2mmm"},"width":40,"predicted":{"text":"шаабан"}},
    {"id":"b2n#7","cell":{"value":45688.5,"format":"B2d mmmm yyyy"},"width":40,"predicted":{"text":"2 шаабан 1446"}},
    {"id":"b2n#8","cell":{"value":45688.5,"format":"B2dddd"},"width":40,"predicted":{"text":"пятница"}},
    {"id":"b2n#9","cell":{"value":45001.25,"format":"B2mmmm"},"width":40,"predicted":{"text":"шаабан"}},
    {"id":"b2n#10","cell":{"value":45001.25,"format":"B2mmm"},"width":40,"predicted":{"text":"шаабан"}},
    {"id":"b2n#11","cell":{"value":45001.25,"format":"B2d mmmm yyyy"},"width":40,"predicted":{"text":"24 шаабан 1444"}},
    {"id":"b2n#12","cell":{"value":45001.25,"format":"B2dddd"},"width":40,"predicted":{"text":"четверг"}},
    {"id":"b2n#13","cell":{"value":44562.75,"format":"B2mmmm"},"width":40,"predicted":{"text":"джумад-уль-авваль"}},
    {"id":"b2n#14","cell":{"value":44562.75,"format":"B2mmm"},"width":40,"predicted":{"text":"джумад-уль-авваль"}},
    {"id":"b2n#15","cell":{"value":44562.75,"format":"B2d mmmm yyyy"},"width":40,"predicted":{"text":"28 джумад-уль-авваль 1443"}},
    {"id":"b2n#16","cell":{"value":44562.75,"format":"B2dddd"},"width":40,"predicted":{"text":"суббота"}},
  ]
}
