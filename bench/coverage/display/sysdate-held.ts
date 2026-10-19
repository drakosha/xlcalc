/**
 * Held-out display cases of the system's long date and time tags (Excel en-US, `display sysdate-held`), each answer predicted and
 * committed before Excel showed the case. `[$-x-sysdate]` and `[$-x-systime]` are the culture tags of the system's long
 * date and time ([MS-XLSX] formatCode16), as `[$-F800]` and `[$-F400]` are their codes (ECMA-376 numFmts): each is
 * predicted to show what the engine shows for `[$-F800]` (`[$-F400]`) followed by the same code, the system's format
 * whatever the code after the tag, and the cells `f800` and `f400` show the same codes with `[$-F800]` and `[$-F400]` in the
 * same run, as controls of the system format Excel runs with. A currency's tag is its symbol up to the first hyphen,
 * then the culture (`[$€-x-euro2]`, `[$₽-ru-RU]`): predicted to show the symbol alone.
 */
import type { DisplayCase } from './cases.js'

export function buildDisplayCases(): DisplayCase[] {
  return [
    {"id":"sysdate#1","cell":{"value":45688.5,"format":"[$-x-sysdate]dddd, mmmm dd, yyyy"},"width":40,"predicted":{"text":"Friday, January 31, 2025"}},
    {"id":"f800#1","cell":{"value":45688.5,"format":"[$-F800]dddd, mmmm dd, yyyy"},"width":40,"predicted":{"text":"Friday, January 31, 2025"}},
    {"id":"sysdate#2","cell":{"value":45688.5,"format":"[$-x-sysdate]yyyy-mm-dd"},"width":40,"predicted":{"text":"Friday, January 31, 2025"}},
    {"id":"f800#2","cell":{"value":45688.5,"format":"[$-F800]yyyy-mm-dd"},"width":40,"predicted":{"text":"Friday, January 31, 2025"}},
    {"id":"sysdate#3","cell":{"value":45688.5,"format":"[$-x-sysdate]d mmm"},"width":40,"predicted":{"text":"Friday, January 31, 2025"}},
    {"id":"f800#3","cell":{"value":45688.5,"format":"[$-F800]d mmm"},"width":40,"predicted":{"text":"Friday, January 31, 2025"}},
    {"id":"sysdate#4","cell":{"value":45688.5,"format":"[$-x-sysdate]mmmm"},"width":40,"predicted":{"text":"Friday, January 31, 2025"}},
    {"id":"f800#4","cell":{"value":45688.5,"format":"[$-F800]mmmm"},"width":40,"predicted":{"text":"Friday, January 31, 2025"}},
    {"id":"sysdate#5","cell":{"value":45688.5,"format":"[$-x-sysdate]h:mm"},"width":40,"predicted":{"text":"Friday, January 31, 2025"}},
    {"id":"f800#5","cell":{"value":45688.5,"format":"[$-F800]h:mm"},"width":40,"predicted":{"text":"Friday, January 31, 2025"}},
    {"id":"sysdate#6","cell":{"value":45001.25,"format":"[$-x-sysdate]dddd, mmmm dd, yyyy"},"width":40,"predicted":{"text":"Thursday, March 16, 2023"}},
    {"id":"f800#6","cell":{"value":45001.25,"format":"[$-F800]dddd, mmmm dd, yyyy"},"width":40,"predicted":{"text":"Thursday, March 16, 2023"}},
    {"id":"sysdate#7","cell":{"value":45001.25,"format":"[$-x-sysdate]yyyy-mm-dd"},"width":40,"predicted":{"text":"Thursday, March 16, 2023"}},
    {"id":"f800#7","cell":{"value":45001.25,"format":"[$-F800]yyyy-mm-dd"},"width":40,"predicted":{"text":"Thursday, March 16, 2023"}},
    {"id":"sysdate#8","cell":{"value":45001.25,"format":"[$-x-sysdate]d mmm"},"width":40,"predicted":{"text":"Thursday, March 16, 2023"}},
    {"id":"f800#8","cell":{"value":45001.25,"format":"[$-F800]d mmm"},"width":40,"predicted":{"text":"Thursday, March 16, 2023"}},
    {"id":"sysdate#9","cell":{"value":45001.25,"format":"[$-x-sysdate]mmmm"},"width":40,"predicted":{"text":"Thursday, March 16, 2023"}},
    {"id":"f800#9","cell":{"value":45001.25,"format":"[$-F800]mmmm"},"width":40,"predicted":{"text":"Thursday, March 16, 2023"}},
    {"id":"sysdate#10","cell":{"value":45001.25,"format":"[$-x-sysdate]h:mm"},"width":40,"predicted":{"text":"Thursday, March 16, 2023"}},
    {"id":"f800#10","cell":{"value":45001.25,"format":"[$-F800]h:mm"},"width":40,"predicted":{"text":"Thursday, March 16, 2023"}},
    {"id":"sysdate#11","cell":{"value":44562.75,"format":"[$-x-sysdate]dddd, mmmm dd, yyyy"},"width":40,"predicted":{"text":"Saturday, January 1, 2022"}},
    {"id":"f800#11","cell":{"value":44562.75,"format":"[$-F800]dddd, mmmm dd, yyyy"},"width":40,"predicted":{"text":"Saturday, January 1, 2022"}},
    {"id":"sysdate#12","cell":{"value":44562.75,"format":"[$-x-sysdate]yyyy-mm-dd"},"width":40,"predicted":{"text":"Saturday, January 1, 2022"}},
    {"id":"f800#12","cell":{"value":44562.75,"format":"[$-F800]yyyy-mm-dd"},"width":40,"predicted":{"text":"Saturday, January 1, 2022"}},
    {"id":"sysdate#13","cell":{"value":44562.75,"format":"[$-x-sysdate]d mmm"},"width":40,"predicted":{"text":"Saturday, January 1, 2022"}},
    {"id":"f800#13","cell":{"value":44562.75,"format":"[$-F800]d mmm"},"width":40,"predicted":{"text":"Saturday, January 1, 2022"}},
    {"id":"sysdate#14","cell":{"value":44562.75,"format":"[$-x-sysdate]mmmm"},"width":40,"predicted":{"text":"Saturday, January 1, 2022"}},
    {"id":"f800#14","cell":{"value":44562.75,"format":"[$-F800]mmmm"},"width":40,"predicted":{"text":"Saturday, January 1, 2022"}},
    {"id":"sysdate#15","cell":{"value":44562.75,"format":"[$-x-sysdate]h:mm"},"width":40,"predicted":{"text":"Saturday, January 1, 2022"}},
    {"id":"f800#15","cell":{"value":44562.75,"format":"[$-F800]h:mm"},"width":40,"predicted":{"text":"Saturday, January 1, 2022"}},
    {"id":"systime#1","cell":{"value":45688.5,"format":"[$-x-systime]h:mm:ss AM/PM"},"width":40,"predicted":{"text":"12:00:00 PM"}},
    {"id":"f400#1","cell":{"value":45688.5,"format":"[$-F400]h:mm:ss AM/PM"},"width":40,"predicted":{"text":"12:00:00 PM"}},
    {"id":"systime#2","cell":{"value":45688.5,"format":"[$-x-systime]hh:mm"},"width":40,"predicted":{"text":"12:00:00 PM"}},
    {"id":"f400#2","cell":{"value":45688.5,"format":"[$-F400]hh:mm"},"width":40,"predicted":{"text":"12:00:00 PM"}},
    {"id":"systime#3","cell":{"value":45688.5,"format":"[$-x-systime]h:mm:ss"},"width":40,"predicted":{"text":"12:00:00 PM"}},
    {"id":"f400#3","cell":{"value":45688.5,"format":"[$-F400]h:mm:ss"},"width":40,"predicted":{"text":"12:00:00 PM"}},
    {"id":"systime#4","cell":{"value":45688.5,"format":"[$-x-systime]yyyy-mm-dd"},"width":40,"predicted":{"text":"12:00:00 PM"}},
    {"id":"f400#4","cell":{"value":45688.5,"format":"[$-F400]yyyy-mm-dd"},"width":40,"predicted":{"text":"12:00:00 PM"}},
    {"id":"systime#5","cell":{"value":45001.25,"format":"[$-x-systime]h:mm:ss AM/PM"},"width":40,"predicted":{"text":"6:00:00 AM"}},
    {"id":"f400#5","cell":{"value":45001.25,"format":"[$-F400]h:mm:ss AM/PM"},"width":40,"predicted":{"text":"6:00:00 AM"}},
    {"id":"systime#6","cell":{"value":45001.25,"format":"[$-x-systime]hh:mm"},"width":40,"predicted":{"text":"6:00:00 AM"}},
    {"id":"f400#6","cell":{"value":45001.25,"format":"[$-F400]hh:mm"},"width":40,"predicted":{"text":"6:00:00 AM"}},
    {"id":"systime#7","cell":{"value":45001.25,"format":"[$-x-systime]h:mm:ss"},"width":40,"predicted":{"text":"6:00:00 AM"}},
    {"id":"f400#7","cell":{"value":45001.25,"format":"[$-F400]h:mm:ss"},"width":40,"predicted":{"text":"6:00:00 AM"}},
    {"id":"systime#8","cell":{"value":45001.25,"format":"[$-x-systime]yyyy-mm-dd"},"width":40,"predicted":{"text":"6:00:00 AM"}},
    {"id":"f400#8","cell":{"value":45001.25,"format":"[$-F400]yyyy-mm-dd"},"width":40,"predicted":{"text":"6:00:00 AM"}},
    {"id":"systime#9","cell":{"value":44562.75,"format":"[$-x-systime]h:mm:ss AM/PM"},"width":40,"predicted":{"text":"6:00:00 PM"}},
    {"id":"f400#9","cell":{"value":44562.75,"format":"[$-F400]h:mm:ss AM/PM"},"width":40,"predicted":{"text":"6:00:00 PM"}},
    {"id":"systime#10","cell":{"value":44562.75,"format":"[$-x-systime]hh:mm"},"width":40,"predicted":{"text":"6:00:00 PM"}},
    {"id":"f400#10","cell":{"value":44562.75,"format":"[$-F400]hh:mm"},"width":40,"predicted":{"text":"6:00:00 PM"}},
    {"id":"systime#11","cell":{"value":44562.75,"format":"[$-x-systime]h:mm:ss"},"width":40,"predicted":{"text":"6:00:00 PM"}},
    {"id":"f400#11","cell":{"value":44562.75,"format":"[$-F400]h:mm:ss"},"width":40,"predicted":{"text":"6:00:00 PM"}},
    {"id":"systime#12","cell":{"value":44562.75,"format":"[$-x-systime]yyyy-mm-dd"},"width":40,"predicted":{"text":"6:00:00 PM"}},
    {"id":"f400#12","cell":{"value":44562.75,"format":"[$-F400]yyyy-mm-dd"},"width":40,"predicted":{"text":"6:00:00 PM"}},
    {"id":"currency#1","cell":{"value":1.5,"format":"[$€-x-euro2] #,##0.00"},"width":40,"predicted":{"text":"€ 1.50"}},
    {"id":"currency#2","cell":{"value":1.5,"format":"#,##0.00 [$₽-ru-RU]"},"width":40,"predicted":{"text":"1.50 ₽"}},
    {"id":"currency#3","cell":{"value":1.5,"format":"[$$-en-US]#,##0.00"},"width":40,"predicted":{"text":"$1.50"}},
  ]
}
