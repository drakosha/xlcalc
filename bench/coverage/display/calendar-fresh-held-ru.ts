/**
 * Held-out display cases of the calendar prefixes on other dates (Excel ru-RU, `display calendar-fresh-held-ru ru`), each answer predicted by the
 * engine after the rules the sets `calendar-held`, `calendar-names-held` and `calendar-hijri-held` gave, and committed
 * before Excel showed the case: `B1`'s English names and `B2`'s tabular Hijri dates with their Arabic names, from 1924 to
 * 2100, with times.
 */
import type { DisplayCase } from './cases.js'

export function buildDisplayCases(): DisplayCase[] {
  return [
    {"id":"fresh#1","cell":{"value":9132,"format":"B2d mmmm yyyy"},"width":40,"predicted":{"text":"5 جمادى الثانية 1343"}},
    {"id":"fresh#2","cell":{"value":9132,"format":"B2dddd"},"width":40,"predicted":{"text":"الأربعاء"}},
    {"id":"fresh#3","cell":{"value":9132,"format":"B2mmm yy"},"width":40,"predicted":{"text":"جمادى الثانية 43"}},
    {"id":"fresh#4","cell":{"value":9132,"format":"B2ddd dd/mm h:mm"},"width":40,"predicted":{"text":"الأربعاء 05.06 0:00"}},
    {"id":"fresh#5","cell":{"value":9132,"format":"B1dddd d mmmm yyyy"},"width":40,"predicted":{"text":"Wednesday 31 December 1924"}},
    {"id":"fresh#6","cell":{"value":9132,"format":"b1mmm d, yyyy"},"width":40,"predicted":{"text":"Dec 31, 1924"}},
    {"id":"fresh#7","cell":{"value":18429.770833333332,"format":"B2d mmmm yyyy"},"width":40,"predicted":{"text":"29 شعبان 1369"}},
    {"id":"fresh#8","cell":{"value":18429.770833333332,"format":"B2dddd"},"width":40,"predicted":{"text":"الخميس"}},
    {"id":"fresh#9","cell":{"value":18429.770833333332,"format":"B2mmm yy"},"width":40,"predicted":{"text":"شعبان 69"}},
    {"id":"fresh#10","cell":{"value":18429.770833333332,"format":"B2ddd dd/mm h:mm"},"width":40,"predicted":{"text":"الخميس 29.08 18:30"}},
    {"id":"fresh#11","cell":{"value":18429.770833333332,"format":"B1dddd d mmmm yyyy"},"width":40,"predicted":{"text":"Thursday 15 June 1950"}},
    {"id":"fresh#12","cell":{"value":18429.770833333332,"format":"b1mmm d, yyyy"},"width":40,"predicted":{"text":"Jun 15, 1950"}},
    {"id":"fresh#13","cell":{"value":32085,"format":"B2d mmmm yyyy"},"width":40,"predicted":{"text":"13 ربيع الأول 1408"}},
    {"id":"fresh#14","cell":{"value":32085,"format":"B2dddd"},"width":40,"predicted":{"text":"الأربعاء"}},
    {"id":"fresh#15","cell":{"value":32085,"format":"B2mmm yy"},"width":40,"predicted":{"text":"ربيع الأول 08"}},
    {"id":"fresh#16","cell":{"value":32085,"format":"B2ddd dd/mm h:mm"},"width":40,"predicted":{"text":"الأربعاء 13.03 0:00"}},
    {"id":"fresh#17","cell":{"value":32085,"format":"B1dddd d mmmm yyyy"},"width":40,"predicted":{"text":"Wednesday 4 November 1987"}},
    {"id":"fresh#18","cell":{"value":32085,"format":"b1mmm d, yyyy"},"width":40,"predicted":{"text":"Nov 4, 1987"}},
    {"id":"fresh#19","cell":{"value":54484,"format":"B2d mmmm yyyy"},"width":40,"predicted":{"text":"28 جمادى الأولى 1471"}},
    {"id":"fresh#20","cell":{"value":54484,"format":"B2dddd"},"width":40,"predicted":{"text":"الثلاثاء"}},
    {"id":"fresh#21","cell":{"value":54484,"format":"B2mmm yy"},"width":40,"predicted":{"text":"جمادى الأولى 71"}},
    {"id":"fresh#22","cell":{"value":54484,"format":"B2ddd dd/mm h:mm"},"width":40,"predicted":{"text":"الثلاثاء 28.05 0:00"}},
    {"id":"fresh#23","cell":{"value":54484,"format":"B1dddd d mmmm yyyy"},"width":40,"predicted":{"text":"Tuesday 2 March 2049"}},
    {"id":"fresh#24","cell":{"value":54484,"format":"b1mmm d, yyyy"},"width":40,"predicted":{"text":"Mar 2, 2049"}},
    {"id":"fresh#25","cell":{"value":64151.26041666667,"format":"B2d mmmm yyyy"},"width":40,"predicted":{"text":"9 رمضان 1498"}},
    {"id":"fresh#26","cell":{"value":64151.26041666667,"format":"B2dddd"},"width":40,"predicted":{"text":"الثلاثاء"}},
    {"id":"fresh#27","cell":{"value":64151.26041666667,"format":"B2mmm yy"},"width":40,"predicted":{"text":"رمضان 98"}},
    {"id":"fresh#28","cell":{"value":64151.26041666667,"format":"B2ddd dd/mm h:mm"},"width":40,"predicted":{"text":"الثلاثاء 09.09 6:15"}},
    {"id":"fresh#29","cell":{"value":64151.26041666667,"format":"B1dddd d mmmm yyyy"},"width":40,"predicted":{"text":"Tuesday 20 August 2075"}},
    {"id":"fresh#30","cell":{"value":64151.26041666667,"format":"b1mmm d, yyyy"},"width":40,"predicted":{"text":"Aug 20, 2075"}},
    {"id":"fresh#31","cell":{"value":73110,"format":"B2d mmmm yyyy"},"width":40,"predicted":{"text":"20 ذو الحجة 1523"}},
    {"id":"fresh#32","cell":{"value":73110,"format":"B2dddd"},"width":40,"predicted":{"text":"الإثنين"}},
    {"id":"fresh#33","cell":{"value":73110,"format":"B2mmm yy"},"width":40,"predicted":{"text":"ذو الحجة 23"}},
    {"id":"fresh#34","cell":{"value":73110,"format":"B2ddd dd/mm h:mm"},"width":40,"predicted":{"text":"الإثنين 20.12 0:00"}},
    {"id":"fresh#35","cell":{"value":73110,"format":"B1dddd d mmmm yyyy"},"width":40,"predicted":{"text":"Monday 1 March 2100"}},
    {"id":"fresh#36","cell":{"value":73110,"format":"b1mmm d, yyyy"},"width":40,"predicted":{"text":"Mar 1, 2100"}},
  ]
}
