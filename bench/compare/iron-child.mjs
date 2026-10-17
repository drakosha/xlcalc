// One workload of bench/hyperformula in IronCalc or xlcalc, by that benchmark's child and protocol: node iron-child.mjs <workload> <ironcalc|xl-closures|xl-eval>
import './iron.mjs'
await import('../hyperformula/child.mjs')
