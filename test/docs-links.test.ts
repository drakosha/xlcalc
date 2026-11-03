/**
 * The links of the README and of `docs/` lead somewhere: every relative link names a file that exists, and an anchor
 * (`#...`) one of its headings, as GitHub makes them (lower case, punctuation dropped, spaces as `-`, a repeated heading
 * numbered `-1`, `-2`).
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { dirname, join, relative, resolve } from 'node:path'
import { describe, expect, it } from 'vitest'

const root = resolve(import.meta.dirname, '..')
const files = ['README.md', ...readdirSync(join(root, 'docs')).filter((f) => f.endsWith('.md')).map((f) => `docs/${f}`)]

/** A heading's anchor as GitHub writes it. */
const slug = (heading: string) =>
  heading
    .trim()
    .toLowerCase()
    .replace(/[^\p{L}\p{N}\p{M}\s_-]/gu, '')
    .replace(/\s/g, '-')

/**
 * A Markdown file's lines outside fenced code blocks, as CommonMark reads fences: a block closes only at a line of its
 * fence's character alone, at least as long, so ``` followed by text on its line leaves the block open to the end.
 */
function outsideCode(text: string): string {
  const out: string[] = []
  let fence: string | undefined
  for (const line of text.split('\n')) {
    const m = /^ {0,3}(`{3,}|~{3,})(.*)$/.exec(line)
    if (fence === undefined) {
      if (m && !(m[1][0] === '`' && m[2].includes('`'))) fence = m[1]
      else out.push(line)
    } else if (m && m[1][0] === fence[0] && m[1].length >= fence.length && !m[2].trim()) fence = undefined
  }
  return out.join('\n')
}

/** The anchors of a Markdown file's headings, outside code blocks. */
function anchors(text: string): Set<string> {
  const out = new Set<string>()
  const seen = new Map<string, number>()
  for (const m of outsideCode(text).matchAll(/^#{1,6} (.+)$/gm)) {
    const base = slug(m[1])
    const n = seen.get(base) ?? 0
    seen.set(base, n + 1)
    out.add(n ? `${base}-${n}` : base)
  }
  for (const m of text.matchAll(/<a id="([^"]+)"><\/a>/g)) out.add(m[1])
  return out
}

const texts = new Map(files.map((f) => [f, readFileSync(join(root, f), 'utf8')]))

describe('links of the README and docs/', () => {
  it('lead to files and headings that exist', () => {
    const broken: string[] = []
    for (const [file, text] of texts)
      for (const m of text.replace(/```[\s\S]*?```/g, '').matchAll(/\]\(([^)\s]+)\)/g)) {
        const link = m[1]
        if (/^[a-z]+:/i.test(link)) continue
        const [path, anchor] = link.split('#')
        const target = path ? relative(root, resolve(join(root, dirname(file)), path)) : file
        if (!existsSync(join(root, target))) broken.push(`${file}: ${link} (no file)`)
        else if (anchor !== undefined && target.endsWith('.md')) {
          const own = texts.get(target) ?? readFileSync(join(root, target), 'utf8')
          if (!anchors(own).has(anchor)) broken.push(`${file}: ${link} (no heading)`)
        }
      }
    expect(broken).toEqual([])
  })
})
