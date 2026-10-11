#!/bin/sh
# The fixtures of test/viewport-saved.test.ts: book.json computed once by xlcalc 0.11.0 and 0.16.3 (installed from the
# release tarballs into a temporary directory, never by the test) and by this tree's code. Run from anywhere.
set -e
HERE=$(cd "$(dirname "$0")" && pwd)
ROOT=$(cd "$HERE/../../.." && pwd)
RELEASES=${XLCALC_RELEASES:-$HOME/Developer/formula/release}
TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT
for v in 0.11.0 0.16.3; do
  mkdir -p "$TMP/$v"
  (cd "$TMP/$v" && echo '{"name":"fixture-'$v'","private":true}' > package.json && npm install --no-audit --no-fund "$RELEASES/xlcalc-$v.tgz" > /dev/null)
  node "$HERE/compute.mjs" "$TMP/$v/node_modules/xlcalc/dist/index.js" "$HERE/result-$v.json"
done
v=$(node -p "require('$ROOT/package.json').version")
(cd "$ROOT" && npx tsx "$HERE/compute.mjs" "$ROOT/src/index.ts" "$HERE/result-$v.json")
