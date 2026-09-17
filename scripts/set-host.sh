#!/usr/bin/env bash
# Replace the placeholder host in src/manifest.xml with the real HTTPS host.
#
#   scripts/set-host.sh addin.firma.dk            # in place
#   scripts/set-host.sh addin.firma.dk out.xml    # write to another file
#
# Only the host part is replaced; paths (/taskpane.html, /assets/...) stay as they are.
set -euo pipefail

if [[ $# -lt 1 ]]; then
  echo "usage: $0 <host[:port]> [output-file]" >&2
  exit 2
fi

host="$1"
placeholder="addin.example.com"
src="$(cd "$(dirname "$0")/.." && pwd)/src/manifest.xml"
out="${2:-$src}"

if ! grep -q "$placeholder" "$src"; then
  echo "warning: placeholder '$placeholder' not found in $src – nothing to do" >&2
fi

sed "s#https://${placeholder}#https://${host}#g" "$src" > "${out}.tmp"
mv "${out}.tmp" "$out"
echo "wrote $out with host https://${host} ($(grep -c "https://${host}" "$out") occurrences)"
