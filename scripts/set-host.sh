#!/usr/bin/env bash
# Copyright (c) 2026 Cisco and/or its affiliates.
#
# This software is licensed to you under the terms of the Cisco Sample
# Code License, Version 1.1 (the "License"). You may obtain a copy of the
# License at
#
#                https://developer.cisco.com/docs/licenses
#
# All use of the material herein must be in accordance with the terms of
# the License. All rights not expressly granted by the License are
# reserved. Unless required by applicable law or agreed to separately in
# writing, software distributed under the License is distributed on an "AS
# IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express
# or implied.
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
