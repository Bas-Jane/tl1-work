#!/bin/bash
set -e
cd "$(cd "$(dirname "$0")" && pwd)"
if command -v python3 >/dev/null 2>&1; then
  exec python3 launch.py "$@"
else
  exec python launch.py "$@"
fi
