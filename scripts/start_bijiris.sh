#!/bin/zsh
set -eu

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT_DIR"

PUBLIC_URL="$(python3 - <<'PY'
import json
from pathlib import Path

path = Path("data/config.json")
if not path.exists():
    print("")
else:
    data = json.loads(path.read_text(encoding="utf-8"))
    print(str(data.get("public_base_url", "")).strip())
PY
)"

if [ -n "$PUBLIC_URL" ]; then
  export BIJIRIS_PUBLIC_BASE_URL="$PUBLIC_URL"
fi

exec python3 server.py --host 0.0.0.0 --port 8123
