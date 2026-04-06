#!/bin/zsh
set -eu

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT_DIR"

CONFIG_PATH="${BIJIRIS_CLOUDFLARED_CONFIG:-$ROOT_DIR/cloudflared/config.yml}"

if [ ! -f "$CONFIG_PATH" ]; then
  echo "cloudflared config not found: $CONFIG_PATH" >&2
  echo "Copy cloudflared/config.template.yml to cloudflared/config.yml and set tunnel UUID, credentials-file, hostname." >&2
  exit 1
fi

exec cloudflared tunnel --config "$CONFIG_PATH" run
