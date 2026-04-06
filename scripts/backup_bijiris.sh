#!/bin/zsh
set -eu

ROOT_DIR="$(cd "$(dirname "$0")/.." && pwd)"
cd "$ROOT_DIR"

mkdir -p data/backups
STAMP="$(date +%Y%m%d-%H%M%S)"
BACKUP_PATH="data/backups/bijiris-backup-${STAMP}.tar.gz"

tar -czf "$BACKUP_PATH" data/bijiris.db data/config.json data/uploads
echo "$ROOT_DIR/$BACKUP_PATH"
