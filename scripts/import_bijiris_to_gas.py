#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import json
import mimetypes
import sys
import urllib.error
import urllib.request
from pathlib import Path

from export_bijiris_for_gas import export_data, DEFAULT_CONFIG_PATH, DEFAULT_DB_PATH, DEFAULT_UPLOADS_DIR


def post_json(url: str, payload: dict) -> dict:
    request = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "text/plain;charset=utf-8"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=120) as response:
        raw = response.read().decode("utf-8")
    return json.loads(raw)


def upload_asset(url: str, password: str, kind: str, metadata: dict, file_path: Path) -> dict:
    mime_type = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
    payload = {
        "action": "uploadLocalAsset",
        "password": password,
        "kind": kind,
        "file": {
            "fieldName": metadata.get("fieldName", ""),
            "name": file_path.name,
            "type": mime_type,
            "size": file_path.stat().st_size,
            "base64": base64.b64encode(file_path.read_bytes()).decode("ascii"),
        },
    }
    payload.update(metadata)
    return post_json(url, payload)


def main() -> None:
    parser = argparse.ArgumentParser(description="Import local Bijiris data into GAS storage.")
    parser.add_argument("--gas-url", required=True)
    parser.add_argument("--password", default="bijiris-admin")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB_PATH)
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG_PATH)
    parser.add_argument("--uploads", type=Path, default=DEFAULT_UPLOADS_DIR)
    args = parser.parse_args()

    export_payload = export_data(args.db, args.config, args.uploads)
    initialized = post_json(
        args.gas_url,
        {
            "action": "initializeData",
            "password": args.password,
            "data": export_payload,
        },
    )
    if int(initialized.get("statusCode", 500)) >= 300:
      raise SystemExit(initialized.get("error", "GAS 初期化に失敗しました。"))

    uploaded = 0
    for response in export_payload.get("responses", []):
        for file_item in response.get("files", []):
            local_file = Path(str(((file_item.get("localFile") or {}).get("path") or "")).strip())
            if not local_file.exists():
                continue
            result = upload_asset(
                args.gas_url,
                args.password,
                "responseFile",
                {
                    "responseId": response["id"],
                    "fileId": file_item["id"],
                    "fieldName": file_item.get("fieldKey", ""),
                },
                local_file,
            )
            if int(result.get("statusCode", 500)) >= 300:
                raise SystemExit(result.get("error", f"responseFile upload failed: {local_file.name}"))
            uploaded += 1

    for record in export_payload.get("profileRecords", []):
        image = record.get("image") or {}
        local_file = Path(str(((image.get("localFile") or {}).get("path") or "")).strip())
        if not local_file.exists():
            continue
        result = upload_asset(
            args.gas_url,
            args.password,
            "profileRecord",
            {
                "respondentId": record["respondentId"],
                "recordId": record["id"],
            },
            local_file,
        )
        if int(result.get("statusCode", 500)) >= 300:
            raise SystemExit(result.get("error", f"profileRecord upload failed: {local_file.name}"))
        uploaded += 1

    print(
        json.dumps(
            {
                "initialized": initialized.get("data") or {},
                "uploadedLocalImages": uploaded,
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    try:
        main()
    except urllib.error.HTTPError as error:
        body = error.read().decode("utf-8", errors="replace")
        print(body or str(error), file=sys.stderr)
        raise
