#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sqlite3
from pathlib import Path
from typing import Any


ROOT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_DB_PATH = ROOT_DIR / "data" / "bijiris.db"
DEFAULT_CONFIG_PATH = ROOT_DIR / "data" / "config.json"
DEFAULT_UPLOADS_DIR = ROOT_DIR / "data" / "uploads"


def read_json(path: Path) -> dict[str, Any]:
    if not path.exists():
      return {}
    return json.loads(path.read_text(encoding="utf-8"))


def rows_to_dicts(rows: list[sqlite3.Row]) -> list[dict[str, Any]]:
    return [dict(row) for row in rows]


def decode_json(value: Any, fallback: Any) -> Any:
    raw = str(value or "").strip()
    if not raw:
        return fallback
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return fallback


def next_id(items: list[dict[str, Any]], key: str = "id") -> int:
    values = [int(item.get(key) or 0) for item in items]
    return (max(values) if values else 0) + 1


def normalize_field(row: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": int(row["id"]),
        "label": row["label"],
        "key": row["field_key"],
        "type": row["type"],
        "required": bool(row["required"]),
        "options": decode_json(row.get("options"), []),
        "placeholder": row.get("placeholder") or "",
        "helpText": row.get("help_text") or "",
        "visibilityFieldKey": row.get("visibility_field_key") or "",
        "visibilityValues": decode_json(row.get("visibility_values"), []),
        "accept": row.get("accept") or "",
        "allowMultiple": bool(row.get("allow_multiple")),
        "allowOther": bool(row.get("allow_other")),
        "sortOrder": int(row.get("sort_order") or 0),
    }


def response_file_payload(
    row: dict[str, Any],
    annotation: dict[str, Any] | None,
    uploads_dir: Path,
) -> dict[str, Any]:
    stored_name = str(row.get("stored_name") or "").strip()
    relative_path = str(row.get("relative_path") or "").strip()
    local_path = uploads_dir / stored_name if stored_name else None
    image_url = relative_path
    preview_url = relative_path
    if relative_path.startswith("http"):
        image_url = relative_path
        preview_url = relative_path

    return {
        "id": int(row["id"]),
        "fieldKey": row.get("field_key") or "",
        "label": row.get("label") or "",
        "originalName": row.get("original_name") or "",
        "storedName": stored_name,
        "mimeType": row.get("mime_type") or "",
        "size": int(row.get("size") or 0),
        "relativePath": relative_path,
        "url": image_url,
        "previewUrl": preview_url,
        "annotation": {
            "title": (annotation or {}).get("title") or "",
            "entryDate": (annotation or {}).get("entry_date") or "",
            "memo": (annotation or {}).get("memo") or "",
            "createdAt": (annotation or {}).get("created_at") or row.get("response_created_at") or "",
            "updatedAt": (annotation or {}).get("updated_at") or row.get("response_created_at") or "",
        },
        "localFile": {
            "exists": bool(local_path and local_path.exists()),
            "path": str(local_path) if local_path and local_path.exists() else "",
        },
    }


def profile_record_payload(row: dict[str, Any], uploads_dir: Path) -> dict[str, Any]:
    stored_name = str(row.get("image_stored_name") or "").strip()
    relative_path = str(row.get("image_relative_path") or "").strip()
    local_path = uploads_dir / stored_name if stored_name else None
    image = None
    if relative_path or stored_name:
        image = {
            "originalName": row.get("image_original_name") or "",
            "storedName": stored_name,
            "mimeType": row.get("image_mime_type") or "",
            "relativePath": relative_path,
            "url": relative_path,
            "previewUrl": relative_path,
            "localFile": {
                "exists": bool(local_path and local_path.exists()),
                "path": str(local_path) if local_path and local_path.exists() else "",
            },
        }
    return {
        "id": int(row["id"]),
        "recordId": int(row["id"]),
        "sourceType": "manual",
        "respondentId": row["respondent_id"],
        "respondentName": row.get("respondent_name") or "",
        "title": row.get("title") or "",
        "date": row.get("entry_date") or "",
        "memo": row.get("memo") or "",
        "createdAt": row.get("created_at") or "",
        "updatedAt": row.get("updated_at") or "",
        "editable": True,
        "deletable": True,
        "sourceLabel": "管理者追加",
        "image": image,
    }


def measurement_payload(row: dict[str, Any]) -> dict[str, Any]:
    def label(value: Any) -> str:
        amount = float(value)
        return str(int(amount)) if amount.is_integer() else f"{amount:.1f}"

    return {
        "id": int(row["id"]),
        "recordId": int(row["id"]),
        "respondentId": row["respondent_id"],
        "respondentName": row.get("respondent_name") or "",
        "date": row.get("entry_date") or "",
        "category": row.get("measurement_category") or "",
        "waist": float(row["waist"]),
        "hip": float(row["hip"]),
        "thigh": float(row["thigh"]),
        "waistLabel": label(row["waist"]),
        "hipLabel": label(row["hip"]),
        "thighLabel": label(row["thigh"]),
        "createdAt": row.get("created_at") or "",
        "updatedAt": row.get("updated_at") or "",
        "editable": True,
        "deletable": True,
    }


def export_data(db_path: Path, config_path: Path, uploads_dir: Path) -> dict[str, Any]:
    config = read_json(config_path)
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        forms_rows = rows_to_dicts(conn.execute("SELECT * FROM forms ORDER BY updated_at DESC, id DESC").fetchall())
        field_rows = rows_to_dicts(
            conn.execute("SELECT * FROM fields ORDER BY form_id ASC, sort_order ASC, id ASC").fetchall()
        )
        respondents_rows = rows_to_dicts(
            conn.execute("SELECT * FROM respondents ORDER BY respondent_name COLLATE NOCASE ASC, respondent_id ASC").fetchall()
        )
        responses_rows = rows_to_dicts(
            conn.execute("SELECT * FROM responses ORDER BY created_at ASC, id ASC").fetchall()
        )
        answer_rows = rows_to_dicts(
            conn.execute("SELECT * FROM response_answers ORDER BY response_id ASC, id ASC").fetchall()
        )
        file_rows = rows_to_dicts(
            conn.execute("SELECT * FROM response_files ORDER BY response_id ASC, id ASC").fetchall()
        )
        annotation_rows = rows_to_dicts(
            conn.execute("SELECT * FROM response_file_annotations ORDER BY response_file_id ASC").fetchall()
        )
        profile_rows = rows_to_dicts(
            conn.execute("SELECT * FROM respondent_profile_records ORDER BY entry_date DESC, updated_at DESC, id DESC").fetchall()
        )
        measurement_rows = rows_to_dicts(
            conn.execute(
                "SELECT * FROM respondent_measurements ORDER BY entry_date ASC, updated_at ASC, id ASC"
            ).fetchall()
        )
    finally:
        conn.close()

    fields_by_form: dict[int, list[dict[str, Any]]] = {}
    for row in field_rows:
        fields_by_form.setdefault(int(row["form_id"]), []).append(normalize_field(row))

    forms = []
    for row in forms_rows:
        forms.append(
            {
                "id": int(row["id"]),
                "title": row["title"],
                "slug": row["slug"],
                "description": row.get("description") or "",
                "successMessage": row.get("success_message") or "",
                "categoryLabel": row.get("category_label") or "分類",
                "categoryOptions": decode_json(row.get("category_options"), []),
                "isActive": bool(row.get("is_active")),
                "createdAt": row.get("created_at") or "",
                "updatedAt": row.get("updated_at") or "",
                "fields": fields_by_form.get(int(row["id"]), []),
            }
        )

    answers_by_response: dict[int, list[dict[str, Any]]] = {}
    for row in answer_rows:
        answers_by_response.setdefault(int(row["response_id"]), []).append(
            {
                "id": int(row["id"]),
                "fieldKey": row.get("field_key") or "",
                "label": row.get("label") or "",
                "value": row.get("value") or "",
            }
        )

    annotations_by_file_id = {int(row["response_file_id"]): row for row in annotation_rows}
    files_by_response: dict[int, list[dict[str, Any]]] = {}
    for row in file_rows:
        files_by_response.setdefault(int(row["response_id"]), []).append(
            response_file_payload(row, annotations_by_file_id.get(int(row["id"])), uploads_dir)
        )

    form_titles = {int(form["id"]): form["title"] for form in forms}
    responses = []
    for row in responses_rows:
        response_id = int(row["id"])
        responses.append(
            {
                "id": response_id,
                "formId": int(row["form_id"]),
                "formTitle": form_titles.get(int(row["form_id"]), ""),
                "respondentId": row["respondent_id"],
                "respondentName": row.get("respondent_name") or "",
                "respondentEmail": row.get("respondent_email") or "",
                "category": row.get("category") or "",
                "notes": row.get("notes") or "",
                "createdAt": row.get("created_at") or "",
                "ipAddress": row.get("ip_address") or "",
                "userAgent": row.get("user_agent") or "",
                "answers": answers_by_response.get(response_id, []),
                "files": files_by_response.get(response_id, []),
            }
        )

    respondents = [
        {
            "respondentId": row["respondent_id"],
            "respondentName": row.get("respondent_name") or "",
            "createdAt": row.get("created_at") or "",
            "updatedAt": row.get("updated_at") or "",
            "ticketSheetManualValue": row.get("ticket_sheet_manual_value") or "",
            "currentTicketBookType": row.get("current_ticket_book_type") or "",
            "currentTicketStampCount": int(row.get("current_ticket_stamp_count") or 0),
            "currentTicketStampManualEnabled": bool(row.get("current_ticket_stamp_manual_enabled")),
        }
        for row in respondents_rows
    ]

    profile_records = [profile_record_payload(row, uploads_dir) for row in profile_rows]
    measurements = [measurement_payload(row) for row in measurement_rows]

    return {
        "meta": {
            "schemaVersion": 1,
            "exportedAt": __import__("datetime").datetime.now(__import__("datetime").timezone.utc).replace(microsecond=0).isoformat(),
        },
        "settings": {
            "adminUsername": config.get("admin_username") or "admin",
            "adminPasswordSha256": config.get("admin_password_sha256") or "",
            "publicBaseUrl": config.get("public_base_url") or "",
        },
        "counters": {
            "nextFormId": next_id(forms),
            "nextFieldId": next_id(field_rows),
            "nextResponseId": next_id(responses),
            "nextResponseFileId": next_id(file_rows),
            "nextProfileRecordId": next_id(profile_rows),
            "nextMeasurementId": next_id(measurement_rows),
        },
        "forms": forms,
        "respondents": respondents,
        "responses": responses,
        "profileRecords": profile_records,
        "measurements": measurements,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Export Bijiris SQLite data to GAS-friendly JSON.")
    parser.add_argument("--db", type=Path, default=DEFAULT_DB_PATH)
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG_PATH)
    parser.add_argument("--uploads", type=Path, default=DEFAULT_UPLOADS_DIR)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    payload = export_data(args.db, args.config, args.uploads)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
