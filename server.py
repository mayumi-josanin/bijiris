#!/usr/bin/env python3
from __future__ import annotations

import argparse
import base64
import csv
import datetime as dt
import hashlib
import hmac
import io
import json
import mimetypes
import os
import re
import secrets
import sqlite3
import socket
import subprocess
import sys
import tarfile
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import warnings
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from itertools import zip_longest

warnings.filterwarnings(
    "ignore",
    message="'cgi' is deprecated and slated for removal in Python 3.13",
    category=DeprecationWarning,
)

import cgi

ROOT_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("BIJIRIS_DATA_DIR", str(ROOT_DIR / "data"))).expanduser()
UPLOADS_DIR = DATA_DIR / "uploads"
BACKUPS_DIR = DATA_DIR / "backups"
PUBLIC_DIR = ROOT_DIR / "public"
DB_PATH = DATA_DIR / "bijiris.db"
CONFIG_PATH = DATA_DIR / "config.json"
SESSION_COOKIE = "bijiris_session"
DEFAULT_PASSWORD = "bijiris-admin"
DEFAULT_SERVER_PORT = 8123
DEFAULT_SERVER_HOST = "0.0.0.0"
FILE_SIZE_LIMIT = 10 * 1024 * 1024
TOTAL_UPLOAD_LIMIT = 30 * 1024 * 1024
ALLOWED_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".heic", ".heif"}
ALLOWED_FIELD_TYPES = {"short_text", "long_text", "select", "radio", "checkbox", "file"}
MEASUREMENT_CATEGORIES = (
    "モニター",
    "回数券",
    "トライアル",
    "単発",
    "初回お試し",
    "乗り放題キャンペーン",
    "その他",
)
TICKET_END_FORM_SLUG = "bijiris-ticket-end"
TICKET_SHEET_FIELD_KEY = "ticket_sheet_number"
TICKET_SHEET_FIELD_LABEL = "今回終了した回数券"
TICKET_BOOK_TYPES = ("6回券", "10回券")
TREATMENT_SURVEY_FORM_SLUG = "bijiris-treatment-survey"
TICKET_VISIT_COUNT_FIELD_KEY = "ticket_visit_count"
MEASUREMENT_IMPORT_COLUMN_HINTS: dict[str, tuple[str, ...]] = {
    "name": ("お名前", "名前", "氏名", "顧客名", "患者名", "利用者名", "会員名", "name"),
    "date": (
        "計測日",
        "測定日",
        "計測年月日",
        "測定年月日",
        "日付",
        "年月日",
        "日にち",
        "来院日",
        "実施日",
        "施術日",
        "計測日時",
        "測定日時",
        "date",
        "timestamp",
    ),
    "waist": ("ウエスト", "腹囲", "waist"),
    "hip": ("ヒップ", "臀囲", "hip"),
    "thigh": ("太もも", "太腿", "大腿", "もも", "thigh"),
}
MEASUREMENT_SITE_COLUMN_HINTS: tuple[str, ...] = ("計測箇所", "測定箇所", "計測部位", "測定部位", "部位", "項目")
PUBLIC_BASE_URL_ENV = "BIJIRIS_PUBLIC_BASE_URL"
RENDER_EXTERNAL_URL_ENV = "RENDER_EXTERNAL_URL"
REFERENCE_FORM_DEFINITIONS: list[dict[str, Any]] = [
    {
        "title": "ビジリスモニター終了アンケート",
        "slug": "bijiris-monitor-end",
        "description": (
            "6回のモニター参加ありがとうございました。体の変化や率直な感想をお聞かせください。"
            " 写真や掲載可否、継続プランもこのフォームで受け付けます。"
        ),
        "success_message": "モニター終了アンケートを受け付けました。ありがとうございました。",
        "category_label": "",
        "category_options": [],
        "fields": [
            {
                "label": "ビジリスの刺激や使用感はいかがでしたか？",
                "field_key": "stimulus_feeling",
                "type": "radio",
                "required": 1,
                "options": [
                    "痛気持ちよくて効いている感じがした",
                    "最初は驚いたが慣れた",
                    "リラックスして座っていられた",
                ],
                "placeholder": "",
                "help_text": "",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 0,
            },
            {
                "label": "ビジリスを6回体験して、日常生活で変化を感じたことはありますか？（複数回答可）",
                "field_key": "daily_changes",
                "type": "checkbox",
                "required": 0,
                "options": [
                    "立っている時や座っている時の姿勢が楽になった",
                    "長時間歩いても疲れにくくなった",
                    "尿漏れや頻尿が気にならなくなった",
                    "ズボンやスカートが緩くなった気がする",
                    "冷え性が良くなった（体がポカポカする）",
                    "便通が良くなった",
                    "腰痛・股関節痛が軽くなった",
                    "睡眠の質が良くなった",
                    "特に変化は感じなかった",
                    "階段の上り下りが楽になった",
                ],
                "placeholder": "",
                "help_text": "",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 1,
                "allow_other": 1,
                "sort_order": 1,
            },
            {
                "label": "今後もっと改善したい部分はありますか？",
                "field_key": "future_improvements",
                "type": "checkbox",
                "required": 0,
                "options": [
                    "もっとお腹周りを引き締めたい",
                    "痛みのない生活を送りたい",
                    "姿勢をもっと良くしたい",
                    "今の良い状態をキープしたい",
                    "睡眠の質を高めたい",
                    "妊娠しやすい体づくりをしたい",
                    "トイレトラブルを改善したい",
                ],
                "placeholder": "",
                "help_text": "",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 1,
                "allow_other": 1,
                "sort_order": 2,
            },
            {
                "label": "計測写真(1回目)",
                "field_key": "first_measurement_images",
                "type": "file",
                "required": 1,
                "options": [],
                "placeholder": "",
                "help_text": "1回目計測時の写真2枚(後ろ向き・横向き)をご提出ください",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "image/*",
                "allow_multiple": 1,
                "sort_order": 3,
            },
            {
                "label": "計測写真(6回目)",
                "field_key": "sixth_measurement_images",
                "type": "file",
                "required": 1,
                "options": [],
                "placeholder": "",
                "help_text": "6回目計測時の写真2枚(後ろ向き・横向き)をご提出ください",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "image/*",
                "allow_multiple": 1,
                "sort_order": 4,
            },
            {
                "label": "今回の測定結果や感想を、個人が特定できない形でSNSや院内掲示に掲載させていただいてもよろしいでしょうか？",
                "field_key": "sns_permission",
                "type": "radio",
                "required": 1,
                "options": ["写真・数値・感想すべてOK", "掲載不可"],
                "placeholder": "",
                "help_text": (
                    "※ お名前はイニシャルや仮名を使用します。お写真は顔やお部屋の背景などを隠し、"
                    "プライバシーに最大限配慮いたします。同じ悩みを持つ方への励みになりますので、"
                    "ご協力いただければ幸いです。\n"
                    "※ ご承諾いただけた方は、回数券をモニター特別価格よりもさらに特別割引でご購入いただけます"
                ),
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 5,
            },
            {
                "label": "(任意)あなたの身体の変化を「見える化」！「ビフォーアフター写真＆個別アドバイスシート」をプレゼント",
                "field_key": "advice_sheet",
                "type": "radio",
                "required": 0,
                "options": ["希望する", "希望しない"],
                "placeholder": "",
                "help_text": "",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 6,
            },
            {
                "label": "今後はどうしていきたいですか？",
                "field_key": "future_plan",
                "type": "radio",
                "required": 1,
                "options": ["6回の効果を無駄にしないために継続していきたい", "継続は考えていない"],
                "placeholder": "",
                "help_text": "上記のモニター特典を参考にしてみてください",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 7,
            },
            {
                "label": "継続していきたい方へ\n＊今後の継続プラン(回数券)について、ご希望のコースをお選びください",
                "field_key": "continuation_course",
                "type": "radio",
                "required": 0,
                "options": [
                    "SNS掲載御礼！特別感謝コース（30%OFF）",
                    "モニター様優待コース（10%OFF）",
                    "単発コース(都度払いコース) ＊割引がありません",
                ],
                "placeholder": "",
                "help_text": (
                    "◯特別割引に関しまして\n"
                    "使用されている回数券の有効期限が切れるまでに新たな回数券を購入する場合に規定の割引が適用されます\n\n"
                    "有効期限が切れて期間が空いてからの購入の場合の割引は以下になります↓↓↓\n"
                    "・SNS掲載御礼！特別感謝コース（30%OFF）→SNS掲載御礼！特別感謝コース（10%OFF）\n"
                    "⚠️継続していきたい方は回数券の有効期限が切れる前に購入することを忘れずに！"
                ),
                "visibility_field_key": "future_plan",
                "visibility_values": ["6回の効果を無駄にしないために継続していきたい"],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 8,
            },
            {
                "label": "ご質問・ご相談(自由記述)",
                "field_key": "questions",
                "type": "long_text",
                "required": 0,
                "options": [],
                "placeholder": "",
                "help_text": "ご質問やご相談等があればご記載ください　＊内容はなんでも構いません！",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 9,
            },
        ],
    },
    {
        "title": "ビジリス回数券終了アンケート",
        "slug": "bijiris-ticket-end",
        "description": "回数券利用後の体感や変化、計測写真、今後の改善希望をお聞かせください。",
        "success_message": "回数券終了アンケートを受け付けました。ありがとうございました。",
        "category_label": "",
        "category_options": [],
        "fields": [
            {
                "label": "本日のビジリスの体感はいかがでしたか？",
                "field_key": "body_feedback",
                "type": "long_text",
                "required": 0,
                "options": [],
                "placeholder": "以前と比べて変化したことなどがあればご記載ください",
                "help_text": "以前と比べて変化したことなどがあればご記載ください",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 0,
            },
            {
                "label": "日常生活で変化を感じたことはありますか？（複数回答可）",
                "field_key": "daily_changes",
                "type": "checkbox",
                "required": 0,
                "options": [
                    "立っている時や座っている時の姿勢が楽になった",
                    "長時間歩いても疲れにくくなった",
                    "尿漏れや頻尿が気にならなくなった",
                    "ズボンやスカートが緩くなった気がする",
                    "冷え性が良くなった（体がポカポカする）",
                    "便通が良くなった",
                    "腰痛・股関節痛が軽くなった",
                    "睡眠の質が良くなった",
                    "階段の上り下りが楽になった",
                    "特に変化は感じなかった",
                ],
                "placeholder": "",
                "help_text": "",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 1,
                "allow_other": 1,
                "sort_order": 1,
            },
            {
                "label": "今後もっと改善したい部分はありますか？",
                "field_key": "future_improvements",
                "type": "checkbox",
                "required": 0,
                "options": [
                    "もっとお腹周りを引き締めたい",
                    "痛みのない生活を送りたい",
                    "姿勢をもっと良くしたい",
                    "今の良い状態をキープしたい",
                    "睡眠の質を高めたい",
                    "妊娠しやすい体づくりをしたい",
                    "トイレトラブルを改善したい",
                ],
                "placeholder": "",
                "help_text": "",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 1,
                "allow_other": 1,
                "sort_order": 2,
            },
            {
                "label": "計測写真(1回目)",
                "field_key": "first_measurement_images",
                "type": "file",
                "required": 0,
                "options": [],
                "placeholder": "",
                "help_text": "モニター1回目計測時の写真2枚(後ろ向き・横向き)をご提出ください",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "image/*",
                "allow_multiple": 1,
                "sort_order": 3,
            },
            {
                "label": "計測写真(6回目or10回目)",
                "field_key": "last_measurement_images",
                "type": "file",
                "required": 1,
                "options": [],
                "placeholder": "",
                "help_text": "回数券6回目or10回目計測時の写真2枚(後ろ向き・横向き)をご提出ください",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "image/*",
                "allow_multiple": 1,
                "sort_order": 4,
            },
            {
                "label": "ご質問・ご相談(自由記述)",
                "field_key": "questions",
                "type": "long_text",
                "required": 0,
                "options": [],
                "placeholder": "",
                "help_text": "ご質問やご相談等があればご記載ください　＊内容はなんでも構いません！",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 5,
            },
        ],
    },
    {
        "title": "ビジリス施術アンケート",
        "slug": "bijiris-treatment-survey",
        "description": "本日の体感と施術内容、施術回数に応じた状況、ご質問をお聞かせください。",
        "success_message": "施術アンケートを受け付けました。ありがとうございました。",
        "category_label": "",
        "category_options": [],
        "fields": [
            {
                "label": "本日のビジリスの体感はいかがでしたか？",
                "field_key": "body_feedback",
                "type": "long_text",
                "required": 0,
                "options": [],
                "placeholder": "以前と比べて変化したことなどがあればご記載ください",
                "help_text": "以前と比べて変化したことなどがあればご記載ください",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 0,
            },
            {
                "label": "施術内容",
                "field_key": "treatment_type",
                "type": "select",
                "required": 1,
                "options": ["初回お試し", "単発", "モニター", "回数券", "乗り放題キャンペーン", "トライアル"],
                "placeholder": "",
                "help_text": "",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 1,
            },
            {
                "label": "単発と答えた方へ",
                "field_key": "single_visit_count",
                "type": "select",
                "required": 0,
                "options": ["1回目", "2回目", "3回目", "4回目", "5回目", "6回目", "7回目", "8回目", "9回目", "10回目", "11回目以上"],
                "placeholder": "",
                "help_text": "施術内容で単発と答えた方は単発での施術回数が何回目なのかを教えてください",
                "visibility_field_key": "treatment_type",
                "visibility_values": ["単発"],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 2,
            },
            {
                "label": "モニターと答えた方へ",
                "field_key": "monitor_visit_count",
                "type": "select",
                "required": 0,
                "options": ["1回目", "2回目", "3回目", "4回目", "5回目", "6回目"],
                "placeholder": "",
                "help_text": "施術内容でモニターと答えた方はモニターでの施術回数が何回目なのかを教えてください",
                "visibility_field_key": "treatment_type",
                "visibility_values": ["モニター"],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 3,
            },
            {
                "label": "回数券と答えた方へ\n現在お使いの回数券は何回券ですか？",
                "field_key": "ticket_book_type_answer",
                "type": "select",
                "required": 0,
                "options": ["6回券", "10回券"],
                "placeholder": "",
                "help_text": "施術内容で回数券と答えた方は、現在お使いの回数券が6回券か10回券かを教えてください",
                "visibility_field_key": "treatment_type",
                "visibility_values": ["回数券"],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 4,
            },
            {
                "label": "現在お使いの回数券で何回目の施術ですか？",
                "field_key": "ticket_visit_count",
                "type": "select",
                "required": 0,
                "options": ["1回目", "2回目", "3回目", "4回目", "5回目", "6回目", "7回目", "8回目", "9回目", "10回目"],
                "placeholder": "",
                "help_text": (
                    "施術内容で回数券と答えた方は回数券での施術回数が何回目なのかを教えてください\n"
                    "＊ただし、施術回数は累計ではなく現在お使いの回数券の施術回数が何回目なのかを教えてください"
                ),
                "visibility_field_key": "treatment_type",
                "visibility_values": ["回数券"],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 5,
            },
            {
                "label": "乗り放題キャンペーンと答えた方へ",
                "field_key": "campaign_visit_count",
                "type": "select",
                "required": 0,
                "options": ["1回目", "2回目", "3回目", "4回目", "5回目", "6回目", "7回目", "8回目", "9回目", "10回目", "11回目以上"],
                "placeholder": "",
                "help_text": (
                    "施術内容で乗り放題キャンペーンと答えた方は乗り放題キャンペーンでの施術回数が何回目なのかを教えてください\n"
                    "＊ただし、施術回数は累計ではなく乗り放題キャンペーンでの施術回数が何回目なのかを教えてください"
                ),
                "visibility_field_key": "treatment_type",
                "visibility_values": ["乗り放題キャンペーン"],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 6,
            },
            {
                "label": "トライアルと答えた方へ",
                "field_key": "trial_visit_count",
                "type": "select",
                "required": 0,
                "options": ["1回目", "2回目", "3回目", "4回目", "5回目", "6回目", "7回目", "8回目"],
                "placeholder": "",
                "help_text": (
                    "施術内容でトライアルと答えた方はトライアルでの施術回数が何回目なのかを教えてください\n"
                    "＊ただし、施術回数は累計ではなくトライアルでの施術回数が何回目なのかを教えてください"
                ),
                "visibility_field_key": "treatment_type",
                "visibility_values": ["トライアル"],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 7,
            },
            {
                "label": "ご質問・ご相談",
                "field_key": "questions",
                "type": "long_text",
                "required": 0,
                "options": [],
                "placeholder": "",
                "help_text": "ご質問やご相談等があればご記載ください　＊内容はなんでも構いません！",
                "visibility_field_key": "",
                "visibility_values": [],
                "accept": "",
                "allow_multiple": 0,
                "sort_order": 8,
            },
        ],
    },
]


def now_iso() -> str:
    return dt.datetime.now(dt.timezone.utc).astimezone().replace(microsecond=0).isoformat()


def slugify(value: str) -> str:
    value = unicodedata.normalize("NFKC", value or "").strip().lower()
    value = re.sub(r"[^\w\s-]", "", value)
    value = re.sub(r"[-\s]+", "-", value).strip("-_")
    return value or f"form-{secrets.token_hex(3)}"


def field_keyify(value: str) -> str:
    key = slugify(value).replace("-", "_")
    key = re.sub(r"[^a-z0-9_]", "", key)
    return key or f"field_{secrets.token_hex(3)}"


def sha256_hex(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def b64url_encode(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).rstrip(b"=").decode("ascii")


def b64url_decode(value: str) -> bytes:
    padding = "=" * (-len(value) % 4)
    return base64.urlsafe_b64decode(value + padding)


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    BACKUPS_DIR.mkdir(parents=True, exist_ok=True)
    PUBLIC_DIR.mkdir(parents=True, exist_ok=True)


def load_config() -> dict[str, Any]:
    ensure_dirs()
    if CONFIG_PATH.exists():
        with CONFIG_PATH.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
    else:
        data = {
            "admin_username": "admin",
            "admin_password_sha256": sha256_hex(DEFAULT_PASSWORD),
            "app_secret": secrets.token_hex(32),
            "public_base_url": "",
        }
        with CONFIG_PATH.open("w", encoding="utf-8") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
    if "app_secret" not in data:
        data["app_secret"] = secrets.token_hex(32)
    if "admin_username" not in data:
        data["admin_username"] = "admin"
    if "admin_password_sha256" not in data:
        data["admin_password_sha256"] = sha256_hex(DEFAULT_PASSWORD)
    if "public_base_url" not in data:
        data["public_base_url"] = ""
    return data


CONFIG = load_config()


def save_config(data: dict[str, Any]) -> dict[str, Any]:
    ensure_dirs()
    payload = {
        "admin_username": str(data.get("admin_username") or "admin").strip() or "admin",
        "admin_password_sha256": str(data.get("admin_password_sha256") or sha256_hex(DEFAULT_PASSWORD)).strip(),
        "app_secret": str(data.get("app_secret") or secrets.token_hex(32)).strip(),
        "public_base_url": str(data.get("public_base_url") or "").strip(),
    }
    temp_path = CONFIG_PATH.with_suffix(".tmp")
    with temp_path.open("w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)
    temp_path.replace(CONFIG_PATH)
    CONFIG.clear()
    CONFIG.update(payload)
    return payload


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def column_exists(conn: sqlite3.Connection, table_name: str, column_name: str) -> bool:
    rows = conn.execute(f"PRAGMA table_info({table_name})").fetchall()
    return any(row["name"] == column_name for row in rows)


def ensure_column(conn: sqlite3.Connection, table_name: str, column_name: str, definition: str) -> None:
    if not column_exists(conn, table_name, column_name):
        conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {definition}")


def remove_identity_fields(conn: sqlite3.Connection) -> None:
    form_ids = [
        row["form_id"]
        for row in conn.execute(
            """
            SELECT DISTINCT form_id
            FROM fields
            WHERE field_key IN ('name_display', 'respondent_id', 'respondent_name')
            """
        ).fetchall()
    ]
    if not form_ids:
        return

    conn.execute(
        """
        DELETE FROM fields
        WHERE field_key IN ('name_display', 'respondent_id', 'respondent_name')
        """
    )
    for form_id in form_ids:
        remaining = conn.execute(
            """
            SELECT id
            FROM fields
            WHERE form_id = ?
            ORDER BY sort_order ASC, id ASC
            """,
            (form_id,),
        ).fetchall()
        for index, row in enumerate(remaining):
            conn.execute("UPDATE fields SET sort_order = ? WHERE id = ?", (index, row["id"]))


def sync_respondents_registry(conn: sqlite3.Connection) -> None:
    timestamp = now_iso()
    response_rows = conn.execute(
        """
        SELECT respondent_id, MAX(respondent_name) AS respondent_name
        FROM responses
        GROUP BY respondent_id
        """
    ).fetchall()
    record_rows = conn.execute(
        """
        SELECT respondent_id, MAX(respondent_name) AS respondent_name
        FROM respondent_profile_records
        GROUP BY respondent_id
        """
    ).fetchall()
    registry = {}
    for row in response_rows + record_rows:
        respondent_id = str(row["respondent_id"] or "").strip()
        respondent_name = normalize_respondent_name(str(row["respondent_name"] or ""))
        if respondent_id and respondent_name:
            registry[respondent_id] = respondent_name
    for respondent_id, respondent_name in registry.items():
        existing = conn.execute(
            "SELECT respondent_id FROM respondents WHERE respondent_id = ?",
            (respondent_id,),
        ).fetchone()
        if existing:
            conn.execute(
                "UPDATE respondents SET respondent_name = ?, updated_at = ? WHERE respondent_id = ?",
                (respondent_name, timestamp, respondent_id),
            )
        else:
            conn.execute(
                "INSERT INTO respondents (respondent_id, respondent_name, created_at, updated_at) VALUES (?, ?, ?, ?)",
                (respondent_id, respondent_name, timestamp, timestamp),
            )


def normalize_respondent_name(value: str) -> str:
    name = unicodedata.normalize("NFKC", str(value or "")).strip()
    return re.sub(r"\s+", " ", name)


def respondent_name_match_key(value: str) -> str:
    normalized = normalize_respondent_name(value)
    return re.sub(r"\s+", "", normalized).lower()


def respondent_name_key(value: str) -> str:
    return normalize_respondent_name(value).lower()


def normalize_ticket_sheet_value(value: str) -> str:
    normalized = normalize_respondent_name(str(value or "")).translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    match = re.fullmatch(r"(\d{1,3})(?:枚目)?", normalized)
    if not match:
        raise ValueError("回数券が何枚目かを数字で入力してください。")
    number = int(match.group(1))
    if number <= 0:
        raise ValueError("回数券が何枚目かは1以上で入力してください。")
    return f"{number}枚目"


def normalize_optional_ticket_sheet_value(value: Any) -> str:
    text = normalize_respondent_name(str(value or ""))
    if not text:
        return ""
    return normalize_ticket_sheet_value(text)


def normalize_ticket_book_type(value: Any) -> str:
    text = normalize_respondent_name(str(value or ""))
    if not text:
        return ""
    if text not in TICKET_BOOK_TYPES:
        raise ValueError("回数券種別は 6回券 または 10回券 を選択してください。")
    return text


def ticket_book_stamp_max(ticket_book_type: str) -> int:
    if ticket_book_type == "6回券":
        return 6
    if ticket_book_type == "10回券":
        return 10
    return 10


def ticket_book_stamp_display_max(ticket_book_type: str) -> int:
    if ticket_book_type == "6回券":
        return 5
    if ticket_book_type == "10回券":
        return 9
    return 0


def normalize_ticket_stamp_count(value: Any, ticket_book_type: str) -> int:
    text = normalize_respondent_name(str(value or "")).translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    if not text:
        return 0
    if not re.fullmatch(r"\d{1,2}", text):
        raise ValueError("スタンプ数は数字で入力してください。")
    count = int(text)
    max_count = ticket_book_stamp_display_max(ticket_book_type)
    if count < 0:
        raise ValueError("スタンプ数は0以上で入力してください。")
    if count > max_count:
        raise ValueError(f"スタンプ数は {max_count} 以下で入力してください。")
    return count


def parse_ticket_visit_count(value: Any) -> int:
    text = normalize_respondent_name(str(value or "")).translate(str.maketrans("０１２３４５６７８９", "0123456789"))
    match = re.search(r"(\d{1,2})", text)
    return int(match.group(1)) if match else 0


def effective_ticket_stamp_count(ticket_book_type: str, raw_count: int) -> int:
    if raw_count <= 0:
        return 0
    max_count = ticket_book_stamp_display_max(ticket_book_type)
    if max_count <= 0:
        return raw_count
    return min(raw_count, max_count)


def fetch_respondent_row(conn: sqlite3.Connection, respondent_id: str) -> sqlite3.Row | None:
    return conn.execute(
        """
        SELECT *
        FROM respondents
        WHERE respondent_id = ?
        """,
        (respondent_id,),
    ).fetchone()


def update_respondent_ticket_sheet_manual_value(
    conn: sqlite3.Connection,
    respondent_id: str,
    ticket_sheet_manual_value: str,
) -> None:
    with conn:
        conn.execute(
            """
            UPDATE respondents
            SET ticket_sheet_manual_value = ?, updated_at = ?
            WHERE respondent_id = ?
            """,
            (ticket_sheet_manual_value, now_iso(), respondent_id),
        )


def update_respondent_ticket_status(
    conn: sqlite3.Connection,
    respondent_id: str,
    *,
    ticket_sheet_manual_value: str,
    current_ticket_book_type: str,
    current_ticket_stamp_count: int,
    current_ticket_stamp_manual_enabled: bool,
) -> None:
    with conn:
        conn.execute(
            """
            UPDATE respondents
            SET
                ticket_sheet_manual_value = ?,
                current_ticket_book_type = ?,
                current_ticket_stamp_count = ?,
                current_ticket_stamp_manual_enabled = ?,
                updated_at = ?
            WHERE respondent_id = ?
            """,
            (
                ticket_sheet_manual_value,
                current_ticket_book_type,
                current_ticket_stamp_count,
                1 if current_ticket_stamp_manual_enabled else 0,
                now_iso(),
                respondent_id,
            ),
        )


def fetch_respondent_row_by_name(
    conn: sqlite3.Connection,
    respondent_name: str,
) -> sqlite3.Row | None:
    normalized_name = normalize_respondent_name(respondent_name)
    if not normalized_name:
        return None
    target_key = respondent_name_match_key(normalized_name)
    rows = conn.execute(
        """
        SELECT *
        FROM respondents
        ORDER BY updated_at DESC, created_at DESC, respondent_id ASC
        """
    ).fetchall()
    for row in rows:
        if respondent_name_match_key(str(row["respondent_name"] or "")) == target_key:
            return row
    return None


def ensure_respondent_registry(
    conn: sqlite3.Connection,
    respondent_name: str,
    *,
    respondent_id: str | None = None,
    ticket_sheet_manual_value: str | None = None,
    current_ticket_book_type: str | None = None,
    current_ticket_stamp_count: int | None = None,
    current_ticket_stamp_manual_enabled: bool | None = None,
) -> dict[str, str]:
    normalized_name = normalize_respondent_name(respondent_name)
    if not normalized_name:
        raise ValueError("お名前は必須です。")
    registry_id = respondent_id or respondent_name_key(normalized_name)
    timestamp = now_iso()
    existing = fetch_respondent_row(conn, registry_id)
    with conn:
        if existing:
            if (
                ticket_sheet_manual_value is None
                and current_ticket_book_type is None
                and current_ticket_stamp_count is None
                and current_ticket_stamp_manual_enabled is None
            ):
                conn.execute(
                    """
                    UPDATE respondents
                    SET respondent_name = ?, updated_at = ?
                    WHERE respondent_id = ?
                    """,
                    (normalized_name, timestamp, registry_id),
                )
            else:
                conn.execute(
                    """
                    UPDATE respondents
                    SET
                        respondent_name = ?,
                        ticket_sheet_manual_value = ?,
                        current_ticket_book_type = ?,
                        current_ticket_stamp_count = ?,
                        current_ticket_stamp_manual_enabled = ?,
                        updated_at = ?
                    WHERE respondent_id = ?
                    """,
                    (
                        normalized_name,
                        ticket_sheet_manual_value or "",
                        current_ticket_book_type or "",
                        current_ticket_stamp_count or 0,
                        1 if current_ticket_stamp_manual_enabled else 0,
                        timestamp,
                        registry_id,
                    ),
                )
        else:
            conn.execute(
                """
                INSERT INTO respondents (
                    respondent_id, respondent_name, ticket_sheet_manual_value,
                    current_ticket_book_type, current_ticket_stamp_count,
                    current_ticket_stamp_manual_enabled, created_at, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    registry_id,
                    normalized_name,
                    ticket_sheet_manual_value or "",
                    current_ticket_book_type or "",
                    current_ticket_stamp_count or 0,
                    1 if current_ticket_stamp_manual_enabled else 0,
                    timestamp,
                    timestamp,
                ),
            )
    return {"respondentId": registry_id, "respondentName": normalized_name}


def insert_form_definition(conn: sqlite3.Connection, form_def: dict[str, Any]) -> None:
    timestamp = now_iso()
    cursor = conn.execute(
        """
        INSERT INTO forms (
            title, slug, description, success_message,
            category_label, category_options, is_active, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, 1, ?, ?)
        """,
        (
            form_def["title"],
            form_def["slug"],
            form_def["description"],
            form_def["success_message"],
            form_def.get("category_label", ""),
            json.dumps(form_def.get("category_options", []), ensure_ascii=False),
            timestamp,
            timestamp,
        ),
    )
    form_id = cursor.lastrowid
    for field in form_def["fields"]:
        conn.execute(
            """
            INSERT INTO fields (
                form_id, label, field_key, type, required, options, placeholder,
                help_text, visibility_field_key, visibility_values, accept, allow_multiple, allow_other, sort_order
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                form_id,
                field["label"],
                field["field_key"],
                field["type"],
                field["required"],
                json.dumps(field["options"], ensure_ascii=False),
                field["placeholder"],
                field.get("help_text", ""),
                field.get("visibility_field_key", ""),
                json.dumps(field.get("visibility_values", []), ensure_ascii=False),
                field.get("accept", ""),
                field.get("allow_multiple", 0),
                field.get("allow_other", 0),
                field["sort_order"],
            ),
        )


def normalize_reference_fields(fields: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        {
            "label": field["label"],
            "field_key": field["field_key"],
            "type": field["type"],
            "required": int(field["required"]),
            "options": list(field["options"]),
            "placeholder": field.get("placeholder", ""),
            "help_text": field.get("help_text", ""),
            "visibility_field_key": field.get("visibility_field_key", ""),
            "visibility_values": list(field.get("visibility_values", [])),
            "accept": field.get("accept", ""),
            "allow_multiple": int(field.get("allow_multiple", 0)),
            "allow_other": int(field.get("allow_other", 0)),
            "sort_order": int(field["sort_order"]),
        }
        for field in fields
    ]


def current_reference_field_signature(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    forms = conn.execute(
        """
        SELECT id, title, slug
        FROM forms
        ORDER BY id ASC
        """
    ).fetchall()
    signature: list[dict[str, Any]] = []
    for form in forms:
        field_rows = conn.execute(
            """
            SELECT
                label,
                field_key,
                type,
                required,
                options,
                placeholder,
                help_text,
                visibility_field_key,
                visibility_values,
                accept,
                allow_multiple,
                allow_other,
                sort_order
            FROM fields
            WHERE form_id = ?
            ORDER BY sort_order ASC, id ASC
            """,
            (form["id"],),
        ).fetchall()
        signature.append(
            {
                "title": form["title"],
                "slug": form["slug"],
                "fields": [
                    {
                        "label": row["label"],
                        "field_key": row["field_key"],
                        "type": row["type"],
                        "required": int(row["required"]),
                        "options": json.loads(row["options"] or "[]"),
                        "placeholder": row["placeholder"] or "",
                        "help_text": row["help_text"] or "",
                        "visibility_field_key": row["visibility_field_key"] or "",
                        "visibility_values": json.loads(row["visibility_values"] or "[]"),
                        "accept": row["accept"] or "",
                        "allow_multiple": int(row["allow_multiple"]),
                        "allow_other": int(row["allow_other"]),
                        "sort_order": int(row["sort_order"]),
                    }
                    for row in field_rows
                ],
            }
        )
    return signature


def expected_reference_field_signature() -> list[dict[str, Any]]:
    return [
        {
            "title": form_def["title"],
            "slug": form_def["slug"],
            "fields": normalize_reference_fields(form_def["fields"]),
        }
        for form_def in REFERENCE_FORM_DEFINITIONS
    ]


def ensure_reference_forms(conn: sqlite3.Connection) -> None:
    response_count = conn.execute("SELECT COUNT(*) AS count FROM responses").fetchone()["count"]
    if response_count > 0:
        return

    existing_slugs = [row["slug"] for row in conn.execute("SELECT slug FROM forms ORDER BY id ASC").fetchall()]
    reference_slugs = [form_def["slug"] for form_def in REFERENCE_FORM_DEFINITIONS]
    if existing_slugs and existing_slugs != ["genba-houkoku"]:
        if existing_slugs != reference_slugs:
            return
        if current_reference_field_signature(conn) == expected_reference_field_signature():
            return

    conn.execute("DELETE FROM forms")
    for form_def in REFERENCE_FORM_DEFINITIONS:
        insert_form_definition(conn, form_def)


def migrate_legacy_respondent_profiles(conn: sqlite3.Connection) -> None:
    has_legacy_table = conn.execute(
        """
        SELECT COUNT(*) AS count
        FROM sqlite_master
        WHERE type = 'table' AND name = 'respondent_profiles'
        """
    ).fetchone()["count"]
    if not has_legacy_table:
        return

    rows = conn.execute(
        """
        SELECT *
        FROM respondent_profiles
        ORDER BY updated_at DESC, created_at DESC, respondent_id ASC
        """
    ).fetchall()
    if not rows:
        return

    migrated_any = False
    for row in rows:
        if not row["image_relative_path"] and not row["profile_date"]:
            continue
        entry_date = row["profile_date"] or str(row["updated_at"] or row["created_at"] or "")[:10]
        if not entry_date:
            entry_date = dt.date.today().isoformat()
        title = "旧プロフィール画像" if row["image_relative_path"] else "旧プロフィール記録"
        memo = "旧データから移行"
        conn.execute(
            """
            INSERT INTO respondent_profile_records (
                respondent_id, respondent_name, title, entry_date, memo,
                image_original_name, image_stored_name, image_mime_type, image_relative_path,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                row["respondent_id"],
                row["respondent_name"] or "",
                title,
                entry_date,
                memo,
                row["image_original_name"] or "",
                row["image_stored_name"] or "",
                row["image_mime_type"] or "",
                row["image_relative_path"] or "",
                row["created_at"] or now_iso(),
                row["updated_at"] or now_iso(),
            ),
        )
        migrated_any = True

    if migrated_any:
        conn.execute("DELETE FROM respondent_profiles")


def init_db() -> None:
    ensure_dirs()
    conn = get_connection()
    try:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS forms (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                slug TEXT NOT NULL UNIQUE,
                description TEXT NOT NULL DEFAULT '',
                success_message TEXT NOT NULL DEFAULT '',
                category_label TEXT NOT NULL DEFAULT '分類',
                category_options TEXT NOT NULL DEFAULT '[]',
                is_active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS fields (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                form_id INTEGER NOT NULL REFERENCES forms(id) ON DELETE CASCADE,
                label TEXT NOT NULL,
                field_key TEXT NOT NULL,
                type TEXT NOT NULL,
                required INTEGER NOT NULL DEFAULT 0,
                options TEXT NOT NULL DEFAULT '[]',
                placeholder TEXT NOT NULL DEFAULT '',
                sort_order INTEGER NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS responses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                form_id INTEGER NOT NULL REFERENCES forms(id) ON DELETE CASCADE,
                respondent_id TEXT NOT NULL,
                respondent_name TEXT NOT NULL,
                respondent_email TEXT NOT NULL DEFAULT '',
                category TEXT NOT NULL,
                notes TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                ip_address TEXT NOT NULL DEFAULT '',
                user_agent TEXT NOT NULL DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS respondents (
                respondent_id TEXT PRIMARY KEY,
                respondent_name TEXT NOT NULL,
                ticket_sheet_manual_value TEXT NOT NULL DEFAULT '',
                current_ticket_book_type TEXT NOT NULL DEFAULT '',
                current_ticket_stamp_count INTEGER NOT NULL DEFAULT 0,
                current_ticket_stamp_manual_enabled INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS response_answers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                response_id INTEGER NOT NULL REFERENCES responses(id) ON DELETE CASCADE,
                field_key TEXT NOT NULL,
                label TEXT NOT NULL,
                value TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS response_files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                response_id INTEGER NOT NULL REFERENCES responses(id) ON DELETE CASCADE,
                original_name TEXT NOT NULL,
                stored_name TEXT NOT NULL UNIQUE,
                mime_type TEXT NOT NULL,
                size INTEGER NOT NULL DEFAULT 0,
                relative_path TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS respondent_profiles (
                respondent_id TEXT PRIMARY KEY,
                respondent_name TEXT NOT NULL DEFAULT '',
                profile_date TEXT NOT NULL DEFAULT '',
                image_original_name TEXT NOT NULL DEFAULT '',
                image_stored_name TEXT NOT NULL DEFAULT '',
                image_mime_type TEXT NOT NULL DEFAULT '',
                image_relative_path TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS respondent_profile_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                respondent_id TEXT NOT NULL,
                respondent_name TEXT NOT NULL DEFAULT '',
                title TEXT NOT NULL DEFAULT '',
                entry_date TEXT NOT NULL DEFAULT '',
                memo TEXT NOT NULL DEFAULT '',
                image_original_name TEXT NOT NULL DEFAULT '',
                image_stored_name TEXT NOT NULL DEFAULT '',
                image_mime_type TEXT NOT NULL DEFAULT '',
                image_relative_path TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS response_file_annotations (
                response_file_id INTEGER PRIMARY KEY REFERENCES response_files(id) ON DELETE CASCADE,
                title TEXT NOT NULL DEFAULT '',
                entry_date TEXT NOT NULL DEFAULT '',
                memo TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS respondent_measurements (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                respondent_id TEXT NOT NULL,
                respondent_name TEXT NOT NULL DEFAULT '',
                entry_date TEXT NOT NULL,
                measurement_category TEXT NOT NULL DEFAULT '',
                waist REAL NOT NULL,
                hip REAL NOT NULL,
                thigh REAL NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_forms_slug ON forms(slug);
            CREATE INDEX IF NOT EXISTS idx_respondents_name ON respondents(respondent_name);
            CREATE INDEX IF NOT EXISTS idx_responses_form_id ON responses(form_id);
            CREATE INDEX IF NOT EXISTS idx_responses_respondent_id ON responses(respondent_id);
            CREATE INDEX IF NOT EXISTS idx_responses_category ON responses(category);
            CREATE INDEX IF NOT EXISTS idx_respondent_profiles_name ON respondent_profiles(respondent_name);
            CREATE INDEX IF NOT EXISTS idx_respondent_profile_records_respondent_id
            ON respondent_profile_records(respondent_id);
            CREATE INDEX IF NOT EXISTS idx_respondent_profile_records_entry_date
            ON respondent_profile_records(entry_date DESC, updated_at DESC);
            CREATE INDEX IF NOT EXISTS idx_response_file_annotations_entry_date
            ON response_file_annotations(entry_date DESC, updated_at DESC);
            CREATE INDEX IF NOT EXISTS idx_respondent_measurements_respondent_id
            ON respondent_measurements(respondent_id);
            CREATE INDEX IF NOT EXISTS idx_respondent_measurements_entry_date
            ON respondent_measurements(entry_date DESC, updated_at DESC);
            """
        )
        ensure_column(conn, "fields", "help_text", "TEXT NOT NULL DEFAULT ''")
        ensure_column(conn, "fields", "visibility_field_key", "TEXT NOT NULL DEFAULT ''")
        ensure_column(conn, "fields", "visibility_values", "TEXT NOT NULL DEFAULT '[]'")
        ensure_column(conn, "fields", "accept", "TEXT NOT NULL DEFAULT ''")
        ensure_column(conn, "fields", "allow_multiple", "INTEGER NOT NULL DEFAULT 0")
        ensure_column(conn, "fields", "allow_other", "INTEGER NOT NULL DEFAULT 0")
        ensure_column(conn, "respondents", "ticket_sheet_manual_value", "TEXT NOT NULL DEFAULT ''")
        ensure_column(conn, "respondents", "current_ticket_book_type", "TEXT NOT NULL DEFAULT ''")
        ensure_column(conn, "respondents", "current_ticket_stamp_count", "INTEGER NOT NULL DEFAULT 0")
        ensure_column(conn, "respondents", "current_ticket_stamp_manual_enabled", "INTEGER NOT NULL DEFAULT 0")
        conn.execute(
            """
            UPDATE respondents
            SET current_ticket_stamp_manual_enabled = 1
            WHERE current_ticket_stamp_count > 0 AND current_ticket_stamp_manual_enabled = 0
            """
        )
        ensure_column(conn, "response_files", "field_key", "TEXT NOT NULL DEFAULT ''")
        ensure_column(conn, "response_files", "label", "TEXT NOT NULL DEFAULT ''")
        ensure_column(conn, "respondent_measurements", "measurement_category", "TEXT NOT NULL DEFAULT ''")
        sync_upload_url_columns(conn)
        remove_identity_fields(conn)
        migrate_legacy_respondent_profiles(conn)
        sync_respondents_registry(conn)
        conn.commit()
        ensure_reference_forms(conn)
        conn.commit()
    finally:
        conn.close()


def sign_session(username: str) -> str:
    payload = json.dumps(
        {"u": username, "exp": int(dt.datetime.now().timestamp()) + (14 * 24 * 60 * 60)},
        ensure_ascii=False,
        separators=(",", ":"),
    ).encode("utf-8")
    encoded_payload = b64url_encode(payload)
    signature = hmac.new(
        CONFIG["app_secret"].encode("utf-8"),
        encoded_payload.encode("utf-8"),
        hashlib.sha256,
    ).digest()
    return f"{encoded_payload}.{b64url_encode(signature)}"


def verify_session(token: str | None) -> str | None:
    if not token or "." not in token:
        return None
    encoded_payload, encoded_signature = token.split(".", 1)
    expected = hmac.new(
        CONFIG["app_secret"].encode("utf-8"),
        encoded_payload.encode("utf-8"),
        hashlib.sha256,
    ).digest()
    try:
        actual = b64url_decode(encoded_signature)
        if not hmac.compare_digest(expected, actual):
            return None
        payload = json.loads(b64url_decode(encoded_payload).decode("utf-8"))
    except Exception:
        return None
    if payload.get("exp", 0) < int(dt.datetime.now().timestamp()):
        return None
    return payload.get("u")


def parse_cookie_header(header: str | None) -> dict[str, str]:
    result: dict[str, str] = {}
    if not header:
        return result
    for part in header.split(";"):
        if "=" not in part:
            continue
        key, value = part.strip().split("=", 1)
        result[key] = value
    return result


def parse_json_body(handler: BaseHTTPRequestHandler) -> dict[str, Any]:
    length = int(handler.headers.get("Content-Length", "0") or "0")
    raw = handler.rfile.read(length) if length > 0 else b"{}"
    if not raw:
        return {}
    return json.loads(raw.decode("utf-8"))


def read_public_asset(name: str) -> Path | None:
    target = (PUBLIC_DIR / name).resolve()
    try:
        target.relative_to(PUBLIC_DIR.resolve())
    except ValueError:
        return None
    return target if target.exists() and target.is_file() else None


def read_upload_target(name: str) -> Path | None:
    target = (UPLOADS_DIR / name).resolve()
    try:
        target.relative_to(UPLOADS_DIR.resolve())
    except ValueError:
        return None
    return target if target.exists() and target.is_file() else None


def clean_form_payload(payload: dict[str, Any], current_form_id: int | None = None) -> dict[str, Any]:
    title = str(payload.get("title", "")).strip()
    if not title:
        raise ValueError("フォーム名は必須です。")

    slug = slugify(str(payload.get("slug", "")).strip() or title)
    description = str(payload.get("description", "")).strip()
    success_message = str(payload.get("successMessage", "")).strip() or "送信ありがとうございました。"
    category_label = str(payload.get("categoryLabel", "")).strip() or "分類"

    raw_options = payload.get("categoryOptions") or []
    category_options = [str(option).strip() for option in raw_options if str(option).strip()]

    raw_fields = payload.get("fields") or []
    cleaned_fields: list[dict[str, Any]] = []
    seen_keys: set[str] = set()
    for index, raw_field in enumerate(raw_fields):
        label = str(raw_field.get("label", "")).strip()
        if not label:
            raise ValueError("追加項目のラベルは必須です。")
        field_type = str(raw_field.get("type", "")).strip()
        if field_type not in ALLOWED_FIELD_TYPES:
            raise ValueError("未対応の項目タイプが含まれています。")
        field_key = field_keyify(str(raw_field.get("key", "")).strip() or label)
        if field_key in seen_keys:
            raise ValueError("追加項目のキーが重複しています。")
        seen_keys.add(field_key)

        options = raw_field.get("options") or []
        clean_options = [str(option).strip() for option in options if str(option).strip()]
        if field_type in {"select", "radio", "checkbox"} and not clean_options:
            raise ValueError(f"{label} の選択肢を設定してください。")
        visibility_field_key = field_keyify(str(raw_field.get("visibilityFieldKey", "")).strip()) if str(raw_field.get("visibilityFieldKey", "")).strip() else ""
        visibility_values = [
            str(option).strip() for option in (raw_field.get("visibilityValues") or []) if str(option).strip()
        ]
        allow_multiple = 1 if raw_field.get("allowMultiple") else 0
        allow_other = 1 if field_type == "checkbox" and raw_field.get("allowOther") else 0
        accept = str(raw_field.get("accept", "")).strip()
        help_text = str(raw_field.get("helpText", "")).strip()
        if field_type == "file":
            clean_options = []
            accept = accept or "image/*"

        cleaned_fields.append(
            {
                "label": label,
                "field_key": field_key,
                "type": field_type,
                "required": 1 if raw_field.get("required") else 0,
                "options": clean_options,
                "placeholder": str(raw_field.get("placeholder", "")).strip(),
                "help_text": help_text,
                "visibility_field_key": visibility_field_key,
                "visibility_values": visibility_values,
                "accept": accept,
                "allow_multiple": allow_multiple,
                "allow_other": allow_other,
                "sort_order": index,
            }
        )

    return {
        "title": title,
        "slug": slug,
        "description": description,
        "success_message": success_message,
        "category_label": category_label,
        "category_options": category_options,
        "is_active": 1 if payload.get("isActive", True) else 0,
        "fields": cleaned_fields,
        "current_form_id": current_form_id,
    }


def fetch_fields_map(conn: sqlite3.Connection, form_ids: list[int]) -> dict[int, list[dict[str, Any]]]:
    if not form_ids:
        return {}
    placeholders = ",".join("?" for _ in form_ids)
    rows = conn.execute(
        f"""
        SELECT * FROM fields
        WHERE form_id IN ({placeholders})
        ORDER BY sort_order ASC, id ASC
        """,
        form_ids,
    ).fetchall()
    result: dict[int, list[dict[str, Any]]] = {form_id: [] for form_id in form_ids}
    for row in rows:
        result.setdefault(row["form_id"], []).append(
            {
                "id": row["id"],
                "label": row["label"],
                "key": row["field_key"],
                "type": row["type"],
                "required": bool(row["required"]),
                "options": json.loads(row["options"] or "[]"),
                "placeholder": row["placeholder"] or "",
                "helpText": row["help_text"] or "",
                "visibilityFieldKey": row["visibility_field_key"] or "",
                "visibilityValues": json.loads(row["visibility_values"] or "[]"),
                "accept": row["accept"] or "",
                "allowMultiple": bool(row["allow_multiple"]),
                "allowOther": bool(row["allow_other"]),
            }
        )
    return result


def fetch_forms(conn: sqlite3.Connection, include_inactive: bool = True) -> list[dict[str, Any]]:
    where = "" if include_inactive else "WHERE forms.is_active = 1"
    rows = conn.execute(
        f"""
        SELECT
            forms.*,
            COUNT(responses.id) AS response_count,
            COUNT(DISTINCT responses.respondent_id) AS respondent_count
        FROM forms
        LEFT JOIN responses ON responses.form_id = forms.id
        {where}
        GROUP BY forms.id
        ORDER BY forms.updated_at DESC, forms.id DESC
        """
    ).fetchall()
    form_ids = [row["id"] for row in rows]
    fields_map = fetch_fields_map(conn, form_ids)
    forms: list[dict[str, Any]] = []
    for row in rows:
        forms.append(
            {
                "id": row["id"],
                "title": row["title"],
                "slug": row["slug"],
                "description": row["description"],
                "successMessage": row["success_message"],
                "categoryLabel": row["category_label"],
                "categoryOptions": json.loads(row["category_options"] or "[]"),
                "isActive": bool(row["is_active"]),
                "createdAt": row["created_at"],
                "updatedAt": row["updated_at"],
                "responseCount": row["response_count"],
                "respondentCount": row["respondent_count"],
                "fields": fields_map.get(row["id"], []),
            }
        )
    return forms


def fetch_form_by_id(conn: sqlite3.Connection, form_id: int) -> dict[str, Any] | None:
    for form in fetch_forms(conn, include_inactive=True):
        if form["id"] == form_id:
            return form
    return None


def fetch_form_by_slug(conn: sqlite3.Connection, slug: str, include_inactive: bool = False) -> dict[str, Any] | None:
    forms = fetch_forms(conn, include_inactive=include_inactive)
    for form in forms:
        if form["slug"] == slug:
            return form
    return None


def fetch_public_forms(conn: sqlite3.Connection) -> list[dict[str, Any]]:
    return [
        {
            "id": form["id"],
            "title": form["title"],
            "slug": form["slug"],
            "description": form["description"],
            "questionCount": len(form["fields"]),
        }
        for form in fetch_forms(conn, include_inactive=False)
    ]


def save_form(conn: sqlite3.Connection, payload: dict[str, Any], form_id: int | None = None) -> dict[str, Any]:
    cleaned = clean_form_payload(payload, current_form_id=form_id)
    existing = conn.execute(
        "SELECT id FROM forms WHERE slug = ? AND (? IS NULL OR id != ?)",
        (cleaned["slug"], form_id, form_id),
    ).fetchone()
    if existing:
        raise ValueError("同じURLスラッグのフォームが既に存在します。")

    timestamp = now_iso()
    with conn:
        if form_id is None:
            cursor = conn.execute(
                """
                INSERT INTO forms (
                    title, slug, description, success_message,
                    category_label, category_options, is_active, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    cleaned["title"],
                    cleaned["slug"],
                    cleaned["description"],
                    cleaned["success_message"],
                    cleaned["category_label"],
                    json.dumps(cleaned["category_options"], ensure_ascii=False),
                    cleaned["is_active"],
                    timestamp,
                    timestamp,
                ),
            )
            form_id = cursor.lastrowid
        else:
            conn.execute(
                """
                UPDATE forms
                SET title = ?, slug = ?, description = ?, success_message = ?,
                    category_label = ?, category_options = ?, is_active = ?, updated_at = ?
                WHERE id = ?
                """,
                (
                    cleaned["title"],
                    cleaned["slug"],
                    cleaned["description"],
                    cleaned["success_message"],
                    cleaned["category_label"],
                    json.dumps(cleaned["category_options"], ensure_ascii=False),
                    cleaned["is_active"],
                    timestamp,
                    form_id,
                ),
            )
            conn.execute("DELETE FROM fields WHERE form_id = ?", (form_id,))

        for field in cleaned["fields"]:
            conn.execute(
                """
                INSERT INTO fields (
                    form_id, label, field_key, type, required, options, placeholder,
                    help_text, visibility_field_key, visibility_values, accept, allow_multiple, allow_other, sort_order
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    form_id,
                    field["label"],
                    field["field_key"],
                    field["type"],
                    field["required"],
                    json.dumps(field["options"], ensure_ascii=False),
                    field["placeholder"],
                    field["help_text"],
                    field["visibility_field_key"],
                    json.dumps(field["visibility_values"], ensure_ascii=False),
                    field["accept"],
                    field["allow_multiple"],
                    field["allow_other"],
                    field["sort_order"],
                ),
            )

    saved = fetch_form_by_id(conn, int(form_id))
    if not saved:
        raise RuntimeError("フォームの保存に失敗しました。")
    return saved


def toggle_form(conn: sqlite3.Connection, form_id: int) -> dict[str, Any]:
    form = fetch_form_by_id(conn, form_id)
    if not form:
        raise ValueError("フォームが見つかりません。")
    with conn:
        conn.execute(
            "UPDATE forms SET is_active = ?, updated_at = ? WHERE id = ?",
            (0 if form["isActive"] else 1, now_iso(), form_id),
        )
    updated = fetch_form_by_id(conn, form_id)
    if not updated:
        raise RuntimeError("フォームの更新に失敗しました。")
    return updated


def fetch_files_map(conn: sqlite3.Connection, response_ids: list[int]) -> dict[int, list[dict[str, Any]]]:
    if not response_ids:
        return {}
    placeholders = ",".join("?" for _ in response_ids)
    rows = conn.execute(
        f"""
        SELECT * FROM response_files
        WHERE response_id IN ({placeholders})
        ORDER BY id ASC
        """,
        response_ids,
    ).fetchall()
    result: dict[int, list[dict[str, Any]]] = {response_id: [] for response_id in response_ids}
    for row in rows:
        result.setdefault(row["response_id"], []).append(
            {
                "id": row["id"],
                "fieldKey": row["field_key"],
                "label": row["label"],
                "originalName": row["original_name"],
                "mimeType": row["mime_type"],
                "size": row["size"],
                "url": resolve_public_image_url(row["relative_path"], row["stored_name"]),
                "previewUrl": resolve_public_image_preview_url(row["relative_path"], row["stored_name"]),
            }
        )
    return result


def response_file_image_payload(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "originalName": row["original_name"] or "",
        "storedName": row["stored_name"] or "",
        "mimeType": row["mime_type"] or "",
        "url": resolve_public_image_url(row["relative_path"], row["stored_name"]),
        "previewUrl": resolve_public_image_preview_url(row["relative_path"], row["stored_name"]),
    }


def default_public_base_url() -> str:
    configured = os.environ.get(PUBLIC_BASE_URL_ENV, "").strip()
    if configured:
        return configured.rstrip("/")
    render_external = os.environ.get(RENDER_EXTERNAL_URL_ENV, "").strip()
    if render_external:
        return render_external.rstrip("/")
    configured = str(CONFIG.get("public_base_url") or "").strip()
    if configured:
        return configured.rstrip("/")
    lan_ip = detect_lan_ipv4()
    if lan_ip:
        return f"http://{lan_ip}:{DEFAULT_SERVER_PORT}"
    return f"http://127.0.0.1:{DEFAULT_SERVER_PORT}"


def local_upload_path(stored_name: Any) -> str:
    raw = str(stored_name or "").strip()
    if not raw:
        return ""
    return f"/uploads/{urllib.parse.quote(raw)}"


def public_upload_url(stored_name: Any) -> str:
    relative = local_upload_path(stored_name)
    if not relative:
        return ""
    return f"{default_public_base_url()}{relative}"


def public_base_url_source() -> str:
    if os.environ.get(PUBLIC_BASE_URL_ENV, "").strip():
        return "env"
    if os.environ.get(RENDER_EXTERNAL_URL_ENV, "").strip():
        return "render"
    if str(CONFIG.get("public_base_url") or "").strip():
        return "config"
    return "auto"


def normalize_public_base_url(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    parsed = urllib.parse.urlparse(raw)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise ValueError("公開URLは http または https の完全なURLで入力してください。")
    normalized = parsed._replace(path="", params="", query="", fragment="")
    return normalized.geturl().rstrip("/")


def update_public_base_url_config(value: Any) -> str:
    normalized = normalize_public_base_url(value) if str(value or "").strip() else ""
    updated = dict(CONFIG)
    updated["public_base_url"] = normalized
    save_config(updated)
    return normalized


def update_admin_password(current_password: str, new_password: str, confirm_password: str) -> None:
    if sha256_hex(str(current_password or "")) != CONFIG["admin_password_sha256"]:
        raise ValueError("現在のパスワードが正しくありません。")
    new_value = str(new_password or "")
    confirm_value = str(confirm_password or "")
    if len(new_value) < 8:
        raise ValueError("新しいパスワードは8文字以上にしてください。")
    if new_value != confirm_value:
        raise ValueError("新しいパスワードと確認用パスワードが一致しません。")
    updated = dict(CONFIG)
    updated["admin_password_sha256"] = sha256_hex(new_value)
    save_config(updated)


def upload_stored_name_from_source(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    if raw.startswith("/uploads/"):
        return urllib.parse.unquote(raw.replace("/uploads/", "", 1))
    if re.match(r"^https?://", raw, re.IGNORECASE):
        parsed = urllib.parse.urlparse(raw)
        if parsed.path.startswith("/uploads/"):
            return urllib.parse.unquote(parsed.path.replace("/uploads/", "", 1))
    return ""


def google_drive_file_id_from_url(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw or not re.match(r"^https?://", raw, re.IGNORECASE):
        return ""
    parsed = urllib.parse.urlparse(raw)
    if "drive.google.com" not in parsed.netloc.lower():
        return ""
    query = urllib.parse.parse_qs(parsed.query)
    file_id = (query.get("id") or [""])[-1].strip()
    if not file_id:
        match = re.search(r"/file/d/([a-zA-Z0-9_-]+)", parsed.path)
        if match:
            file_id = match.group(1)
    return file_id


def google_drive_thumbnail_url(file_id: str, *, width: int = 2000) -> str:
    normalized = str(file_id or "").strip()
    if not normalized:
        return ""
    width_value = max(320, min(int(width or 2000), 4000))
    return f"https://drive.google.com/thumbnail?id={urllib.parse.quote(normalized)}&sz=w{width_value}"


def resolve_public_image_url(value: Any, stored_name: Any = "") -> str:
    raw = str(value or "").strip()
    if not raw:
        return public_upload_url(stored_name)
    if re.match(r"^https?://", raw, re.IGNORECASE):
        file_id = google_drive_file_id_from_url(raw)
        if file_id:
            return google_drive_thumbnail_url(file_id, width=2000)
        return raw
    if raw.startswith("/uploads/"):
        return f"{default_public_base_url()}{raw}"
    return public_upload_url(raw)


def resolve_public_image_preview_url(value: Any, stored_name: Any = "") -> str:
    raw = str(value or "").strip()
    if not raw:
        return public_upload_url(stored_name)
    file_id = google_drive_file_id_from_url(raw)
    if file_id:
        return google_drive_thumbnail_url(file_id, width=1200)
    return resolve_public_image_url(raw, stored_name)


def download_remote_image(source_url: str) -> tuple[bytes, str]:
    raw = str(source_url or "").strip()
    if not raw or not re.match(r"^https?://", raw, re.IGNORECASE):
        raise ValueError("画像URLが不正です。")
    request = urllib.request.Request(
        raw,
        headers={
            "User-Agent": "Bijiris/1.0",
            "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
            "Referer": "https://drive.google.com/",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            payload = response.read()
            content_type = response.headers.get_content_type() or "application/octet-stream"
    except urllib.error.HTTPError as exc:
        if exc.code in {401, 403}:
            raise PermissionError("画像URLにアクセスする権限がありません。") from exc
        raise ValueError("画像URLの取得に失敗しました。") from exc
    except urllib.error.URLError as exc:
        raise ValueError("画像URLへ接続できませんでした。") from exc
    if not payload:
        raise ValueError("画像データが空でした。")
    return payload, content_type


def sync_upload_url_columns(conn: sqlite3.Connection) -> None:
    response_rows = conn.execute(
        """
        SELECT id, stored_name, relative_path
        FROM response_files
        WHERE stored_name <> ''
        """
    ).fetchall()
    for row in response_rows:
        current_path = str(row["relative_path"] or "").strip()
        desired = local_upload_path(row["stored_name"])
        if desired and should_sync_local_upload_path(current_path, row["stored_name"]) and current_path != desired:
            conn.execute(
                "UPDATE response_files SET relative_path = ? WHERE id = ?",
                (desired, row["id"]),
            )

    profile_rows = conn.execute(
        """
        SELECT id, image_stored_name, image_relative_path
        FROM respondent_profile_records
        WHERE image_stored_name <> ''
        """
    ).fetchall()
    for row in profile_rows:
        current_path = str(row["image_relative_path"] or "").strip()
        desired = local_upload_path(row["image_stored_name"])
        if desired and should_sync_local_upload_path(current_path, row["image_stored_name"]) and current_path != desired:
            conn.execute(
                "UPDATE respondent_profile_records SET image_relative_path = ? WHERE id = ?",
                (desired, row["id"]),
            )


def build_backup_filename() -> str:
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    return f"bijiris-backup-{stamp}.tar.gz"


def create_backup_archive() -> dict[str, Any]:
    ensure_dirs()
    backup_name = build_backup_filename()
    backup_path = BACKUPS_DIR / backup_name
    with tarfile.open(backup_path, "w:gz") as archive:
        if DB_PATH.exists():
            archive.add(DB_PATH, arcname="data/bijiris.db")
        if CONFIG_PATH.exists():
            archive.add(CONFIG_PATH, arcname="data/config.json")
        if UPLOADS_DIR.exists():
            archive.add(UPLOADS_DIR, arcname="data/uploads")
    stat = backup_path.stat()
    return {
        "name": backup_name,
        "path": str(backup_path),
        "sizeBytes": stat.st_size,
        "createdAt": dt.datetime.fromtimestamp(stat.st_mtime, tz=dt.timezone.utc)
        .astimezone()
        .replace(microsecond=0)
        .isoformat(),
    }


def list_backup_archives(limit: int = 10) -> list[dict[str, Any]]:
    ensure_dirs()
    archives = sorted(BACKUPS_DIR.glob("bijiris-backup-*.tar.gz"), key=lambda path: path.stat().st_mtime, reverse=True)
    result: list[dict[str, Any]] = []
    for path in archives[: max(1, limit)]:
        stat = path.stat()
        result.append(
            {
                "name": path.name,
                "path": str(path),
                "sizeBytes": stat.st_size,
                "createdAt": dt.datetime.fromtimestamp(stat.st_mtime, tz=dt.timezone.utc)
                .astimezone()
                .replace(microsecond=0)
                .isoformat(),
                "downloadUrl": f"/api/admin/backups/{urllib.parse.quote(path.name)}/download",
            }
        )
    return result


def find_backup_archive(name: str) -> Path | None:
    raw = Path(str(name or "").strip()).name
    if not raw:
        return None
    candidate = BACKUPS_DIR / raw
    if not candidate.exists() or not candidate.is_file():
        return None
    return candidate


def admin_operations_status(conn: sqlite3.Connection, *, public_base_url: str, server_port: int) -> dict[str, Any]:
    uploads_count = int(
        conn.execute("SELECT COUNT(*) AS count FROM response_files WHERE relative_path <> ''").fetchone()["count"]
    )
    profile_images_count = int(
        conn.execute(
            "SELECT COUNT(*) AS count FROM respondent_profile_records WHERE image_relative_path <> ''"
        ).fetchone()["count"]
    )
    local_upload_count = int(
        conn.execute(
            """
            SELECT
                (SELECT COUNT(*) FROM response_files WHERE relative_path LIKE '/uploads/%' OR relative_path LIKE 'http%/uploads/%')
                +
                (SELECT COUNT(*) FROM respondent_profile_records WHERE image_relative_path LIKE '/uploads/%' OR image_relative_path LIKE 'http%/uploads/%')
                AS count
            """
        ).fetchone()["count"]
    )
    external_image_count = uploads_count + profile_images_count - local_upload_count
    db_size = DB_PATH.stat().st_size if DB_PATH.exists() else 0
    upload_file_count = len([path for path in UPLOADS_DIR.iterdir() if path.is_file()]) if UPLOADS_DIR.exists() else 0
    backups = list_backup_archives(limit=5)
    return {
        "localUrl": f"http://127.0.0.1:{server_port}",
        "publicUrl": public_base_url,
        "publicBaseUrlSource": public_base_url_source(),
        "configuredPublicBaseUrl": str(CONFIG.get("public_base_url") or ""),
        "defaultPasswordInUse": CONFIG["admin_password_sha256"] == sha256_hex(DEFAULT_PASSWORD),
        "databasePath": str(DB_PATH),
        "uploadsPath": str(UPLOADS_DIR),
        "backupsPath": str(BACKUPS_DIR),
        "databaseSizeBytes": db_size,
        "uploadFileCount": upload_file_count,
        "localImageCount": local_upload_count,
        "externalImageCount": external_image_count,
        "responseImageCount": uploads_count,
        "profileImageCount": profile_images_count,
        "backupCount": len(list(BACKUPS_DIR.glob('bijiris-backup-*.tar.gz'))) if BACKUPS_DIR.exists() else 0,
        "latestBackup": backups[0] if backups else None,
        "publicUrlIsTemporary": "trycloudflare.com" in public_base_url,
    }

def should_sync_local_upload_path(current_path: Any, stored_name: Any) -> bool:
    path = str(current_path or "").strip()
    raw_stored_name = str(stored_name or "").strip()
    if not raw_stored_name:
        return False
    if not path or path == raw_stored_name:
        return True
    if path.startswith("/uploads/"):
        return True
    parsed = urllib.parse.urlparse(path)
    if parsed.scheme and parsed.netloc:
        return parsed.path == f"/uploads/{urllib.parse.quote(raw_stored_name)}"
    return False


def response_file_default_title(row: sqlite3.Row) -> str:
    label = str(row["label"] or "").strip()
    if label:
        form_title = str(row["form_title"] or "").strip()
        return f"{form_title} / {label}" if form_title else label
    return str(row["original_name"] or "").strip() or "アンケート画像"


def response_file_default_date(row: sqlite3.Row) -> str:
    return str(row["response_created_at"] or "")[:10]


def fetch_response_file_annotation_row(conn: sqlite3.Connection, response_file_id: int) -> sqlite3.Row | None:
    return conn.execute(
        """
        SELECT *
        FROM response_file_annotations
        WHERE response_file_id = ?
        """,
        (response_file_id,),
    ).fetchone()


def fetch_response_file_annotations_map(
    conn: sqlite3.Connection,
    response_file_ids: list[int],
) -> dict[int, sqlite3.Row]:
    unique_ids = list(dict.fromkeys(response_file_id for response_file_id in response_file_ids if response_file_id))
    if not unique_ids:
        return {}
    placeholders = ",".join("?" for _ in unique_ids)
    rows = conn.execute(
        f"""
        SELECT *
        FROM response_file_annotations
        WHERE response_file_id IN ({placeholders})
        """,
        unique_ids,
    ).fetchall()
    return {int(row["response_file_id"]): row for row in rows}


def respondent_profile_record_payload(row: sqlite3.Row) -> dict[str, Any]:
    image = None
    if row["image_relative_path"]:
        image = {
            "originalName": row["image_original_name"] or "",
            "storedName": row["image_stored_name"] or "",
            "mimeType": row["image_mime_type"] or "",
            "url": resolve_public_image_url(row["image_relative_path"], row["image_stored_name"]),
            "previewUrl": resolve_public_image_preview_url(row["image_relative_path"], row["image_stored_name"]),
        }
    return {
        "id": row["id"],
        "recordId": row["id"],
        "sourceType": "manual",
        "title": row["title"] or "",
        "date": row["entry_date"] or "",
        "memo": row["memo"] or "",
        "createdAt": row["created_at"],
        "updatedAt": row["updated_at"],
        "image": image,
        "editable": True,
        "deletable": True,
        "sourceLabel": "管理者追加",
    }


def response_file_record_payload(row: sqlite3.Row, annotation: sqlite3.Row | None = None) -> dict[str, Any]:
    title = str(annotation["title"]).strip() if annotation and annotation["title"] else response_file_default_title(row)
    entry_date = str(annotation["entry_date"]).strip() if annotation and annotation["entry_date"] else response_file_default_date(row)
    memo = str(annotation["memo"]).strip() if annotation and annotation["memo"] else ""
    created_at = annotation["created_at"] if annotation and annotation["created_at"] else row["response_created_at"]
    updated_at = annotation["updated_at"] if annotation and annotation["updated_at"] else row["response_created_at"]
    source_label = str(row["label"] or "").strip() or "アンケート画像"
    return {
        "id": f"response:{row['id']}",
        "recordId": row["id"],
        "sourceType": "response",
        "title": title,
        "date": entry_date,
        "memo": memo,
        "createdAt": created_at,
        "updatedAt": updated_at,
        "image": response_file_image_payload(row),
        "editable": True,
        "deletable": False,
        "sourceLabel": source_label,
        "formTitle": row["form_title"] or "",
        "responseId": row["response_id"],
        "responseCreatedAt": row["response_created_at"],
    }


def empty_respondent_profile_summary() -> dict[str, Any]:
    return {
        "profileDate": "",
        "profileTitle": "",
        "profileMemo": "",
        "profileImage": None,
        "profileRecordCount": 0,
    }


def respondent_profile_payload(row: sqlite3.Row | None) -> dict[str, Any]:
    if not row:
        return empty_respondent_profile_summary()
    if "stored_name" in row.keys():
        record = response_file_record_payload(row)
    else:
        record = respondent_profile_record_payload(row)
    return {
        "profileDate": record["date"],
        "profileTitle": record["title"],
        "profileMemo": record["memo"],
        "profileImage": record["image"],
        "profileRecordCount": 1,
    }


def fetch_respondent_profiles_map(
    conn: sqlite3.Connection,
    respondent_ids: list[str],
) -> dict[str, dict[str, Any]]:
    unique_ids = list(dict.fromkeys(respondent_id for respondent_id in respondent_ids if respondent_id))
    if not unique_ids:
        return {}
    result: dict[str, dict[str, Any]] = {}
    for respondent_id in unique_ids:
        records = fetch_respondent_profile_records(conn, respondent_id)
        summary = empty_respondent_profile_summary()
        if records:
            summary.update(
                {
                    "profileDate": records[0]["date"],
                    "profileTitle": records[0]["title"],
                    "profileMemo": records[0]["memo"],
                    "profileImage": records[0]["image"],
                    "profileRecordCount": len(records),
                }
            )
        result[respondent_id] = summary
    return result


def fetch_respondent_profile_row(conn: sqlite3.Connection, respondent_id: str) -> sqlite3.Row | None:
    return conn.execute(
        """
        SELECT *
        FROM respondent_profile_records
        WHERE respondent_id = ?
        ORDER BY entry_date DESC, updated_at DESC, id DESC
        LIMIT 1
        """,
        (respondent_id,),
    ).fetchone()


def fetch_respondent_profile_record_row(
    conn: sqlite3.Connection,
    respondent_id: str,
    record_id: int,
) -> sqlite3.Row | None:
    return conn.execute(
        """
        SELECT *
        FROM respondent_profile_records
        WHERE respondent_id = ? AND id = ?
        """,
        (respondent_id, record_id),
    ).fetchone()


def fetch_respondent_profile_records(conn: sqlite3.Connection, respondent_id: str) -> list[dict[str, Any]]:
    manual_rows = conn.execute(
        """
        SELECT *
        FROM respondent_profile_records
        WHERE respondent_id = ?
        ORDER BY entry_date DESC, updated_at DESC, id DESC
        """,
        (respondent_id,),
    ).fetchall()
    response_rows = conn.execute(
        """
        SELECT
            response_files.*,
            responses.respondent_id AS respondent_id,
            responses.respondent_name AS respondent_name,
            responses.created_at AS response_created_at,
            forms.title AS form_title
        FROM response_files
        JOIN responses ON responses.id = response_files.response_id
        LEFT JOIN forms ON forms.id = responses.form_id
        WHERE responses.respondent_id = ?
        ORDER BY responses.created_at DESC, response_files.id DESC
        """,
        (respondent_id,),
    ).fetchall()
    annotations_map = fetch_response_file_annotations_map(conn, [int(row["id"]) for row in response_rows])
    records = [respondent_profile_record_payload(row) for row in manual_rows]
    records.extend(
        response_file_record_payload(row, annotations_map.get(int(row["id"])))
        for row in response_rows
    )
    records.sort(key=lambda record: (record["date"] or "", record["updatedAt"] or "", str(record["id"])), reverse=True)
    return records


def measurement_image_link_payload(record: dict[str, Any]) -> dict[str, str]:
    image = record.get("image") or {}
    title = str(record.get("title") or "").strip()
    source_label = str(record.get("sourceLabel") or "").strip()
    original_name = str(image.get("originalName") or "").strip()
    label = title or source_label or original_name or "画像"
    return {
        "label": label,
        "url": str(image.get("url") or ""),
        "previewUrl": str(image.get("previewUrl") or ""),
        "title": title,
        "sourceLabel": source_label,
        "originalName": original_name,
    }


def attach_measurement_image_links(
    conn: sqlite3.Connection,
    measurement_records: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    respondent_ids = list(
        dict.fromkeys(str(record.get("respondentId") or "") for record in measurement_records if record.get("respondentId"))
    )
    if not respondent_ids:
        return measurement_records

    image_map: dict[str, dict[str, list[dict[str, str]]]] = {}
    for respondent_id in respondent_ids:
        by_date: dict[str, list[dict[str, str]]] = {}
        for record in fetch_respondent_profile_records(conn, respondent_id):
            entry_date = str(record.get("date") or "").strip()
            image = record.get("image") or {}
            if not entry_date or not image.get("url"):
                continue
            by_date.setdefault(entry_date, []).append(measurement_image_link_payload(record))
        image_map[respondent_id] = by_date

    for record in measurement_records:
        respondent_id = str(record.get("respondentId") or "")
        entry_date = str(record.get("date") or "").strip()
        record["imageLinks"] = image_map.get(respondent_id, {}).get(entry_date, [])
    return measurement_records


def respondent_measurement_payload(row: sqlite3.Row) -> dict[str, Any]:
    return {
        "id": row["id"],
        "recordId": row["id"],
        "respondentId": row["respondent_id"],
        "respondentName": row["respondent_name"],
        "date": row["entry_date"] or "",
        "category": row["measurement_category"] or "",
        "waist": float(row["waist"]),
        "hip": float(row["hip"]),
        "thigh": float(row["thigh"]),
        "waistLabel": format_measurement_value(row["waist"]),
        "hipLabel": format_measurement_value(row["hip"]),
        "thighLabel": format_measurement_value(row["thigh"]),
        "createdAt": row["created_at"],
        "updatedAt": row["updated_at"],
        "editable": True,
        "deletable": True,
    }


def empty_measurement_summary() -> dict[str, Any]:
    return {
        "measurementCount": 0,
        "latestMeasurementDate": "",
        "latestMeasurements": None,
    }


def fetch_respondent_measurement_row(conn: sqlite3.Connection, respondent_id: str, record_id: int) -> sqlite3.Row | None:
    return conn.execute(
        """
        SELECT *
        FROM respondent_measurements
        WHERE respondent_id = ? AND id = ?
        """,
        (respondent_id, record_id),
    ).fetchone()


def fetch_respondent_measurement_row_by_date(
    conn: sqlite3.Connection,
    respondent_id: str,
    entry_date: str,
) -> sqlite3.Row | None:
    return conn.execute(
        """
        SELECT *
        FROM respondent_measurements
        WHERE respondent_id = ? AND entry_date = ?
        ORDER BY updated_at DESC, id DESC
        LIMIT 1
        """,
        (respondent_id, entry_date),
    ).fetchone()


def fetch_respondent_measurement_records(conn: sqlite3.Connection, respondent_id: str) -> list[dict[str, Any]]:
    rows = conn.execute(
        """
        SELECT *
        FROM respondent_measurements
        WHERE respondent_id = ?
        ORDER BY entry_date ASC, updated_at ASC, id ASC
        """,
        (respondent_id,),
    ).fetchall()
    return attach_measurement_image_links(conn, [respondent_measurement_payload(row) for row in rows])


def list_measurement_records(
    conn: sqlite3.Connection,
    *,
    respondent_id: str | None = None,
    respondent_name: str = "",
    query: str = "",
    limit: int = 500,
) -> list[dict[str, Any]]:
    clauses: list[str] = []
    params: list[Any] = []
    if respondent_id:
        clauses.append("respondent_id = ?")
        params.append(respondent_id)
    if respondent_name:
        clauses.append("respondent_name LIKE ?")
        params.append(f"%{respondent_name}%")
    if query:
        pattern = f"%{query}%"
        clauses.append("(respondent_id LIKE ? OR respondent_name LIKE ? OR entry_date LIKE ? OR measurement_category LIKE ?)")
        params.extend([pattern, pattern, pattern, pattern])
    sql = """
        SELECT *
        FROM respondent_measurements
    """
    if clauses:
        sql += " WHERE " + " AND ".join(clauses)
    sql += " ORDER BY entry_date ASC, updated_at ASC, id ASC LIMIT ?"
    params.append(limit)
    rows = conn.execute(sql, params).fetchall()
    return attach_measurement_image_links(conn, [respondent_measurement_payload(row) for row in rows])


def fetch_respondent_measurements_map(
    conn: sqlite3.Connection,
    respondent_ids: list[str],
) -> dict[str, dict[str, Any]]:
    unique_ids = list(dict.fromkeys(respondent_id for respondent_id in respondent_ids if respondent_id))
    if not unique_ids:
        return {}
    result: dict[str, dict[str, Any]] = {}
    for respondent_id in unique_ids:
        records = fetch_respondent_measurement_records(conn, respondent_id)
        summary = empty_measurement_summary()
        if records:
            latest = records[-1]
            summary.update(
                {
                    "measurementCount": len(records),
                    "latestMeasurementDate": latest["date"],
                    "latestMeasurements": {
                        "category": latest["category"],
                        "waist": latest["waist"],
                        "hip": latest["hip"],
                        "thigh": latest["thigh"],
                        "waistLabel": latest["waistLabel"],
                        "hipLabel": latest["hipLabel"],
                        "thighLabel": latest["thighLabel"],
                    },
                }
            )
        result[respondent_id] = summary
    return result


def create_respondent_measurement_record(
    conn: sqlite3.Connection,
    respondent_id: str,
    respondent_name: str,
    *,
    entry_date: str,
    category: Any = "",
    waist: Any,
    hip: Any,
    thigh: Any,
) -> dict[str, Any]:
    entry_date_value = validate_profile_date(entry_date)
    if not entry_date_value:
        raise ValueError("計測日を入力してください。")
    category_value = validate_measurement_category(category)
    waist_value = validate_measurement_value(waist, "ウエスト")
    hip_value = validate_measurement_value(hip, "ヒップ")
    thigh_value = validate_measurement_value(thigh, "太もも")
    timestamp = now_iso()
    with conn:
        cursor = conn.execute(
            """
            INSERT INTO respondent_measurements (
                respondent_id, respondent_name, entry_date, measurement_category, waist, hip, thigh, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                respondent_id,
                respondent_name,
                entry_date_value,
                category_value,
                waist_value,
                hip_value,
                thigh_value,
                timestamp,
                timestamp,
            ),
        )
    row = fetch_respondent_measurement_row(conn, respondent_id, int(cursor.lastrowid))
    if not row:
        raise RuntimeError("計測記録の保存に失敗しました。")
    return respondent_measurement_payload(row)


def update_respondent_measurement_record(
    conn: sqlite3.Connection,
    respondent_id: str,
    record_id: int,
    *,
    entry_date: Any,
    category: Any = "",
    waist: Any,
    hip: Any,
    thigh: Any,
) -> dict[str, Any]:
    row = fetch_respondent_measurement_row(conn, respondent_id, record_id)
    if not row:
        raise LookupError("計測記録が見つかりません。")
    entry_date_value = validate_profile_date(entry_date) or str(row["entry_date"] or "")
    if not entry_date_value:
        raise ValueError("計測日を入力してください。")
    if category == "":
        category_value = ""
    elif str(category or "").strip():
        category_value = validate_measurement_category(category)
    else:
        category_value = str(row["measurement_category"] or "")
    waist_value = validate_measurement_value(waist if str(waist or "").strip() else row["waist"], "ウエスト")
    hip_value = validate_measurement_value(hip if str(hip or "").strip() else row["hip"], "ヒップ")
    thigh_value = validate_measurement_value(thigh if str(thigh or "").strip() else row["thigh"], "太もも")
    with conn:
        conn.execute(
            """
            UPDATE respondent_measurements
            SET entry_date = ?, measurement_category = ?, waist = ?, hip = ?, thigh = ?, updated_at = ?
            WHERE respondent_id = ? AND id = ?
            """,
            (entry_date_value, category_value, waist_value, hip_value, thigh_value, now_iso(), respondent_id, record_id),
        )
    updated = fetch_respondent_measurement_row(conn, respondent_id, record_id)
    if not updated:
        raise RuntimeError("計測記録の更新に失敗しました。")
    return respondent_measurement_payload(updated)


def delete_respondent_measurement_record(conn: sqlite3.Connection, respondent_id: str, record_id: int) -> dict[str, Any]:
    row = fetch_respondent_measurement_row(conn, respondent_id, record_id)
    if not row:
        raise LookupError("計測記録が見つかりません。")
    with conn:
        deleted_count = conn.execute(
            "DELETE FROM respondent_measurements WHERE respondent_id = ? AND id = ?",
            (respondent_id, record_id),
        ).rowcount
    return {"deletedCount": deleted_count}


def delete_all_respondent_measurements(conn: sqlite3.Connection, respondent_id: str) -> None:
    with conn:
        conn.execute("DELETE FROM respondent_measurements WHERE respondent_id = ?", (respondent_id,))


def save_respondent_measurement_record(
    conn: sqlite3.Connection,
    respondent_id: str,
    respondent_name: str,
    *,
    entry_date: str,
    category: Any = "",
    waist: Any,
    hip: Any,
    thigh: Any,
) -> tuple[str, dict[str, Any]]:
    normalized_entry_date = validate_profile_date(entry_date)
    if not normalized_entry_date:
        raise ValueError("計測日を入力してください。")
    category_value = validate_measurement_category(category)
    waist_value = validate_measurement_value(waist, "ウエスト")
    hip_value = validate_measurement_value(hip, "ヒップ")
    thigh_value = validate_measurement_value(thigh, "太もも")
    existing = fetch_respondent_measurement_row_by_date(conn, respondent_id, normalized_entry_date)
    if existing:
        if (
            str(existing["measurement_category"] or "") == category_value
            and
            float(existing["waist"]) == waist_value
            and float(existing["hip"]) == hip_value
            and float(existing["thigh"]) == thigh_value
        ):
            return "skipped", respondent_measurement_payload(existing)
        updated = update_respondent_measurement_record(
            conn,
            respondent_id,
            int(existing["id"]),
            entry_date=normalized_entry_date,
            category=category_value,
            waist=waist_value,
            hip=hip_value,
            thigh=thigh_value,
        )
        return "updated", updated
    created = create_respondent_measurement_record(
        conn,
        respondent_id,
        respondent_name,
        entry_date=normalized_entry_date,
        category=category_value,
        waist=waist_value,
        hip=hip_value,
        thigh=thigh_value,
    )
    return "created", created


def move_respondent_measurements(
    conn: sqlite3.Connection,
    old_respondent_id: str,
    new_respondent_id: str,
    respondent_name: str,
) -> None:
    timestamp = now_iso()
    with conn:
        if old_respondent_id != new_respondent_id:
            conn.execute(
                """
                UPDATE respondent_measurements
                SET respondent_id = ?, respondent_name = ?, updated_at = ?
                WHERE respondent_id = ?
                """,
                (new_respondent_id, respondent_name, timestamp, old_respondent_id),
            )
        else:
            conn.execute(
                """
                UPDATE respondent_measurements
                SET respondent_name = ?, updated_at = ?
                WHERE respondent_id = ?
                """,
                (respondent_name, timestamp, old_respondent_id),
            )


def fetch_latest_ticket_stamp_auto_map(
    conn: sqlite3.Connection,
    respondent_ids: list[str],
) -> dict[str, dict[str, Any]]:
    unique_ids = list(dict.fromkeys(respondent_id for respondent_id in respondent_ids if respondent_id))
    if not unique_ids:
        return {}
    placeholders = ",".join("?" for _ in unique_ids)
    rows = conn.execute(
        f"""
        SELECT respondent_id, ticket_visit_count, response_created_at
        FROM (
            SELECT
                responses.respondent_id AS respondent_id,
                response_answers.value AS ticket_visit_count,
                responses.created_at AS response_created_at,
                ROW_NUMBER() OVER (
                    PARTITION BY responses.respondent_id
                    ORDER BY responses.created_at DESC, responses.id DESC, response_answers.id DESC
                ) AS row_number
            FROM responses
            JOIN forms ON forms.id = responses.form_id
            JOIN response_answers ON response_answers.response_id = responses.id
            WHERE
                responses.respondent_id IN ({placeholders})
                AND forms.slug = ?
                AND response_answers.field_key = ?
        ) latest_ticket_stamps
        WHERE row_number = 1
        """,
        [*unique_ids, TREATMENT_SURVEY_FORM_SLUG, TICKET_VISIT_COUNT_FIELD_KEY],
    ).fetchall()
    return {
        str(row["respondent_id"]): {
            "autoCountRaw": parse_ticket_visit_count(row["ticket_visit_count"]),
            "autoCountAt": str(row["response_created_at"] or ""),
        }
        for row in rows
    }


def fetch_latest_ticket_sheet_map(
    conn: sqlite3.Connection,
    respondent_ids: list[str],
) -> dict[str, dict[str, Any]]:
    unique_ids = list(dict.fromkeys(respondent_id for respondent_id in respondent_ids if respondent_id))
    if not unique_ids:
        return {}
    placeholders = ",".join("?" for _ in unique_ids)
    auto_map = fetch_latest_ticket_stamp_auto_map(conn, unique_ids)
    rows = conn.execute(
        f"""
        SELECT
            respondents.respondent_id AS respondent_id,
            respondents.ticket_sheet_manual_value AS ticket_sheet_manual_value,
            respondents.current_ticket_book_type AS current_ticket_book_type,
            respondents.current_ticket_stamp_count AS current_ticket_stamp_count,
            respondents.current_ticket_stamp_manual_enabled AS current_ticket_stamp_manual_enabled,
            respondents.updated_at AS response_created_at
        FROM respondents
        WHERE respondents.respondent_id IN ({placeholders})
        """,
        unique_ids,
    ).fetchall()
    result: dict[str, dict[str, Any]] = {}
    for row in rows:
        respondent_id = str(row["respondent_id"])
        manual_value = str(row["ticket_sheet_manual_value"] or "").strip()
        current_ticket_book_type = str(row["current_ticket_book_type"] or "").strip()
        current_ticket_stamp_manual_value = int(row["current_ticket_stamp_count"] or 0)
        current_ticket_stamp_manual_enabled = int(row["current_ticket_stamp_manual_enabled"] or 0) == 1
        auto_info = auto_map.get(respondent_id, {"autoCountRaw": 0, "autoCountAt": ""})
        auto_count = effective_ticket_stamp_count(current_ticket_book_type, int(auto_info["autoCountRaw"]))
        manual_count = effective_ticket_stamp_count(current_ticket_book_type, current_ticket_stamp_manual_value)
        result[respondent_id] = {
            "latestTicketSheet": manual_value,
            "latestTicketSheetManualValue": manual_value,
            "latestTicketSheetAt": str(row["response_created_at"] or "") if manual_value else "",
            "currentTicketBookType": current_ticket_book_type,
            "currentTicketStampCount": manual_count if current_ticket_stamp_manual_enabled else auto_count,
            "currentTicketStampAutoValue": auto_count,
            "currentTicketStampManualValue": current_ticket_stamp_manual_value if current_ticket_stamp_manual_enabled else "",
            "currentTicketStampManualEnabled": current_ticket_stamp_manual_enabled,
            "currentTicketStampMax": (
                ticket_book_stamp_display_max(current_ticket_book_type) if current_ticket_book_type else 0
            ),
            "currentTicketStampAt": (
                str(row["response_created_at"] or "")
                if current_ticket_stamp_manual_enabled
                else str(auto_info["autoCountAt"] or "")
            ),
        }
    return result


def fetch_answers_map(conn: sqlite3.Connection, response_ids: list[int]) -> dict[int, list[dict[str, Any]]]:
    if not response_ids:
        return {}
    placeholders = ",".join("?" for _ in response_ids)
    rows = conn.execute(
        f"""
        SELECT * FROM response_answers
        WHERE response_id IN ({placeholders})
        ORDER BY id ASC
        """,
        response_ids,
    ).fetchall()
    result: dict[int, list[dict[str, Any]]] = {response_id: [] for response_id in response_ids}
    for row in rows:
        result.setdefault(row["response_id"], []).append(
            {
                "label": row["label"],
                "key": row["field_key"],
                "value": row["value"],
            }
        )
    return result


def list_responses(
    conn: sqlite3.Connection,
    form_id: int | None = None,
    respondent_query: str = "",
    category: str = "",
    limit: int = 60,
) -> list[dict[str, Any]]:
    clauses: list[str] = []
    params: list[Any] = []
    if form_id is not None:
        clauses.append("responses.form_id = ?")
        params.append(form_id)
    if category:
        clauses.append("responses.category = ?")
        params.append(category)
    if respondent_query:
        pattern = f"%{respondent_query}%"
        clauses.append(
            "(responses.respondent_id LIKE ? OR responses.respondent_name LIKE ? OR responses.respondent_email LIKE ?)"
        )
        params.extend([pattern, pattern, pattern])

    sql = """
        SELECT
            responses.*,
            forms.title AS form_title,
            forms.slug AS form_slug
        FROM responses
        JOIN forms ON forms.id = responses.form_id
    """
    if clauses:
        sql += " WHERE " + " AND ".join(clauses)
    sql += " ORDER BY responses.created_at DESC LIMIT ?"
    params.append(limit)
    rows = conn.execute(sql, params).fetchall()
    response_ids = [row["id"] for row in rows]
    files_map = fetch_files_map(conn, response_ids)
    responses: list[dict[str, Any]] = []
    for row in rows:
        responses.append(
            {
                "id": row["id"],
                "formId": row["form_id"],
                "formTitle": row["form_title"],
                "formSlug": row["form_slug"],
                "respondentId": row["respondent_id"],
                "respondentName": row["respondent_name"],
                "respondentEmail": row["respondent_email"],
                "category": row["category"],
                "notes": row["notes"],
                "createdAt": row["created_at"],
                "files": files_map.get(row["id"], []),
            }
        )
    return responses


def fetch_response_detail(conn: sqlite3.Connection, response_id: int) -> dict[str, Any] | None:
    row = conn.execute(
        """
        SELECT responses.*, forms.title AS form_title, forms.slug AS form_slug
        FROM responses
        JOIN forms ON forms.id = responses.form_id
        WHERE responses.id = ?
        """,
        (response_id,),
    ).fetchone()
    if not row:
        return None
    files = fetch_files_map(conn, [response_id]).get(response_id, [])
    answers = fetch_answers_map(conn, [response_id]).get(response_id, [])
    return {
        "response": {
            "id": row["id"],
            "formId": row["form_id"],
            "formTitle": row["form_title"],
            "formSlug": row["form_slug"],
            "respondentId": row["respondent_id"],
            "respondentName": row["respondent_name"],
            "respondentEmail": row["respondent_email"],
            "category": row["category"],
            "notes": row["notes"],
            "createdAt": row["created_at"],
        },
        "answers": answers,
        "files": files,
    }


def category_summary(conn: sqlite3.Connection, form_id: int | None, respondent_query: str = "") -> list[dict[str, Any]]:
    clauses: list[str] = []
    params: list[Any] = []
    if form_id is not None:
        clauses.append("form_id = ?")
        params.append(form_id)
    if respondent_query:
        pattern = f"%{respondent_query}%"
        clauses.append("(respondent_id LIKE ? OR respondent_name LIKE ? OR respondent_email LIKE ?)")
        params.extend([pattern, pattern, pattern])
    sql = "SELECT category, COUNT(*) AS count FROM responses"
    if clauses:
        sql += " WHERE " + " AND ".join(clauses)
    sql += " GROUP BY category ORDER BY count DESC, category ASC"
    return [dict(row) for row in conn.execute(sql, params).fetchall()]


def respondent_summary(
    conn: sqlite3.Connection,
    form_id: int | None = None,
    query: str = "",
    limit: int = 100,
) -> list[dict[str, Any]]:
    clauses: list[str] = []
    params: list[Any] = []
    join_clause = "LEFT JOIN responses ON responses.respondent_id = respondents.respondent_id"
    if form_id is not None:
        join_clause = (
            "LEFT JOIN responses ON responses.respondent_id = respondents.respondent_id "
            "AND responses.form_id = ?"
        )
        params.append(form_id)
    if query:
        pattern = f"%{query}%"
        clauses.append("(respondents.respondent_id LIKE ? OR respondents.respondent_name LIKE ?)")
        params.extend([pattern, pattern])

    sql = f"""
        SELECT
            respondents.respondent_id AS respondent_id,
            respondents.respondent_name AS respondent_name,
            COUNT(responses.id) AS response_count,
            MAX(responses.created_at) AS last_response_at
        FROM respondents
        {join_clause}
    """
    if clauses:
        sql += " WHERE " + " AND ".join(clauses)
    sql += " GROUP BY respondents.respondent_id, respondents.respondent_name"
    if form_id is not None:
        sql += " HAVING COUNT(responses.id) > 0"
    sql += " ORDER BY (MAX(responses.created_at) IS NOT NULL) DESC, MAX(responses.created_at) DESC, respondents.updated_at DESC LIMIT ?"
    params.append(limit)
    rows = conn.execute(sql, params).fetchall()
    profiles_map = fetch_respondent_profiles_map(conn, [row["respondent_id"] for row in rows])
    measurements_map = fetch_respondent_measurements_map(conn, [str(row["respondent_id"]) for row in rows])
    ticket_sheet_map = fetch_latest_ticket_sheet_map(conn, [str(row["respondent_id"]) for row in rows])
    return [
        {
            "respondentId": row["respondent_id"],
            "respondentName": row["respondent_name"],
            "respondentEmail": "",
            "responseCount": row["response_count"],
            "lastResponseAt": row["last_response_at"] or "",
            **profiles_map.get(row["respondent_id"], empty_respondent_profile_summary()),
            **measurements_map.get(str(row["respondent_id"]), empty_measurement_summary()),
            **ticket_sheet_map.get(
                str(row["respondent_id"]),
                {
                    "latestTicketSheet": "",
                    "latestTicketSheetManualValue": "",
                    "latestTicketSheetAt": "",
                    "currentTicketBookType": "",
                    "currentTicketStampCount": 0,
                    "currentTicketStampAutoValue": 0,
                    "currentTicketStampManualValue": "",
                    "currentTicketStampManualEnabled": False,
                    "currentTicketStampMax": 0,
                    "currentTicketStampAt": "",
                },
            ),
        }
        for row in rows
    ]


def respondent_overview(conn: sqlite3.Connection, respondent_id: str, form_id: int | None = None) -> dict[str, Any] | None:
    row = fetch_respondent_row(conn, respondent_id)
    if not row:
        return None
    params: list[Any] = [respondent_id]
    sql = """
        SELECT COUNT(*) AS response_count, MAX(created_at) AS last_response_at
        FROM responses
        WHERE respondent_id = ?
    """
    if form_id is not None:
        sql += " AND form_id = ?"
        params.append(form_id)
    stats = conn.execute(sql, params).fetchone()
    profile = fetch_respondent_profiles_map(conn, [respondent_id]).get(respondent_id, empty_respondent_profile_summary())
    measurement_summary = fetch_respondent_measurements_map(conn, [respondent_id]).get(
        respondent_id,
        empty_measurement_summary(),
    )
    latest_ticket_sheet = fetch_latest_ticket_sheet_map(conn, [respondent_id]).get(
        respondent_id,
        {
            "latestTicketSheet": "",
            "latestTicketSheetManualValue": "",
            "latestTicketSheetAt": "",
            "currentTicketBookType": "",
            "currentTicketStampCount": 0,
            "currentTicketStampAutoValue": 0,
            "currentTicketStampManualValue": "",
            "currentTicketStampManualEnabled": False,
            "currentTicketStampMax": 0,
            "currentTicketStampAt": "",
        },
    )
    return {
        "respondentId": respondent_id,
        "respondentName": row["respondent_name"],
        "respondentEmail": "",
        "responseCount": stats["response_count"] if stats else 0,
        "lastResponseAt": stats["last_response_at"] if stats and stats["last_response_at"] else "",
        **profile,
        **measurement_summary,
        **latest_ticket_sheet,
    }


def respondent_history(conn: sqlite3.Connection, respondent_id: str, form_id: int | None = None) -> list[dict[str, Any]]:
    clauses = ["respondent_id = ?"]
    params: list[Any] = [respondent_id]
    if form_id is not None:
        clauses.append("form_id = ?")
        params.append(form_id)
    rows = conn.execute(
        f"""
        SELECT
            responses.*,
            forms.title AS form_title
        FROM responses
        JOIN forms ON forms.id = responses.form_id
        WHERE {' AND '.join(clauses)}
        ORDER BY responses.created_at ASC, responses.id ASC
        """,
        params,
    ).fetchall()
    response_ids = [row["id"] for row in rows]
    files_map = fetch_files_map(conn, response_ids)
    answers_map = fetch_answers_map(conn, response_ids)
    profiles_map = fetch_respondent_profiles_map(conn, [row["respondent_id"] for row in rows])
    history: list[dict[str, Any]] = []
    for row in rows:
        history.append(
            {
                "id": row["id"],
                "formId": row["form_id"],
                "formTitle": row["form_title"],
                "respondentId": row["respondent_id"],
                "respondentName": row["respondent_name"],
                "respondentEmail": row["respondent_email"],
                "category": row["category"],
                "notes": row["notes"],
                "createdAt": row["created_at"],
                "files": files_map.get(row["id"], []),
                "answers": answers_map.get(row["id"], []),
                **profiles_map.get(row["respondent_id"], empty_respondent_profile_summary()),
            }
        )
    return history


def public_history_file_payload(item: dict[str, Any]) -> dict[str, Any]:
    return {
        "label": str(item.get("label") or ""),
        "originalName": str(item.get("originalName") or ""),
        "mimeType": str(item.get("mimeType") or ""),
        "size": int(item.get("size") or 0),
    }


def public_respondent_history_payload(item: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": item["id"],
        "formId": item["formId"],
        "formTitle": item["formTitle"],
        "category": item["category"],
        "createdAt": item["createdAt"],
        "answers": [
            {
                "label": str(answer.get("label") or ""),
                "value": str(answer.get("value") or ""),
            }
            for answer in item.get("answers", [])
        ],
        "files": [public_history_file_payload(file_item) for file_item in item.get("files", [])],
    }


def parse_optional_form_id(value: Any) -> int | None:
    if value in {None, ""}:
        return None
    if isinstance(value, bool):
        raise ValueError("フォームの指定が不正です。")
    if isinstance(value, int):
        return value
    text = str(value).strip()
    if not text:
        return None
    if not text.isdigit():
        raise ValueError("フォームの指定が不正です。")
    return int(text)


def validate_profile_date(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        return dt.date.fromisoformat(text).isoformat()
    except ValueError as exc:
        raise ValueError("日付は YYYY-MM-DD 形式で入力してください。") from exc


def validate_measurement_value(value: Any, label: str) -> float:
    text = unicodedata.normalize("NFKC", str(value or "")).strip()
    if not text:
        raise ValueError(f"{label}は必須です。")
    normalized = text.replace("cm", "").replace("CM", "").strip()
    match = re.search(r"-?\d+(?:\.\d+)?", normalized)
    if match:
        normalized = match.group(0)
    try:
        amount = round(float(normalized), 1)
    except ValueError as exc:
        raise ValueError(f"{label}は数値で入力してください。") from exc
    if amount <= 0:
        raise ValueError(f"{label}は0より大きい数値で入力してください。")
    return amount


def validate_measurement_category(value: Any) -> str:
    text = normalize_respondent_name(str(value or ""))
    if not text:
        return ""
    if text not in MEASUREMENT_CATEGORIES:
        allowed = " / ".join(MEASUREMENT_CATEGORIES)
        raise ValueError(f"カテゴリは次から選択してください。{allowed}")
    return text


def format_measurement_value(value: Any) -> str:
    if value in {None, ""}:
        return ""
    amount = float(value)
    if amount.is_integer():
        return str(int(amount))
    return f"{amount:.1f}"


def normalize_import_column_name(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).strip().lower()
    if not text:
        return ""
    return re.sub(r"[\s_\-‐‑‒–—―/／・,，、:：;；()（）\[\]［］{}｛｝]+", "", text)


def resolve_measurement_import_url(sheet_url: str) -> tuple[str, str]:
    raw_url = str(sheet_url or "").strip()
    if not raw_url:
        raise ValueError("スプレッドシートURLを入力してください。")

    parsed = urllib.parse.urlparse(raw_url)
    if parsed.scheme not in {"http", "https"}:
        raise ValueError("http または https のURLを入力してください。")

    host = parsed.netloc.lower()
    if "docs.google.com" not in host:
        return raw_url, raw_url

    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", parsed.path)
    if not match:
        return raw_url, raw_url

    query = urllib.parse.parse_qs(parsed.query)
    fragment = urllib.parse.parse_qs(parsed.fragment)
    gid = (query.get("gid") or fragment.get("gid") or ["0"])[-1]
    export_url = f"https://docs.google.com/spreadsheets/d/{match.group(1)}/export?format=csv&gid={urllib.parse.quote(gid)}"
    return raw_url, export_url


def download_measurement_import_text(source_url: str) -> str:
    request = urllib.request.Request(
        source_url,
        headers={
            "User-Agent": "Bijiris/1.0",
            "Accept": "text/csv,text/plain,application/vnd.ms-excel;q=0.9,*/*;q=0.8",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            raw = response.read()
            content_type = response.headers.get_content_type()
            charset = response.headers.get_content_charset() or "utf-8"
    except urllib.error.HTTPError as exc:
        if exc.code in {401, 403}:
            raise ValueError("スプレッドシートを取得できません。リンク共有をオンにしてください。") from exc
        raise ValueError("スプレッドシートの取得に失敗しました。URLを確認してください。") from exc
    except urllib.error.URLError as exc:
        raise ValueError("スプレッドシートへ接続できませんでした。URLを確認してください。") from exc

    if not raw:
        raise ValueError("スプレッドシートにデータがありません。")

    for encoding in [charset, "utf-8-sig", "utf-8", "cp932"]:
        try:
            text = raw.decode(encoding)
            break
        except (LookupError, UnicodeDecodeError):
            continue
    else:
        raise ValueError("スプレッドシートの文字コードを読み取れませんでした。")

    if content_type == "text/html" or text.lstrip().startswith("<!DOCTYPE html") or text.lstrip().startswith("<html"):
        raise ValueError("スプレッドシートをCSVとして取得できませんでした。リンク共有をオンにしてください。")
    return text


def parse_measurement_import_rows(csv_text: str) -> tuple[list[str], list[dict[str, str]]]:
    matrix = parse_measurement_import_matrix(csv_text)
    if not matrix:
        raise ValueError("スプレッドシートのヘッダー行を読み取れませんでした。")

    headers: list[str] | None = None
    data_start_index = 0
    scan_limit = min(len(matrix), 20)
    for index in range(scan_limit):
        candidate = matrix[index]
        if not any(candidate):
            continue
        try:
            detect_measurement_import_columns(candidate)
        except ValueError:
            continue
        headers = candidate
        data_start_index = index + 1
        break

    if headers is None:
        headers = next((row for row in matrix if any(row)), None)
        data_start_index = matrix.index(headers) + 1 if headers else 0

    if not headers or not any(str(header or "").strip() for header in headers):
        raise ValueError("スプレッドシートのヘッダー行を読み取れませんでした。")

    rows = []
    for offset, raw_row in enumerate(matrix[data_start_index:], start=data_start_index + 1):
        values = {
            str(key or "").strip(): str(value or "").strip()
            for key, value in zip_longest(headers, raw_row, fillvalue="")
            if str(key or "").strip()
        }
        values["__row_number__"] = str(offset + 1)
        rows.append(values)
    return headers, rows


def parse_measurement_import_matrix(csv_text: str) -> list[list[str]]:
    sample = csv_text[:2048]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;")
    except csv.Error:
        dialect = csv.excel
    return [
        [str(cell or "").strip() for cell in row]
        for row in csv.reader(io.StringIO(csv_text), dialect=dialect)
    ]


def detect_measurement_import_columns(headers: list[str]) -> dict[str, str]:
    remaining = {header: normalize_import_column_name(header) for header in headers if str(header or "").strip()}
    if not remaining:
        raise ValueError("スプレッドシートの列名が見つかりません。")

    def pick_column(candidates: tuple[str, ...], *, required: bool) -> str:
        normalized_candidates = [normalize_import_column_name(candidate) for candidate in candidates]
        for candidate in normalized_candidates:
            for header, normalized in remaining.items():
                if normalized == candidate:
                    return header
        for candidate in normalized_candidates:
            for header, normalized in remaining.items():
                if candidate and candidate in normalized:
                    return header
        if required:
            raise ValueError(f"列名を判定できませんでした: {' / '.join(candidates)}")
        return ""

    return {
        "name": pick_column(MEASUREMENT_IMPORT_COLUMN_HINTS["name"], required=False),
        "date": pick_column(MEASUREMENT_IMPORT_COLUMN_HINTS["date"], required=True),
        "waist": pick_column(MEASUREMENT_IMPORT_COLUMN_HINTS["waist"], required=True),
        "hip": pick_column(MEASUREMENT_IMPORT_COLUMN_HINTS["hip"], required=True),
        "thigh": pick_column(MEASUREMENT_IMPORT_COLUMN_HINTS["thigh"], required=True),
    }


def measurement_site_key(value: Any) -> str:
    normalized = normalize_import_column_name(value)
    if not normalized:
        return ""
    for site_key in ("waist", "hip", "thigh"):
        for candidate in MEASUREMENT_IMPORT_COLUMN_HINTS[site_key]:
            candidate_normalized = normalize_import_column_name(candidate)
            if normalized == candidate_normalized or candidate_normalized in normalized:
                return site_key
    return ""


def detect_measurement_block_axes(matrix: list[list[str]]) -> tuple[int, int, int] | None:
    normalized_site_hints = [normalize_import_column_name(value) for value in MEASUREMENT_SITE_COLUMN_HINTS]
    normalized_name_hints = [normalize_import_column_name(value) for value in MEASUREMENT_IMPORT_COLUMN_HINTS["name"]]
    scan_limit = min(len(matrix), 30)
    for row_index in range(scan_limit):
        row = matrix[row_index]
        if not any(row):
            continue
        normalized_row = [normalize_import_column_name(cell) for cell in row]
        name_col = next(
            (
                index
                for index, value in enumerate(normalized_row)
                if value and any(value == hint or hint in value for hint in normalized_name_hints)
            ),
            None,
        )
        site_col = next(
            (
                index
                for index, value in enumerate(normalized_row)
                if value and any(value == hint or hint in value for hint in normalized_site_hints)
            ),
            None,
        )
        if name_col is not None and site_col is not None:
            return row_index, name_col, site_col
    return None


def normalize_measurement_import_date(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "")).strip()
    if not text:
        raise ValueError("計測日が空です。")
    text = text.replace("／", "/").replace(".", "-")
    if "T" in text:
        text = text.split("T", 1)[0].strip()
    if " " in text:
        text = text.split(" ", 1)[0].strip()
    match = re.fullmatch(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日", text)
    if match:
        return dt.date(int(match.group(1)), int(match.group(2)), int(match.group(3))).isoformat()
    match = re.fullmatch(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", text)
    if match:
        return dt.date(int(match.group(1)), int(match.group(2)), int(match.group(3))).isoformat()
    try:
        return dt.date.fromisoformat(text).isoformat()
    except ValueError as exc:
        raise ValueError("計測日は YYYY-MM-DD 形式で読める値にしてください。") from exc


def prepare_measurement_import_records_from_rows(
    rows: list[dict[str, str]],
    columns: dict[str, str],
    target_name: str,
) -> tuple[list[dict[str, Any]], int]:
    target_match_key = respondent_name_match_key(target_name)
    matched_row_count = 0
    prepared_records: list[dict[str, Any]] = []
    errors: list[str] = []

    for row in rows:
        row_number = int(str(row.get("__row_number__", "0")) or "0")
        if not any(str(value or "").strip() for key, value in row.items() if key != "__row_number__"):
            continue
        if columns["name"]:
            row_name = normalize_respondent_name(row.get(columns["name"], ""))
            if not row_name or respondent_name_match_key(row_name) != target_match_key:
                continue
        candidate_values = [row.get(columns[key], "") for key in ("date", "waist", "hip", "thigh")]
        if not any(str(value or "").strip() for value in candidate_values):
            continue
        matched_row_count += 1
        try:
            prepared_records.append(
                {
                    "entry_date": normalize_measurement_import_date(row.get(columns["date"], "")),
                    "waist": validate_measurement_value(row.get(columns["waist"], ""), "ウエスト"),
                    "hip": validate_measurement_value(row.get(columns["hip"], ""), "ヒップ"),
                    "thigh": validate_measurement_value(row.get(columns["thigh"], ""), "太もも"),
                }
            )
        except ValueError as exc:
            errors.append(f"{row_number or '?'}行目: {exc}")

    if not matched_row_count:
        if columns["name"]:
            raise ValueError("一致するお名前の計測記録が見つかりませんでした。")
        raise ValueError("取り込める計測記録が見つかりませんでした。")

    if errors:
        preview = " / ".join(errors[:3])
        suffix = "" if len(errors) <= 3 else f" ほか{len(errors) - 3}件"
        raise ValueError(f"取込できない行があります。{preview}{suffix}")

    return prepared_records, matched_row_count


def prepare_measurement_import_records_from_blocks(
    csv_text: str,
    target_name: str,
) -> tuple[list[dict[str, Any]], int, dict[str, str]]:
    matrix = parse_measurement_import_matrix(csv_text)
    axes = detect_measurement_block_axes(matrix)
    if not axes:
        raise ValueError("列名を判定できませんでした。")

    header_row_index, name_col, site_col = axes
    data_col_start = max(name_col, site_col) + 1
    current_dates: dict[int, str] = {}
    current_name = ""
    target_match_key = respondent_name_match_key(target_name)
    grouped_records: dict[str, dict[str, Any]] = {}
    matched_names = 0
    errors: list[str] = []

    for row_index, row in enumerate(matrix[header_row_index + 1 :], start=header_row_index + 2):
        row_name = normalize_respondent_name(row[name_col] if name_col < len(row) else "")
        site_key = measurement_site_key(row[site_col] if site_col < len(row) else "")
        date_cells: dict[int, str] = {}
        for col_index in range(data_col_start, len(row)):
            cell = row[col_index]
            if not cell:
                continue
            try:
                date_cells[col_index] = normalize_measurement_import_date(cell)
            except ValueError:
                continue

        if row_name:
            current_name = row_name

        if date_cells and not site_key:
            current_dates = date_cells
            continue

        if not site_key or respondent_name_match_key(current_name) != target_match_key:
            continue

        matched_names += 1
        if not current_dates:
            errors.append(f"{row_index}行目: 計測日の行が見つかりません。")
            continue

        for col_index, entry_date in current_dates.items():
            if col_index >= len(row):
                continue
            value = str(row[col_index] or "").strip()
            if not value:
                continue
            try:
                numeric_value = validate_measurement_value(
                    value,
                    {"waist": "ウエスト", "hip": "ヒップ", "thigh": "太もも"}[site_key],
                )
            except ValueError as exc:
                errors.append(f"{row_index}行目 ({entry_date}): {exc}")
                continue
            grouped_records.setdefault(entry_date, {"entry_date": entry_date})[site_key] = numeric_value

    if not matched_names:
        raise ValueError("一致するお名前の計測記録が見つかりませんでした。")

    prepared_records: list[dict[str, Any]] = []
    for entry_date in sorted(grouped_records.keys()):
        record = grouped_records[entry_date]
        missing = [label for key, label in (("waist", "ウエスト"), ("hip", "ヒップ"), ("thigh", "太もも")) if key not in record]
        if missing:
            errors.append(f"{entry_date}: {' / '.join(missing)} が不足しています。")
            continue
        prepared_records.append(
            {
                "entry_date": entry_date,
                "waist": record["waist"],
                "hip": record["hip"],
                "thigh": record["thigh"],
            }
        )

    if errors:
        preview = " / ".join(errors[:3])
        suffix = "" if len(errors) <= 3 else f" ほか{len(errors) - 3}件"
        raise ValueError(f"取込できない行があります。{preview}{suffix}")

    if not prepared_records:
        raise ValueError("取り込める計測記録が見つかりませんでした。")

    return prepared_records, len(prepared_records), {"name": "お名前", "measurementSite": "計測箇所", "dates": "横並び日付"}


def import_respondent_measurements_from_sheet(
    conn: sqlite3.Connection,
    respondent_id: str,
    respondent_name: str,
    *,
    sheet_url: str,
) -> dict[str, Any]:
    original_url, download_url = resolve_measurement_import_url(sheet_url)
    csv_text = download_measurement_import_text(download_url)
    target_name = normalize_respondent_name(respondent_name)
    prepared_records: list[dict[str, Any]]
    matched_row_count: int
    columns: dict[str, str]
    tabular_error: ValueError | None = None
    try:
        headers, rows = parse_measurement_import_rows(csv_text)
        columns = detect_measurement_import_columns(headers)
        prepared_records, matched_row_count = prepare_measurement_import_records_from_rows(rows, columns, target_name)
    except ValueError as exc:
        tabular_error = exc
        prepared_records, matched_row_count, columns = prepare_measurement_import_records_from_blocks(csv_text, target_name)

    imported = 0
    updated = 0
    skipped = 0
    records: list[dict[str, Any]] = []
    for record in prepared_records:
        action, payload = save_respondent_measurement_record(
            conn,
            respondent_id,
            respondent_name,
            entry_date=record["entry_date"],
            waist=record["waist"],
            hip=record["hip"],
            thigh=record["thigh"],
        )
        if action == "created":
            imported += 1
        elif action == "updated":
            updated += 1
        else:
            skipped += 1
        records.append(payload)

    return {
        "importedCount": imported,
        "updatedCount": updated,
        "skippedCount": skipped,
        "matchedRowCount": matched_row_count,
        "records": records,
        "sourceUrl": original_url,
        "downloadUrl": download_url,
        "columnMapping": columns,
        "layout": "tabular" if tabular_error is None else "block",
    }


def count_responses_for_respondent(conn: sqlite3.Connection, respondent_id: str) -> int:
    row = conn.execute(
        """
        SELECT COUNT(*) AS response_count
        FROM responses
        WHERE respondent_id = ?
        """,
        (respondent_id,),
    ).fetchone()
    return int(row["response_count"]) if row else 0


def delete_uploaded_file(stored_name: str) -> None:
    if not stored_name:
        return
    target = read_upload_target(stored_name)
    if not target:
        return
    try:
        target.unlink()
    except FileNotFoundError:
        return


def create_respondent_profile_record(
    conn: sqlite3.Connection,
    respondent_id: str,
    respondent_name: str,
    *,
    title: str,
    entry_date: str,
    memo: str,
    uploaded_image: dict[str, Any],
) -> dict[str, Any]:
    title_value = str(title or "").strip()
    if not title_value:
        raise ValueError("タイトルは必須です。")

    entry_date_value = validate_profile_date(entry_date)
    if not entry_date_value:
        raise ValueError("日付は必須です。")

    if not uploaded_image:
        raise ValueError("画像は必須です。")

    timestamp = now_iso()
    stored_name = f"profile_record_{secrets.token_hex(10)}{uploaded_image['extension']}"
    (UPLOADS_DIR / stored_name).write_bytes(uploaded_image["bytes"])

    with conn:
        cursor = conn.execute(
            """
            INSERT INTO respondent_profile_records (
                respondent_id, respondent_name, title, entry_date, memo,
                image_original_name, image_stored_name, image_mime_type, image_relative_path,
                created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                respondent_id,
                respondent_name,
                title_value,
                entry_date_value,
                str(memo or "").strip(),
                uploaded_image["original_name"],
                stored_name,
                uploaded_image["mime_type"],
                local_upload_path(stored_name),
                timestamp,
                timestamp,
            ),
        )

    row = fetch_respondent_profile_record_row(conn, respondent_id, int(cursor.lastrowid))
    if not row:
        raise RuntimeError("画像記録の保存に失敗しました。")
    return respondent_profile_record_payload(row)


def fetch_response_file_record_row(
    conn: sqlite3.Connection,
    respondent_id: str,
    response_file_id: int,
) -> sqlite3.Row | None:
    return conn.execute(
        """
        SELECT
            response_files.*,
            responses.respondent_id AS respondent_id,
            responses.respondent_name AS respondent_name,
            responses.created_at AS response_created_at,
            forms.title AS form_title
        FROM response_files
        JOIN responses ON responses.id = response_files.response_id
        LEFT JOIN forms ON forms.id = responses.form_id
        WHERE responses.respondent_id = ? AND response_files.id = ?
        """,
        (respondent_id, response_file_id),
    ).fetchone()


def update_respondent_profile_record(
    conn: sqlite3.Connection,
    respondent_id: str,
    source_type: str,
    record_id: int,
    *,
    title: str,
    entry_date: str,
    memo: str,
) -> dict[str, Any]:
    timestamp = now_iso()
    title_value = str(title or "").strip()
    memo_value = str(memo or "").strip()

    if source_type == "manual":
        row = fetch_respondent_profile_record_row(conn, respondent_id, record_id)
        if not row:
            raise LookupError("画像記録が見つかりません。")
        date_value = validate_profile_date(entry_date) or row["entry_date"]
        title_final = title_value or row["title"] or "無題"
        if not date_value:
            raise ValueError("日付は必須です。")
        with conn:
            conn.execute(
                """
                UPDATE respondent_profile_records
                SET title = ?, entry_date = ?, memo = ?, updated_at = ?
                WHERE respondent_id = ? AND id = ?
                """,
                (title_final, date_value, memo_value, timestamp, respondent_id, record_id),
            )
        updated = fetch_respondent_profile_record_row(conn, respondent_id, record_id)
        if not updated:
            raise RuntimeError("画像記録の更新に失敗しました。")
        return respondent_profile_record_payload(updated)

    if source_type == "response":
        row = fetch_response_file_record_row(conn, respondent_id, record_id)
        if not row:
            raise LookupError("画像記録が見つかりません。")
        annotation = fetch_response_file_annotation_row(conn, record_id)
        date_value = validate_profile_date(entry_date) or (
            annotation["entry_date"] if annotation and annotation["entry_date"] else response_file_default_date(row)
        )
        title_final = title_value or (
            str(annotation["title"]).strip() if annotation and annotation["title"] else response_file_default_title(row)
        )
        if not date_value:
            raise ValueError("日付は必須です。")
        with conn:
            if annotation:
                conn.execute(
                    """
                    UPDATE response_file_annotations
                    SET title = ?, entry_date = ?, memo = ?, updated_at = ?
                    WHERE response_file_id = ?
                    """,
                    (title_final, date_value, memo_value, timestamp, record_id),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO response_file_annotations (
                        response_file_id, title, entry_date, memo, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (record_id, title_final, date_value, memo_value, timestamp, timestamp),
                )
        updated_annotation = fetch_response_file_annotation_row(conn, record_id)
        return response_file_record_payload(row, updated_annotation)

    raise ValueError("画像記録の種類が不正です。")


def delete_respondent_profile_record(conn: sqlite3.Connection, respondent_id: str, record_id: int) -> dict[str, Any]:
    row = fetch_respondent_profile_record_row(conn, respondent_id, record_id)
    if not row:
        raise LookupError("画像記録が見つかりません。")

    with conn:
        conn.execute(
            "DELETE FROM respondent_profile_records WHERE respondent_id = ? AND id = ?",
            (respondent_id, record_id),
        )

    delete_uploaded_file(row["image_stored_name"])
    return {"deletedId": record_id}


def delete_all_respondent_profile_records(conn: sqlite3.Connection, respondent_id: str) -> None:
    rows = conn.execute(
        """
        SELECT image_stored_name
        FROM respondent_profile_records
        WHERE respondent_id = ?
        """,
        (respondent_id,),
    ).fetchall()
    with conn:
        conn.execute("DELETE FROM respondent_profile_records WHERE respondent_id = ?", (respondent_id,))
    for row in rows:
        delete_uploaded_file(row["image_stored_name"])


def move_respondent_profile(
    conn: sqlite3.Connection,
    old_respondent_id: str,
    new_respondent_id: str,
    respondent_name: str,
) -> None:
    timestamp = now_iso()
    with conn:
        if old_respondent_id == new_respondent_id:
            conn.execute(
                """
                UPDATE respondent_profile_records
                SET respondent_name = ?, updated_at = ?
                WHERE respondent_id = ?
                """,
                (respondent_name, timestamp, old_respondent_id),
            )
            return

        conn.execute(
            """
            UPDATE respondent_profile_records
            SET respondent_id = ?, respondent_name = ?, updated_at = ?
            WHERE respondent_id = ?
            """,
            (new_respondent_id, respondent_name, timestamp, old_respondent_id),
        )


def rename_respondent(
    conn: sqlite3.Connection,
    respondent_id: str,
    new_name: str,
    *,
    form_id: int | None = None,
) -> dict[str, Any]:
    normalized_name = normalize_respondent_name(new_name)
    if not normalized_name:
        raise ValueError("お名前は必須です。")

    clauses = ["respondent_id = ?"]
    params: list[Any] = [respondent_id]
    if form_id is not None:
        clauses.append("form_id = ?")
        params.append(form_id)

    row = conn.execute(
        f"""
        SELECT COUNT(*) AS response_count
        FROM responses
        WHERE {' AND '.join(clauses)}
        """,
        params,
    ).fetchone()
    existing_registry = fetch_respondent_row(conn, respondent_id)
    existing_manual_ticket_sheet = str(existing_registry["ticket_sheet_manual_value"] or "").strip() if existing_registry else ""
    existing_ticket_book_type = str(existing_registry["current_ticket_book_type"] or "").strip() if existing_registry else ""
    existing_ticket_stamp_count = int(existing_registry["current_ticket_stamp_count"] or 0) if existing_registry else 0
    existing_ticket_stamp_manual_enabled = (
        int(existing_registry["current_ticket_stamp_manual_enabled"] or 0) == 1 if existing_registry else False
    )
    if (not row or row["response_count"] <= 0) and not (existing_registry and form_id is None):
        raise LookupError("回答者が見つかりません。")

    updated_respondent_id = respondent_name_key(normalized_name)
    with conn:
        if row and row["response_count"] > 0:
            conn.execute(
                f"""
                UPDATE responses
                SET respondent_id = ?, respondent_name = ?
                WHERE {' AND '.join(clauses)}
                """,
                [updated_respondent_id, normalized_name, *params],
            )
        if existing_registry:
            conn.execute("DELETE FROM respondents WHERE respondent_id = ?", (respondent_id,))
        ensure_respondent_registry(
            conn,
            normalized_name,
            respondent_id=updated_respondent_id,
            ticket_sheet_manual_value=existing_manual_ticket_sheet,
            current_ticket_book_type=existing_ticket_book_type,
            current_ticket_stamp_count=existing_ticket_stamp_count,
            current_ticket_stamp_manual_enabled=existing_ticket_stamp_manual_enabled,
        )
    move_respondent_profile(conn, respondent_id, updated_respondent_id, normalized_name)
    move_respondent_measurements(conn, respondent_id, updated_respondent_id, normalized_name)

    return {
        "respondentId": updated_respondent_id,
        "respondentName": normalized_name,
        "updatedCount": row["response_count"] if row else 0,
    }


def delete_respondent(
    conn: sqlite3.Connection,
    respondent_id: str,
    *,
    form_id: int | None = None,
) -> dict[str, Any]:
    clauses = ["responses.respondent_id = ?"]
    params: list[Any] = [respondent_id]
    if form_id is not None:
        clauses.append("responses.form_id = ?")
        params.append(form_id)

    file_rows = conn.execute(
        f"""
        SELECT response_files.stored_name
        FROM response_files
        JOIN responses ON responses.id = response_files.response_id
        WHERE {' AND '.join(clauses)}
        """,
        params,
    ).fetchall()
    stored_names = [row["stored_name"] for row in file_rows]

    with conn:
        deleted_count = conn.execute(
            f"""
            DELETE FROM responses
            WHERE {' AND '.join(clause.replace('responses.', '') for clause in clauses)}
            """,
            params,
        ).rowcount

    if deleted_count <= 0:
        respondent = fetch_respondent_row(conn, respondent_id)
        if not respondent or form_id is not None:
            raise LookupError("回答者が見つかりません。")
        delete_all_respondent_profile_records(conn, respondent_id)
        delete_all_respondent_measurements(conn, respondent_id)
        with conn:
            conn.execute("DELETE FROM respondents WHERE respondent_id = ?", (respondent_id,))
            conn.execute("DELETE FROM respondent_profiles WHERE respondent_id = ?", (respondent_id,))
        return {"deletedCount": 0, "registryDeleted": True}

    for stored_name in stored_names:
        delete_uploaded_file(stored_name)

    if count_responses_for_respondent(conn, respondent_id) <= 0:
        delete_all_respondent_profile_records(conn, respondent_id)
        delete_all_respondent_measurements(conn, respondent_id)
        with conn:
            conn.execute("DELETE FROM respondents WHERE respondent_id = ?", (respondent_id,))
            conn.execute("DELETE FROM respondent_profiles WHERE respondent_id = ?", (respondent_id,))

    return {"deletedCount": deleted_count}


def validate_uploaded_images(
    files: list[cgi.FieldStorage],
    *,
    label: str = "画像",
    required: bool = False,
) -> list[dict[str, Any]]:
    prepared: list[dict[str, Any]] = []
    total_size = 0
    for file_item in files:
        if not getattr(file_item, "filename", None):
            continue
        original_name = Path(file_item.filename).name
        file_bytes = file_item.file.read()
        if not file_bytes:
            continue
        size = len(file_bytes)
        if size > FILE_SIZE_LIMIT:
            raise ValueError(f"{original_name} は10MB以下にしてください。")
        total_size += size
        if total_size > TOTAL_UPLOAD_LIMIT:
            raise ValueError("アップロードの合計サイズは30MB以下にしてください。")
        mime_type = file_item.type or mimetypes.guess_type(original_name)[0] or "application/octet-stream"
        extension = Path(original_name).suffix.lower()
        if not (mime_type.startswith("image/") or extension in ALLOWED_IMAGE_EXTENSIONS):
            raise ValueError(f"{original_name} は画像ファイルのみアップロードできます。")
        prepared.append(
            {
                "original_name": original_name,
                "bytes": file_bytes,
                "mime_type": mime_type if mime_type.startswith("image/") else "image/octet-stream",
                "extension": extension or mimetypes.guess_extension(mime_type) or ".bin",
                "size": size,
            }
        )
    if required and not prepared:
        raise ValueError(f"{label} は1件以上アップロードしてください。")
    return prepared


class BijirisServer(ThreadingHTTPServer):
    daemon_threads = True


def split_host_port(host_header: str) -> tuple[str, str]:
    value = str(host_header or "").strip()
    if not value:
        return "", ""
    if value.startswith("[") and "]" in value:
        host, _, remainder = value[1:].partition("]")
        port = remainder[1:] if remainder.startswith(":") else ""
        return host, port
    if value.count(":") == 1:
        host, port = value.rsplit(":", 1)
        if port.isdigit():
            return host, port
    return value, ""


def is_loopback_or_unspecified_host(host: str) -> bool:
    normalized = str(host or "").strip().lower()
    return (
        not normalized
        or normalized == "localhost"
        or normalized == "0.0.0.0"
        or normalized == "::"
        or normalized == "::1"
        or normalized.startswith("127.")
    )


def score_ipv4_address(ip: str) -> tuple[int, int]:
    if ip.startswith("192.168."):
        return (0, 0)
    if ip.startswith("10."):
        return (0, 1)
    if re.match(r"^172\.(1[6-9]|2\d|3[0-1])\.", ip):
        return (0, 2)
    if ip.startswith("169.254."):
        return (2, 0)
    return (1, 0)


def detect_lan_ipv4() -> str | None:
    candidates: set[str] = set()

    try:
        udp_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            udp_socket.connect(("192.0.2.1", 80))
            detected = udp_socket.getsockname()[0]
            if detected and not is_loopback_or_unspecified_host(detected):
                candidates.add(detected)
        finally:
            udp_socket.close()
    except OSError:
        pass

    try:
        for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET, socket.SOCK_STREAM):
            detected = str(info[4][0] or "").strip()
            if detected and not is_loopback_or_unspecified_host(detected):
                candidates.add(detected)
    except OSError:
        pass

    try:
        output = subprocess.check_output(["ifconfig"], text=True, stderr=subprocess.DEVNULL)
        blocks = re.split(r"\n(?=\S)", output)
        for block in blocks:
            if "status: active" not in block:
                continue
            for match in re.finditer(r"inet (\d+\.\d+\.\d+\.\d+)", block):
                detected = match.group(1)
                if detected and not is_loopback_or_unspecified_host(detected):
                    candidates.add(detected)
    except (OSError, subprocess.SubprocessError):
        pass

    if not candidates:
        return None
    return sorted(candidates, key=score_ipv4_address)[0]


def preferred_public_base_url(
    request_host_header: str,
    *,
    fallback_port: int,
    scheme: str = "http",
) -> str:
    configured = os.environ.get(PUBLIC_BASE_URL_ENV, "").strip()
    if configured:
        return configured.rstrip("/")
    render_external = os.environ.get(RENDER_EXTERNAL_URL_ENV, "").strip()
    if render_external:
        return render_external.rstrip("/")
    configured = str(CONFIG.get("public_base_url") or "").strip()
    if configured:
        return configured.rstrip("/")

    request_host, request_port = split_host_port(request_host_header)
    port = request_port or str(fallback_port)
    if request_host and not is_loopback_or_unspecified_host(request_host):
        return f"{scheme}://{request_host}:{port}" if port else f"{scheme}://{request_host}"

    lan_ip = detect_lan_ipv4()
    if lan_ip:
        return f"{scheme}://{lan_ip}:{port}" if port else f"{scheme}://{lan_ip}"

    fallback_host = request_host or "127.0.0.1"
    return f"{scheme}://{fallback_host}:{port}" if port else f"{scheme}://{fallback_host}"


class BijirisHandler(BaseHTTPRequestHandler):
    server_version = "Bijiris/1.0"

    def log_message(self, format: str, *args: Any) -> None:
        sys.stdout.write("%s - - [%s] %s\n" % (self.address_string(), self.log_date_time_string(), format % args))

    def write_response_body(self, raw: bytes) -> None:
        try:
            self.wfile.write(raw)
        except (BrokenPipeError, ConnectionResetError):
            # Browsers may cancel image/file requests during navigation; ignore the disconnect.
            return

    def send_json(self, payload: Any, status: int = 200, headers: list[tuple[str, str]] | None = None) -> None:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Cache-Control", "no-store")
        if headers:
            for key, value in headers:
                self.send_header(key, value)
        self.end_headers()
        self.write_response_body(raw)

    def send_text(self, text: str, status: int = 200, content_type: str = "text/plain; charset=utf-8") -> None:
        raw = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.write_response_body(raw)

    def serve_file(
        self,
        path: Path,
        content_type: str | None = None,
        *,
        cache_control: str = "no-store, no-cache, must-revalidate",
    ) -> None:
        raw = path.read_bytes()
        mime = content_type or mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        if mime.startswith("text/") or mime in {"application/javascript", "application/json"}:
            mime = f"{mime}; charset=utf-8"
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Cache-Control", cache_control)
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.end_headers()
        self.write_response_body(raw)

    def not_found(self, message: str = "見つかりません。") -> None:
        self.send_json({"error": message}, status=404)

    def bad_request(self, message: str) -> None:
        self.send_json({"error": message}, status=400)

    def server_error(self, message: str = "サーバーエラーが発生しました。") -> None:
        self.send_json({"error": message}, status=500)

    def get_session_user(self) -> str | None:
        cookies = parse_cookie_header(self.headers.get("Cookie"))
        return verify_session(cookies.get(SESSION_COOKIE))

    def require_admin(self) -> str | None:
        user = self.get_session_user()
        if not user:
            self.send_json({"error": "認証が必要です。"}, status=401)
            return None
        return user

    def get_query_params(self) -> dict[str, str]:
        parsed = urllib.parse.urlparse(self.path)
        return {key: values[-1] for key, values in urllib.parse.parse_qs(parsed.query).items()}

    def get_preferred_public_base_url(self) -> str:
        scheme = self.headers.get("X-Forwarded-Proto", "http").split(",")[0].strip() or "http"
        return preferred_public_base_url(
            self.headers.get("Host", ""),
            fallback_port=int(self.server.server_port),
            scheme=scheme,
        )

    def parse_form_data(self) -> cgi.FieldStorage:
        return cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                "CONTENT_LENGTH": self.headers.get("Content-Length", "0"),
            },
            keep_blank_values=True,
        )

    def serve_admin(self) -> None:
        asset = read_public_asset("admin.html")
        if not asset:
            self.not_found()
            return
        self.serve_file(asset, "text/html; charset=utf-8")

    def serve_respondent(self) -> None:
        asset = read_public_asset("respondent.html")
        if not asset:
            self.not_found()
            return
        self.serve_file(asset, "text/html; charset=utf-8")

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if path in {"/", "/admin", "/admin/"}:
            self.serve_admin()
            return
        if path in {"/f", "/f/"} or path.startswith("/f/"):
            self.serve_respondent()
            return
        if path == "/sw.js":
            asset = read_public_asset("sw.js")
            if not asset:
                self.not_found()
                return
            self.serve_file(asset, "application/javascript; charset=utf-8")
            return
        if path.startswith("/assets/"):
            asset = read_public_asset(path.replace("/assets/", "", 1))
            if not asset:
                self.not_found()
                return
            self.serve_file(asset)
            return
        if path.startswith("/uploads/"):
            if not self.require_admin():
                return
            asset = read_upload_target(path.replace("/uploads/", "", 1))
            if not asset:
                self.not_found()
                return
            self.serve_file(asset)
            return
        if path == "/health":
            self.send_json({"ok": True, "timestamp": now_iso()})
            return

        segments = [segment for segment in path.split("/") if segment]
        try:
            if segments[:2] == ["api", "public"] and len(segments) == 3 and segments[2] == "forms":
                self.handle_public_forms_list()
                return
            if segments[:3] == ["api", "admin", "image-proxy"] and len(segments) == 3:
                self.handle_admin_image_proxy()
                return
            if segments[:3] == ["api", "public", "respondents"] and len(segments) == 4 and segments[3] == "history":
                self.handle_public_respondent_history()
                return
            if segments[:3] == ["api", "public", "forms"] and len(segments) == 4:
                self.handle_public_form_get(segments[3])
                return
            if segments[:3] == ["api", "admin", "bootstrap"] and len(segments) == 3:
                self.handle_admin_bootstrap()
                return
            if segments[:3] == ["api", "admin", "operations"] and len(segments) == 4 and segments[3] == "status":
                self.handle_admin_operations_status()
                return
            if segments[:3] == ["api", "admin", "backups"] and len(segments) == 5 and segments[4] == "download":
                self.handle_admin_backup_download(urllib.parse.unquote(segments[3]))
                return
            if segments[:3] == ["api", "admin", "forms"] and len(segments) == 3:
                self.handle_admin_forms_list()
                return
            if segments[:3] == ["api", "admin", "forms"] and len(segments) == 5 and segments[4] == "responses":
                self.handle_admin_form_responses(int(segments[3]))
                return
            if segments[:3] == ["api", "admin", "responses"] and len(segments) == 4:
                self.handle_admin_response_detail(int(segments[3]))
                return
            if segments[:3] == ["api", "admin", "respondents"] and len(segments) == 3:
                self.handle_admin_respondents()
                return
            if segments[:3] == ["api", "admin", "measurements"] and len(segments) == 3:
                self.handle_admin_measurements()
                return
            if segments[:3] == ["api", "admin", "respondents"] and len(segments) == 5 and segments[4] == "history":
                self.handle_admin_respondent_history(urllib.parse.unquote(segments[3]))
                return
        except ValueError:
            self.bad_request("URLが不正です。")
            return
        except Exception:
            self.server_error()
            raise

        self.not_found()

    def handle_admin_image_proxy(self) -> None:
        if not self.require_admin():
            return
        params = self.get_query_params()
        source = str(params.get("src", "")).strip()
        if not source:
            self.bad_request("画像URLが必要です。")
            return

        stored_name = upload_stored_name_from_source(source)
        if stored_name:
            asset = read_upload_target(stored_name)
            if not asset:
                self.not_found("画像が見つかりません。")
                return
            self.serve_file(asset, cache_control="private, no-store, no-cache, must-revalidate")
            return

        try:
            payload, content_type = download_remote_image(source)
        except PermissionError as exc:
            self.send_json({"error": str(exc)}, status=403)
            return
        except ValueError as exc:
            self.bad_request(str(exc))
            return

        mime = content_type if content_type.startswith("image/") else "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "private, no-store, no-cache, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.end_headers()
        self.write_response_body(payload)

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        segments = [segment for segment in path.split("/") if segment]
        try:
            if path == "/api/admin/login":
                self.handle_admin_login()
                return
            if path == "/api/admin/logout":
                self.handle_admin_logout()
                return
            if path == "/api/admin/settings/public-base-url":
                self.handle_admin_update_public_base_url()
                return
            if path == "/api/admin/settings/password":
                self.handle_admin_update_password()
                return
            if path == "/api/admin/backups/create":
                self.handle_admin_backup_create()
                return
            if path == "/api/admin/forms":
                self.handle_admin_create_form()
                return
            if path == "/api/admin/respondents/create":
                self.handle_admin_create_respondent()
                return
            if segments[:3] == ["api", "admin", "forms"] and len(segments) == 5 and segments[4] == "toggle":
                self.handle_admin_toggle_form(int(segments[3]))
                return
            if segments[:3] == ["api", "admin", "respondents"] and len(segments) == 5 and segments[4] == "rename":
                self.handle_admin_respondent_rename(urllib.parse.unquote(segments[3]))
                return
            if segments[:3] == ["api", "admin", "respondents"] and len(segments) == 5 and segments[4] == "delete":
                self.handle_admin_respondent_delete(urllib.parse.unquote(segments[3]))
                return
            if segments[:3] == ["api", "admin", "respondents"] and len(segments) == 5 and segments[4] == "profile":
                self.handle_admin_respondent_profile_update(urllib.parse.unquote(segments[3]))
                return
            if segments[:3] == ["api", "admin", "respondents"] and len(segments) == 5 and segments[4] == "profile-records":
                self.handle_admin_respondent_profile_record_create(urllib.parse.unquote(segments[3]))
                return
            if segments[:3] == ["api", "admin", "respondents"] and len(segments) == 5 and segments[4] == "measurements":
                self.handle_admin_respondent_measurement_create(urllib.parse.unquote(segments[3]))
                return
            if (
                segments[:3] == ["api", "admin", "respondents"]
                and len(segments) == 6
                and segments[4] == "measurements"
                and segments[5] == "import-sheet"
            ):
                self.handle_admin_respondent_measurement_import(urllib.parse.unquote(segments[3]))
                return
            if (
                segments[:3] == ["api", "admin", "respondents"]
                and len(segments) == 8
                and segments[4] == "profile-records"
                and segments[7] == "update"
            ):
                self.handle_admin_respondent_profile_record_update(
                    urllib.parse.unquote(segments[3]),
                    segments[5],
                    int(segments[6]),
                )
                return
            if (
                segments[:3] == ["api", "admin", "respondents"]
                and len(segments) == 8
                and segments[4] == "profile-records"
                and segments[7] == "delete"
            ):
                self.handle_admin_respondent_profile_record_delete(
                    urllib.parse.unquote(segments[3]),
                    segments[5],
                    int(segments[6]),
                )
                return
            if (
                segments[:3] == ["api", "admin", "respondents"]
                and len(segments) == 7
                and segments[4] == "measurements"
                and segments[6] == "update"
            ):
                self.handle_admin_respondent_measurement_update(
                    urllib.parse.unquote(segments[3]),
                    int(segments[5]),
                )
                return
            if (
                segments[:3] == ["api", "admin", "respondents"]
                and len(segments) == 7
                and segments[4] == "measurements"
                and segments[6] == "delete"
            ):
                self.handle_admin_respondent_measurement_delete(
                    urllib.parse.unquote(segments[3]),
                    int(segments[5]),
                )
                return
            if segments[:3] == ["api", "public", "forms"] and len(segments) == 5 and segments[4] == "submit":
                self.handle_public_form_submit(segments[3])
                return
        except ValueError:
            self.bad_request("入力内容が不正です。")
            return
        except Exception:
            self.server_error()
            raise

        self.not_found()

    def do_PUT(self) -> None:
        segments = [segment for segment in urllib.parse.urlparse(self.path).path.split("/") if segment]
        try:
            if segments[:3] == ["api", "admin", "forms"] and len(segments) == 4:
                self.handle_admin_update_form(int(segments[3]))
                return
        except ValueError:
            self.bad_request("URLが不正です。")
            return
        except Exception:
            self.server_error()
            raise
        self.not_found()

    def handle_public_form_get(self, slug: str) -> None:
        conn = get_connection()
        try:
            form = fetch_form_by_slug(conn, slug, include_inactive=False)
            if not form:
                self.not_found("公開中のフォームが見つかりません。")
                return
            self.send_json({"form": form})
        finally:
            conn.close()

    def handle_public_forms_list(self) -> None:
        conn = get_connection()
        try:
            self.send_json({"forms": fetch_public_forms(conn)})
        finally:
            conn.close()

    def handle_public_respondent_history(self) -> None:
        params = self.get_query_params()
        respondent_name = normalize_respondent_name(params.get("name", ""))
        if not respondent_name:
            self.bad_request("お名前を入力してください。")
            return
        conn = get_connection()
        try:
            respondent = fetch_respondent_row_by_name(conn, respondent_name)
            if not respondent:
                self.not_found("一致する回答履歴が見つかりませんでした。")
                return
            history = respondent_history(conn, str(respondent["respondent_id"]))
            payload = [public_respondent_history_payload(item) for item in history]
            last_response_at = payload[-1]["createdAt"] if payload else ""
            self.send_json(
                {
                    "respondent": {
                        "respondentId": str(respondent["respondent_id"] or ""),
                        "respondentName": respondent_name,
                        "responseCount": len(payload),
                        "lastResponseAt": last_response_at,
                    },
                    "history": payload,
                }
            )
        finally:
            conn.close()

    def handle_public_form_submit(self, slug: str) -> None:
        conn = get_connection()
        try:
            form = fetch_form_by_slug(conn, slug, include_inactive=False)
            if not form:
                self.not_found("公開中のフォームが見つかりません。")
                return

            if "multipart/form-data" not in self.headers.get("Content-Type", ""):
                self.bad_request("送信形式が不正です。")
                return

            form_data = self.parse_form_data()
            respondent_name = normalize_respondent_name(str(form_data.getvalue("respondent_name", "")))
            category = str(form_data.getvalue("category", "")).strip() if form["categoryOptions"] else form["title"]
            notes = ""

            if not respondent_name:
                self.bad_request("お名前は必須です。")
                return
            matched_respondent = fetch_respondent_row_by_name(conn, respondent_name)
            respondent_id = (
                str(matched_respondent["respondent_id"])
                if matched_respondent
                else respondent_name_key(respondent_name)
            )
            if form["categoryOptions"] and category not in form["categoryOptions"]:
                self.bad_request("分類の選択が不正です。")
                return

            answer_values: list[dict[str, str]] = []
            if form["slug"] == TICKET_END_FORM_SLUG:
                ticket_sheet_value = normalize_ticket_sheet_value(form_data.getvalue(TICKET_SHEET_FIELD_KEY, ""))
                answer_values.append(
                    {
                        "label": TICKET_SHEET_FIELD_LABEL,
                        "key": TICKET_SHEET_FIELD_KEY,
                        "value": ticket_sheet_value,
                    }
                )
            prepared_files: list[dict[str, Any]] = []
            for field in form["fields"]:
                visible = True
                if field["visibilityFieldKey"] and field["visibilityValues"]:
                    source_value = form_data.getvalue(field["visibilityFieldKey"], "")
                    if isinstance(source_value, list):
                        source_values = {str(item).strip() for item in source_value if str(item).strip()}
                    else:
                        source_values = {str(source_value).strip()} if str(source_value).strip() else set()
                    visible = bool(source_values.intersection(set(field["visibilityValues"])))

                if field["type"] == "file":
                    if not visible:
                        continue
                    files_raw = form_data[field["key"]] if field["key"] in form_data else []
                    if not isinstance(files_raw, list):
                        files_raw = [files_raw]
                    uploaded = validate_uploaded_images(
                        files_raw,
                        label=field["label"],
                        required=field["required"],
                    )
                    for file_info in uploaded:
                        prepared_files.append({**file_info, "field_key": field["key"], "label": field["label"]})
                    continue

                raw_value = form_data.getvalue(field["key"], "")
                if isinstance(raw_value, list):
                    values = [str(item).strip() for item in raw_value if str(item).strip()]
                else:
                    values = [str(raw_value or "").strip()] if str(raw_value or "").strip() else []

                if field["type"] == "checkbox" and field.get("allowOther"):
                    other_selected = "__other__" in values
                    values = [item for item in values if item != "__other__"]
                    other_text = normalize_respondent_name(str(form_data.getvalue(f"{field['key']}__other_text", "")))
                    if other_selected:
                        if not other_text:
                            self.bad_request(f"{field['label']} のその他を選んだ場合は内容を入力してください。")
                            return
                        values.append(f"その他: {other_text}")

                value = ", ".join(values)

                if visible and field["required"] and not value:
                    self.bad_request(f"{field['label']} は必須です。")
                    return
                if value:
                    answer_values.append({"label": field["label"], "key": field["key"], "value": value})

            with conn:
                ensure_respondent_registry(conn, respondent_name, respondent_id=respondent_id)
                cursor = conn.execute(
                    """
                    INSERT INTO responses (
                        form_id, respondent_id, respondent_name, respondent_email,
                        category, notes, created_at, ip_address, user_agent
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        form["id"],
                        respondent_id,
                        respondent_name,
                        "",
                        category,
                        notes,
                        now_iso(),
                        self.client_address[0],
                        self.headers.get("User-Agent", ""),
                    ),
                )
                response_id = cursor.lastrowid

                for answer in answer_values:
                    conn.execute(
                        """
                        INSERT INTO response_answers (response_id, field_key, label, value)
                        VALUES (?, ?, ?, ?)
                        """,
                        (response_id, answer["key"], answer["label"], answer["value"]),
                    )

                for file_info in prepared_files:
                    stored_name = f"{response_id}_{secrets.token_hex(10)}{file_info['extension']}"
                    target = UPLOADS_DIR / stored_name
                    target.write_bytes(file_info["bytes"])
                    conn.execute(
                        """
                        INSERT INTO response_files (
                            response_id, field_key, label, original_name, stored_name, mime_type, size, relative_path
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            response_id,
                            file_info["field_key"],
                            file_info["label"],
                            file_info["original_name"],
                            stored_name,
                            file_info["mime_type"],
                            file_info["size"],
                            local_upload_path(stored_name),
                        ),
                    )

            self.send_json(
                {
                    "ok": True,
                    "message": form["successMessage"],
                    "responseId": response_id,
                }
            )
        except ValueError as exc:
            self.bad_request(str(exc))
        finally:
            conn.close()

    def handle_admin_login(self) -> None:
        body = parse_json_body(self)
        password = str(body.get("password", ""))
        if sha256_hex(password) != CONFIG["admin_password_sha256"]:
            self.send_json({"error": "パスワードが正しくありません。"}, status=401)
            return
        token = sign_session(CONFIG["admin_username"])
        cookie = (
            f"{SESSION_COOKIE}={token}; Path=/; HttpOnly; SameSite=Lax; Max-Age={14 * 24 * 60 * 60}"
        )
        self.send_json({"ok": True, "username": CONFIG["admin_username"]}, headers=[("Set-Cookie", cookie)])

    def handle_admin_logout(self) -> None:
        cookie = f"{SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0"
        self.send_json({"ok": True}, headers=[("Set-Cookie", cookie)])

    def handle_admin_bootstrap(self) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            public_base_url = self.get_preferred_public_base_url()
            forms = fetch_forms(conn, include_inactive=True)
            stats_row = conn.execute(
                """
                SELECT
                    COUNT(DISTINCT forms.id) AS form_count,
                    COUNT(DISTINCT responses.id) AS response_count,
                    (SELECT COUNT(*) FROM respondents) AS respondent_count
                FROM forms
                LEFT JOIN responses ON responses.form_id = forms.id
                """
            ).fetchone()
            recent = list_responses(conn, limit=8)
            self.send_json(
                {
                    "forms": forms,
                    "stats": {
                        "formCount": stats_row["form_count"],
                        "responseCount": stats_row["response_count"],
                        "respondentCount": stats_row["respondent_count"],
                    },
                    "recentResponses": recent,
                    "publicBaseUrl": public_base_url,
                    "defaultPassword": DEFAULT_PASSWORD if CONFIG["admin_password_sha256"] == sha256_hex(DEFAULT_PASSWORD) else None,
                    "settings": {
                        "configuredPublicBaseUrl": str(CONFIG.get("public_base_url") or ""),
                        "publicBaseUrlSource": public_base_url_source(),
                        "publicBaseUrl": public_base_url,
                        "defaultPasswordInUse": CONFIG["admin_password_sha256"] == sha256_hex(DEFAULT_PASSWORD),
                    },
                    "operationsStatus": admin_operations_status(
                        conn,
                        public_base_url=public_base_url,
                        server_port=int(self.server.server_port),
                    ),
                    "backups": list_backup_archives(limit=5),
                }
            )
        finally:
            conn.close()

    def handle_admin_operations_status(self) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            public_base_url = self.get_preferred_public_base_url()
            self.send_json(
                {
                    "status": admin_operations_status(
                        conn,
                        public_base_url=public_base_url,
                        server_port=int(self.server.server_port),
                    ),
                    "backups": list_backup_archives(limit=10),
                }
            )
        finally:
            conn.close()

    def handle_admin_update_public_base_url(self) -> None:
        if not self.require_admin():
            return
        try:
            payload = parse_json_body(self)
            public_base_url = update_public_base_url_config(payload.get("publicBaseUrl", ""))
            self.send_json(
                {
                    "ok": True,
                    "configuredPublicBaseUrl": public_base_url,
                    "publicBaseUrl": self.get_preferred_public_base_url(),
                    "publicBaseUrlSource": public_base_url_source(),
                }
            )
        except ValueError as exc:
            self.bad_request(str(exc))

    def handle_admin_update_password(self) -> None:
        if not self.require_admin():
            return
        try:
            payload = parse_json_body(self)
            update_admin_password(
                str(payload.get("currentPassword", "")),
                str(payload.get("newPassword", "")),
                str(payload.get("confirmPassword", "")),
            )
            self.send_json({"ok": True})
        except ValueError as exc:
            self.bad_request(str(exc))

    def handle_admin_backup_create(self) -> None:
        if not self.require_admin():
            return
        backup = create_backup_archive()
        self.send_json(
            {
                "ok": True,
                "backup": {
                    **backup,
                    "downloadUrl": f"/api/admin/backups/{urllib.parse.quote(backup['name'])}/download",
                },
                "backups": list_backup_archives(limit=10),
            }
        )

    def handle_admin_backup_download(self, backup_name: str) -> None:
        if not self.require_admin():
            return
        backup = find_backup_archive(backup_name)
        if not backup:
            self.not_found("バックアップファイルが見つかりません。")
            return
        self.send_response(200)
        self.send_header("Content-Type", "application/gzip")
        self.send_header("Content-Length", str(backup.stat().st_size))
        self.send_header("Content-Disposition", f'attachment; filename="{backup.name}"')
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        with backup.open("rb") as fh:
            self.write_response_body(fh.read())

    def handle_admin_forms_list(self) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            self.send_json({"forms": fetch_forms(conn, include_inactive=True)})
        finally:
            conn.close()

    def handle_admin_create_respondent(self) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            payload = parse_json_body(self)
            respondent = ensure_respondent_registry(conn, str(payload.get("name", "")))
            overview = respondent_overview(conn, respondent["respondentId"])
            self.send_json({"ok": True, "respondent": overview or respondent})
        except ValueError as exc:
            self.bad_request(str(exc))
        finally:
            conn.close()

    def handle_admin_create_form(self) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            payload = parse_json_body(self)
            form = save_form(conn, payload)
            self.send_json({"ok": True, "form": form})
        except ValueError as exc:
            self.bad_request(str(exc))
        finally:
            conn.close()

    def handle_admin_update_form(self, form_id: int) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            if not fetch_form_by_id(conn, form_id):
                self.not_found("フォームが見つかりません。")
                return
            payload = parse_json_body(self)
            form = save_form(conn, payload, form_id=form_id)
            self.send_json({"ok": True, "form": form})
        except ValueError as exc:
            self.bad_request(str(exc))
        finally:
            conn.close()

    def handle_admin_toggle_form(self, form_id: int) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            form = toggle_form(conn, form_id)
            self.send_json({"ok": True, "form": form})
        except ValueError as exc:
            self.bad_request(str(exc))
        finally:
            conn.close()

    def handle_admin_form_responses(self, form_id: int) -> None:
        if not self.require_admin():
            return
        params = self.get_query_params()
        respondent_query = params.get("respondent", "").strip()
        category = params.get("category", "").strip()
        conn = get_connection()
        try:
            if not fetch_form_by_id(conn, form_id):
                self.not_found("フォームが見つかりません。")
                return
            responses = list_responses(
                conn,
                form_id=form_id,
                respondent_query=respondent_query,
                category=category,
                limit=100,
            )
            summary = category_summary(conn, form_id=form_id, respondent_query=respondent_query)
            self.send_json({"responses": responses, "categorySummary": summary})
        finally:
            conn.close()

    def handle_admin_response_detail(self, response_id: int) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            detail = fetch_response_detail(conn, response_id)
            if not detail:
                self.not_found("回答が見つかりません。")
                return
            self.send_json(detail)
        finally:
            conn.close()

    def handle_admin_respondents(self) -> None:
        if not self.require_admin():
            return
        params = self.get_query_params()
        form_id_raw = params.get("form_id")
        form_id = int(form_id_raw) if form_id_raw and form_id_raw.isdigit() else None
        query = params.get("q", "").strip()
        limit_raw = params.get("limit", "").strip()
        limit = int(limit_raw) if limit_raw.isdigit() else 100
        limit = max(1, min(limit, 1000))
        conn = get_connection()
        try:
            respondents = respondent_summary(conn, form_id=form_id, query=query, limit=limit)
            self.send_json({"respondents": respondents})
        finally:
            conn.close()

    def handle_admin_measurements(self) -> None:
        if not self.require_admin():
            return
        params = self.get_query_params()
        respondent_id = params.get("respondent_id", "").strip() or None
        respondent_name = params.get("respondent_name", "").strip()
        query = params.get("q", "").strip()
        limit_raw = params.get("limit", "").strip()
        limit = int(limit_raw) if limit_raw.isdigit() else 500
        limit = max(1, min(limit, 2000))
        conn = get_connection()
        try:
            records = list_measurement_records(
                conn,
                respondent_id=respondent_id,
                respondent_name=respondent_name,
                query=query,
                limit=limit,
            )
            self.send_json({"records": records})
        finally:
            conn.close()

    def handle_admin_respondent_history(self, respondent_id: str) -> None:
        if not self.require_admin():
            return
        params = self.get_query_params()
        form_id_raw = params.get("form_id")
        form_id = int(form_id_raw) if form_id_raw and form_id_raw.isdigit() else None
        conn = get_connection()
        try:
            respondent = respondent_overview(conn, respondent_id, form_id=form_id)
            if not respondent:
                self.not_found("回答者が見つかりません。")
                return
            history = respondent_history(conn, respondent_id, form_id=form_id)
            image_records = fetch_respondent_profile_records(conn, respondent_id)
            measurement_records = fetch_respondent_measurement_records(conn, respondent_id)
            self.send_json(
                {
                    "respondent": respondent,
                    "history": history,
                    "imageRecords": image_records,
                    "profileRecords": image_records,
                    "measurementRecords": measurement_records,
                }
            )
        finally:
            conn.close()

    def handle_admin_respondent_rename(self, respondent_id: str) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            payload = parse_json_body(self)
            form_id = parse_optional_form_id(payload.get("formId"))
            if form_id is not None and not fetch_form_by_id(conn, form_id):
                self.not_found("フォームが見つかりません。")
                return
            result = rename_respondent(conn, respondent_id, str(payload.get("name", "")), form_id=form_id)
            self.send_json({"ok": True, **result})
        except ValueError as exc:
            self.bad_request(str(exc))
        except LookupError as exc:
            self.not_found(str(exc))
        finally:
            conn.close()

    def handle_admin_respondent_delete(self, respondent_id: str) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            payload = parse_json_body(self)
            form_id = parse_optional_form_id(payload.get("formId"))
            if form_id is not None and not fetch_form_by_id(conn, form_id):
                self.not_found("フォームが見つかりません。")
                return
            result = delete_respondent(conn, respondent_id, form_id=form_id)
            self.send_json({"ok": True, **result})
        except ValueError as exc:
            self.bad_request(str(exc))
        except LookupError as exc:
            self.not_found(str(exc))
        finally:
            conn.close()

    def handle_admin_respondent_profile_update(self, respondent_id: str) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            content_type = self.headers.get("Content-Type", "")
            if "application/json" in content_type:
                payload = parse_json_body(self)
                form_id = parse_optional_form_id(payload.get("formId"))
                respondent_name = normalize_respondent_name(str(payload.get("name", "")))
                ticket_sheet_manual_value = (
                    normalize_optional_ticket_sheet_value(payload.get("ticketSheet"))
                    if "ticketSheet" in payload
                    else None
                )
                ticket_book_type = (
                    normalize_ticket_book_type(payload.get("ticketBookType"))
                    if "ticketBookType" in payload
                    else None
                )
                ticket_stamp_count = (
                    payload.get("ticketStampCount")
                    if "ticketStampCount" in payload
                    else None
                )
                ticket_stamp_manual_enabled = (
                    str(payload.get("ticketStampManualEnabled", "")).strip().lower() in {"1", "true", "yes", "on"}
                    if "ticketStampManualEnabled" in payload
                    else None
                )
            elif "multipart/form-data" in content_type:
                form_data = self.parse_form_data()
                form_id = parse_optional_form_id(form_data.getvalue("form_id", ""))
                respondent_name = normalize_respondent_name(str(form_data.getvalue("respondent_name", "")))
                ticket_sheet_value_raw = form_data.getvalue("ticket_sheet_manual_value")
                ticket_sheet_manual_value = (
                    normalize_optional_ticket_sheet_value(ticket_sheet_value_raw)
                    if ticket_sheet_value_raw is not None
                    else None
                )
                ticket_book_type_raw = form_data.getvalue("current_ticket_book_type")
                ticket_book_type = (
                    normalize_ticket_book_type(ticket_book_type_raw)
                    if ticket_book_type_raw is not None
                    else None
                )
                ticket_stamp_count = form_data.getvalue("current_ticket_stamp_count")
                ticket_stamp_manual_enabled_raw = form_data.getvalue("current_ticket_stamp_manual_enabled")
                ticket_stamp_manual_enabled = (
                    str(ticket_stamp_manual_enabled_raw or "").strip().lower() in {"1", "true", "yes", "on"}
                    if ticket_stamp_manual_enabled_raw is not None
                    else None
                )
            else:
                self.bad_request("送信形式が不正です。")
                return

            if form_id is not None and not fetch_form_by_id(conn, form_id):
                self.not_found("フォームが見つかりません。")
                return

            respondent = respondent_overview(conn, respondent_id)
            if not respondent:
                self.not_found("回答者が見つかりません。")
                return

            if not respondent_name:
                respondent_name = respondent["respondentName"]
            if not respondent_name:
                self.bad_request("お名前は必須です。")
                return

            existing_ticket_book_type = str(respondent.get("currentTicketBookType") or "")
            resolved_ticket_book_type = existing_ticket_book_type if ticket_book_type is None else ticket_book_type
            resolved_ticket_stamp_manual_enabled = (
                bool(respondent.get("currentTicketStampManualEnabled"))
                if ticket_stamp_manual_enabled is None
                else ticket_stamp_manual_enabled
            )
            resolved_ticket_stamp_count = (
                int(respondent.get("currentTicketStampManualValue") or 0)
                if ticket_stamp_count is None
                else normalize_ticket_stamp_count(ticket_stamp_count, resolved_ticket_book_type)
            )

            target_respondent_id = respondent_id
            if respondent_name != normalize_respondent_name(respondent["respondentName"]):
                renamed = rename_respondent(conn, respondent_id, respondent_name, form_id=form_id)
                target_respondent_id = renamed["respondentId"]
            else:
                move_respondent_profile(conn, target_respondent_id, target_respondent_id, respondent_name)
                move_respondent_measurements(conn, target_respondent_id, target_respondent_id, respondent_name)
                ensure_respondent_registry(conn, respondent_name, respondent_id=target_respondent_id)

            if ticket_sheet_manual_value is not None or ticket_book_type is not None or ticket_stamp_count is not None:
                update_respondent_ticket_status(
                    conn,
                    target_respondent_id,
                    ticket_sheet_manual_value=(
                        ticket_sheet_manual_value
                        if ticket_sheet_manual_value is not None
                        else str(respondent.get("latestTicketSheetManualValue") or "")
                    ),
                    current_ticket_book_type=resolved_ticket_book_type,
                    current_ticket_stamp_count=resolved_ticket_stamp_count if resolved_ticket_stamp_manual_enabled else 0,
                    current_ticket_stamp_manual_enabled=resolved_ticket_stamp_manual_enabled,
                )

            self.send_json(
                {
                    "ok": True,
                    "respondentId": target_respondent_id,
                    "respondentName": respondent_name,
                    "ticketSheet": ticket_sheet_manual_value if ticket_sheet_manual_value is not None else "",
                    "ticketBookType": resolved_ticket_book_type,
                    "ticketStampCount": resolved_ticket_stamp_count,
                    "ticketStampManualEnabled": resolved_ticket_stamp_manual_enabled,
                }
            )
        except ValueError as exc:
            self.bad_request(str(exc))
        except LookupError as exc:
            self.not_found(str(exc))
        finally:
            conn.close()

    def handle_admin_respondent_profile_record_create(self, respondent_id: str) -> None:
        if not self.require_admin():
            return
        if "multipart/form-data" not in self.headers.get("Content-Type", ""):
            self.bad_request("送信形式が不正です。")
            return
        conn = get_connection()
        try:
            form_data = self.parse_form_data()
            form_id = parse_optional_form_id(form_data.getvalue("form_id", ""))
            if form_id is not None and not fetch_form_by_id(conn, form_id):
                self.not_found("フォームが見つかりません。")
                return

            respondent = respondent_overview(conn, respondent_id)
            if not respondent:
                self.not_found("回答者が見つかりません。")
                return

            title = str(form_data.getvalue("title", "")).strip()
            entry_date = str(form_data.getvalue("entry_date", "")).strip()
            memo = str(form_data.getvalue("memo", "")).strip()

            files_raw = form_data["profile_image"] if "profile_image" in form_data else []
            if not isinstance(files_raw, list):
                files_raw = [files_raw]
            uploaded_images = validate_uploaded_images(files_raw, label="回答者画像", required=True)
            if len(uploaded_images) != 1:
                self.bad_request("画像は1枚のみアップロードしてください。")
                return

            record = create_respondent_profile_record(
                conn,
                respondent_id,
                respondent["respondentName"],
                title=title,
                entry_date=entry_date,
                memo=memo,
                uploaded_image=uploaded_images[0],
            )
            self.send_json({"ok": True, "record": record})
        except ValueError as exc:
            self.bad_request(str(exc))
        except LookupError as exc:
            self.not_found(str(exc))
        finally:
            conn.close()

    def handle_admin_respondent_profile_record_update(self, respondent_id: str, source_type: str, record_id: int) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            payload = parse_json_body(self)
            record = update_respondent_profile_record(
                conn,
                respondent_id,
                source_type,
                record_id,
                title=str(payload.get("title", "")),
                entry_date=str(payload.get("entryDate", "")),
                memo=str(payload.get("memo", "")),
            )
            self.send_json({"ok": True, "record": record})
        except ValueError as exc:
            self.bad_request(str(exc))
        except LookupError as exc:
            self.not_found(str(exc))
        finally:
            conn.close()

    def handle_admin_respondent_profile_record_delete(
        self,
        respondent_id: str,
        source_type: str,
        record_id: int,
    ) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            if source_type != "manual":
                self.bad_request("アンケート画像は削除できません。")
                return
            result = delete_respondent_profile_record(conn, respondent_id, record_id)
            self.send_json({"ok": True, **result})
        except LookupError as exc:
            self.not_found(str(exc))
        finally:
            conn.close()

    def handle_admin_respondent_measurement_create(self, respondent_id: str) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            payload = parse_json_body(self)
            form_id = parse_optional_form_id(payload.get("formId"))
            if form_id is not None and not fetch_form_by_id(conn, form_id):
                self.not_found("フォームが見つかりません。")
                return
            respondent = respondent_overview(conn, respondent_id)
            if not respondent:
                self.not_found("回答者が見つかりません。")
                return
            record = create_respondent_measurement_record(
                conn,
                respondent_id,
                respondent["respondentName"],
                entry_date=payload.get("entryDate", ""),
                category=payload.get("category", ""),
                waist=payload.get("waist", ""),
                hip=payload.get("hip", ""),
                thigh=payload.get("thigh", ""),
            )
            self.send_json({"ok": True, "record": record})
        except ValueError as exc:
            self.bad_request(str(exc))
        except LookupError as exc:
            self.not_found(str(exc))
        finally:
            conn.close()

    def handle_admin_respondent_measurement_update(self, respondent_id: str, record_id: int) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            payload = parse_json_body(self)
            record = update_respondent_measurement_record(
                conn,
                respondent_id,
                record_id,
                entry_date=payload.get("entryDate", ""),
                category=payload.get("category", ""),
                waist=payload.get("waist", ""),
                hip=payload.get("hip", ""),
                thigh=payload.get("thigh", ""),
            )
            self.send_json({"ok": True, "record": record})
        except ValueError as exc:
            self.bad_request(str(exc))
        except LookupError as exc:
            self.not_found(str(exc))
        finally:
            conn.close()

    def handle_admin_respondent_measurement_import(self, respondent_id: str) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            payload = parse_json_body(self)
            respondent = respondent_overview(conn, respondent_id)
            if not respondent:
                self.not_found("回答者が見つかりません。")
                return
            result = import_respondent_measurements_from_sheet(
                conn,
                respondent_id,
                respondent["respondentName"],
                sheet_url=str(payload.get("sheetUrl", "")),
            )
            self.send_json({"ok": True, **result})
        except ValueError as exc:
            self.bad_request(str(exc))
        finally:
            conn.close()

    def handle_admin_respondent_measurement_delete(self, respondent_id: str, record_id: int) -> None:
        if not self.require_admin():
            return
        conn = get_connection()
        try:
            result = delete_respondent_measurement_record(conn, respondent_id, record_id)
            self.send_json({"ok": True, **result})
        except LookupError as exc:
            self.not_found(str(exc))
        finally:
            conn.close()


def run_server(host: str, port: int) -> None:
    init_db()
    server = BijirisServer((host, port), BijirisHandler)
    local_url = f"http://127.0.0.1:{port}"
    public_url = preferred_public_base_url(host, fallback_port=port)
    print(f"Bijiris local: {local_url}")
    print(f"Bijiris public: {public_url}")
    if CONFIG["admin_password_sha256"] == sha256_hex(DEFAULT_PASSWORD):
        print(f"Admin password: {DEFAULT_PASSWORD}")
    else:
        print("Admin password: custom password configured")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down.")
    finally:
        server.server_close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Bijiris survey and image management app")
    default_host = os.environ.get("BIJIRIS_HOST", DEFAULT_SERVER_HOST).strip() or DEFAULT_SERVER_HOST
    default_port_text = os.environ.get("PORT") or os.environ.get("BIJIRIS_PORT") or str(DEFAULT_SERVER_PORT)
    try:
        default_port = int(str(default_port_text).strip())
    except ValueError:
        default_port = DEFAULT_SERVER_PORT
    parser.add_argument("--host", default=default_host)
    parser.add_argument("--port", type=int, default=default_port)
    args = parser.parse_args()
    run_server(args.host, args.port)


if __name__ == "__main__":
    main()
