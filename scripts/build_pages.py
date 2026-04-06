#!/usr/bin/env python3
from __future__ import annotations

import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PUBLIC = ROOT / "public"
DOCS = ROOT / "docs"
ASSETS = DOCS / "assets"
ADMIN_DIR = DOCS / "admin"
ICONS_SRC = PUBLIC / "icons"
ICONS_DST = ASSETS / "icons"


RESPONDENT_HTML = """<!DOCTYPE html>
<html lang="ja">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <meta name="format-detection" content="telephone=no" />
    <meta name="theme-color" content="#21374c" />
    <meta name="mobile-web-app-capable" content="yes" />
    <meta name="apple-mobile-web-app-capable" content="yes" />
    <meta name="apple-mobile-web-app-status-bar-style" content="default" />
    <meta name="apple-mobile-web-app-title" content="ビジリス" />
    <title>ビジリス 回答フォーム</title>
    <link rel="manifest" href="./manifest.webmanifest" />
    <link rel="icon" type="image/png" sizes="32x32" href="./assets/icons/favicon-32.png" />
    <link rel="apple-touch-icon" sizes="180x180" href="./assets/icons/apple-touch-icon.png" />
    <link rel="stylesheet" href="./assets/style.css" />
  </head>
  <body class="respondent-body">
    <main class="respondent-shell">
      <section class="respondent-card">
        <p class="eyebrow">BIJIRIS FORM</p>
        <div id="selectorPanel" class="stack-gap">
          <div class="stack-gap compact-gap respondent-intro">
            <h1>アンケートを選択してください</h1>
            <div class="button-row respondent-intro-actions">
              <button id="openHistoryFromSelector" class="secondary-button compact-button" type="button">回答履歴を見る</button>
              <button id="reloadAppFromSelector" class="secondary-button compact-button" type="button">最新版に更新</button>
            </div>
          </div>
          <section id="installCard" class="respondent-install-card stack-gap compact-gap hidden">
            <div class="spread respondent-install-head">
              <div class="stack-gap compact-gap">
                <p class="respondent-kicker">APP INSTALL</p>
                <h2>スマホに追加してアプリのように使えます</h2>
              </div>
              <span id="installStatusBadge" class="pill hidden">インストール済み</span>
            </div>
            <p id="installDescription" class="muted">ホーム画面に追加すると、次回からアプリのようにすぐ開けます。</p>
            <div class="button-row respondent-intro-actions">
              <button id="installAppButton" class="primary-button compact-button" type="button">ホーム画面に追加</button>
              <button id="installGuideToggle" class="ghost-button compact-button hidden" type="button">追加方法を見る</button>
            </div>
            <div id="installGuide" class="note-box hidden"></div>
          </section>
          <div id="selectorList" class="card-list"></div>
        </div>

        <div id="respondentHeader" class="stack-gap compact-gap">
          <div class="button-row respondent-header-actions">
            <button id="backToSelector" class="ghost-button hidden" type="button">アンケート一覧に戻る</button>
            <button id="openHistoryFromHeader" class="ghost-button compact-button hidden" type="button">回答履歴を見る</button>
            <button id="reloadAppButton" class="secondary-button compact-button" type="button">最新版に更新</button>
          </div>
          <div class="respondent-heading-shell">
            <p class="respondent-kicker">SELECTED SURVEY</p>
            <div class="respondent-heading-mark"></div>
          </div>
          <h1 id="formTitle">フォームを読み込み中です</h1>
          <p id="formDescription" class="muted">少しお待ちください。</p>
        </div>

        <div id="formErrorBanner" class="error-banner hidden"></div>

        <section id="ticketStepPanel" class="respondent-step-panel stack-gap hidden">
          <div class="respondent-step-shell stack-gap compact-gap">
            <p class="respondent-kicker">TICKET STEP</p>
            <h2>今回終了した回数券を教えてください</h2>
            <p class="muted">最初に、何枚目の回数券が終了したかを入力してください。その後にアンケートへ進みます。</p>
            <label class="field">
              <span>何枚目の回数券が終了しましたか？</span>
              <input id="ticketSheetInput" type="number" inputmode="numeric" min="1" step="1" placeholder="例: 2" />
            </label>
            <div class="button-row respondent-step-actions">
              <button id="ticketStepContinue" class="primary-button" type="button">質問へ進む</button>
            </div>
            <p id="ticketStepError" class="error-text"></p>
          </div>
        </section>

        <section id="historyPanel" class="respondent-step-panel stack-gap hidden">
          <div class="respondent-step-shell respondent-history-shell stack-gap compact-gap">
            <div class="button-row respondent-step-actions">
              <button id="historyBackButton" class="ghost-button compact-button" type="button">戻る</button>
            </div>
            <p class="respondent-kicker">HISTORY</p>
            <h2>アンケート回答履歴</h2>
            <p class="muted">お名前を入力すると、これまでの回答履歴を確認できます。</p>
            <label class="field">
              <span>お名前</span>
              <input id="historyRespondentName" type="text" autocomplete="name" placeholder="例: 鈴木太郎" />
            </label>
            <div class="button-row respondent-step-actions">
              <button id="historySearchButton" class="primary-button" type="button">履歴を表示</button>
            </div>
            <p id="historySearchError" class="error-text"></p>
            <div id="historySummary" class="note-box hidden"></div>
            <div id="historyResults" class="stack-gap"></div>
          </div>
        </section>

        <form id="respondentForm" class="stack-gap hidden" enctype="multipart/form-data">
          <label class="field">
            <span>お名前</span>
            <input type="text" name="respondent_name" required autocomplete="name" />
          </label>

          <div id="ticketSummary" class="note-box respondent-ticket-summary hidden">
            <div>
              <strong>今回終了した回数券</strong>
              <div id="ticketSummaryValue" class="respondent-ticket-summary-value"></div>
            </div>
            <button id="ticketStepEdit" class="ghost-button compact-button" type="button">変更</button>
          </div>

          <label class="field">
            <span id="categoryLabel">分類</span>
            <select id="categorySelect" name="category" required></select>
          </label>

          <div id="customFields" class="stack-gap"></div>
          <button id="submitButton" class="primary-button" type="submit">送信する</button>
          <p id="submitError" class="error-text"></p>
        </form>

        <div id="successPanel" class="success-panel hidden">
          <h2>送信完了</h2>
          <p id="successMessage"></p>
          <div class="button-row">
            <button id="resetButton" class="secondary-button" type="button">続けて送信する</button>
            <button id="openHistoryFromSuccess" class="ghost-button" type="button">回答履歴を見る</button>
          </div>
        </div>
      </section>
    </main>

    <script src="./assets/config.js"></script>
    <script src="./assets/runtime.js" defer></script>
    <script src="./assets/respondent.js" defer></script>
  </body>
</html>
"""


ADMIN_HTML = """<!DOCTYPE html>
<html lang="ja">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Bijiris 管理画面</title>
    <link rel="stylesheet" href="../assets/style.css" />
  </head>
  <body class="admin-body">
    <div id="loginPanel" class="login-shell">
      <div class="login-card">
        <p class="eyebrow">BIJIRIS</p>
        <h1>回答管理アプリ</h1>
        <p class="muted">管理者ログイン後に、フォーム作成、QRコード共有、回答履歴の確認を行えます。</p>
        <form id="loginForm" class="stack-gap">
          <label class="field">
            <span>管理者パスワード</span>
            <input id="loginPassword" type="password" name="password" autocomplete="current-password" required />
          </label>
          <button class="primary-button" type="submit">ログイン</button>
          <p id="loginHelp" class="inline-help"></p>
          <p id="loginError" class="error-text"></p>
        </form>
      </div>
    </div>

    <!-- Management markup stays identical to the local admin page -->
    <div id="adminApp" class="admin-shell hidden"></div>

    <script src="../assets/config.js"></script>
    <script src="../assets/runtime.js" defer></script>
    <script src="../assets/admin.js" defer></script>
  </body>
</html>
"""


CONFIG_JS = """window.BIJIRIS_CONFIG = Object.assign({}, window.BIJIRIS_CONFIG || {}, {
  apiMode: "gas",
  gasUrl: "",
  siteRootUrl: "",
  respondentUrl: "",
  adminUrl: ""
});
"""


def copy_file(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)


def build_manifest() -> str:
    return """{
  "name": "ビジリス",
  "short_name": "ビジリス",
  "start_url": ".",
  "scope": ".",
  "display": "standalone",
  "background_color": "#f7f5ef",
  "theme_color": "#21374c",
  "icons": [
    {
      "src": "./assets/icons/icon-192.png",
      "sizes": "192x192",
      "type": "image/png"
    },
    {
      "src": "./assets/icons/icon-512.png",
      "sizes": "512x512",
      "type": "image/png"
    },
    {
      "src": "./assets/icons/apple-touch-icon.png",
      "sizes": "180x180",
      "type": "image/png"
    }
  ]
}
"""


def extract_admin_app_markup() -> str:
    admin_html = (PUBLIC / "admin.html").read_text(encoding="utf-8")
    start = admin_html.index('<div id="adminApp"')
    end = admin_html.index("    <script", start)
    return admin_html[start:end].rstrip()


def main() -> None:
    if DOCS.exists():
      shutil.rmtree(DOCS)
    ASSETS.mkdir(parents=True, exist_ok=True)
    ADMIN_DIR.mkdir(parents=True, exist_ok=True)
    ICONS_DST.mkdir(parents=True, exist_ok=True)

    copy_file(PUBLIC / "style.css", ASSETS / "style.css")
    copy_file(PUBLIC / "runtime.js", ASSETS / "runtime.js")
    copy_file(PUBLIC / "respondent.js", ASSETS / "respondent.js")
    copy_file(PUBLIC / "admin.js", ASSETS / "admin.js")
    copy_file(PUBLIC / "sw.js", DOCS / "sw.js")
    for icon in ICONS_SRC.glob("*"):
        if icon.is_file():
            copy_file(icon, ICONS_DST / icon.name)

    (ASSETS / "config.js").write_text(CONFIG_JS, encoding="utf-8")
    (DOCS / "manifest.webmanifest").write_text(build_manifest(), encoding="utf-8")
    (DOCS / "index.html").write_text(RESPONDENT_HTML, encoding="utf-8")

    admin_html = ADMIN_HTML.replace('<div id="adminApp" class="admin-shell hidden"></div>', extract_admin_app_markup())
    (ADMIN_DIR / "index.html").write_text(admin_html, encoding="utf-8")


if __name__ == "__main__":
    main()
