import os
import re
import json
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QTextEdit,
    QPushButton,
    QFileDialog,
    QLabel,
    QListWidget,
    QComboBox,
    QLineEdit,
    QMessageBox,
)

import language_tool_python
import feedparser
import requests
from docx import Document
from pypdf import PdfReader
from whoosh.fields import Schema, TEXT, ID
from whoosh.index import create_in, open_dir

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

APP_DIR = Path.home() / ".velin_desk"
APP_DIR.mkdir(exist_ok=True)
NOTES_DIR = APP_DIR / "notes"
NOTES_DIR.mkdir(exist_ok=True)
RSS_DIR = APP_DIR / "rss_index"
RSS_DIR.mkdir(exist_ok=True)
READ_LATER_FILE = APP_DIR / "read_later.json"


class VelinDesk(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VELIN Desk — Personal Writing Atelier")
        self.resize(1200, 800)

        self.current_text = ""
        self.tool = language_tool_python.LanguageToolPublicAPI("en-US")
        self.cursor_glow = False

        self.central = QWidget()
        self.setCentralWidget(self.central)
        root = QHBoxLayout(self.central)

        left = QVBoxLayout()
        center = QVBoxLayout()
        right = QVBoxLayout()

        root.addLayout(left, 1)
        root.addLayout(center, 3)
        root.addLayout(right, 1)

        self.lang_combo = QComboBox()
        self.lang_combo.addItems(["tr-TR", "en-US", "fr", "it"])
        self.lang_combo.currentTextChanged.connect(self.change_language)

        self.editor = QTextEdit()
        self.editor.setPlaceholderText("Yazmaya başla…")
        self.editor.textChanged.connect(self.typewriter_effect)

        self.spell_btn = QPushButton("Yazım Denetimi")
        self.spell_btn.clicked.connect(self.check_spelling)

        self.open_btn = QPushButton("Dosya Aç (md/pdf/docx)")
        self.open_btn.clicked.connect(self.open_file)

        self.ai_btn = QPushButton("AI ile Geliştir")
        self.ai_btn.clicked.connect(self.ai_improve)

        self.result = QTextEdit()
        self.result.setReadOnly(True)

        center.addWidget(self.lang_combo)
        center.addWidget(self.editor)
        center.addWidget(self.spell_btn)
        center.addWidget(self.open_btn)
        center.addWidget(self.ai_btn)
        center.addWidget(QLabel("Çıktı / Öneriler"))
        center.addWidget(self.result)

        left.addWidget(QLabel("📒 Zettelkasten Notlar"))
        self.note_list = QListWidget()
        self.refresh_notes()
        self.new_note_btn = QPushButton("Yeni Zettel")
        self.new_note_btn.clicked.connect(self.new_zettel)
        left.addWidget(self.note_list)
        left.addWidget(self.new_note_btn)

        right.addWidget(QLabel("📰 RSS + Read Later"))
        self.rss_input = QLineEdit()
        self.rss_input.setPlaceholderText("RSS URL")
        self.rss_btn = QPushButton("RSS Çek + İndeksle")
        self.rss_btn.clicked.connect(self.fetch_rss)
        self.read_later = QListWidget()
        self.load_read_later()
        self.cloud_btn = QPushButton("Cloud Sync (WebDAV)")
        self.cloud_btn.clicked.connect(self.sync_to_webdav)

        right.addWidget(self.rss_input)
        right.addWidget(self.rss_btn)
        right.addWidget(self.read_later)
        right.addWidget(self.cloud_btn)

        self.theme_timer = QTimer()
        self.theme_timer.timeout.connect(self.toggle_cursor_glow)
        self.theme_timer.start(500)

        self.apply_theme()

    def apply_theme(self):
        self.setStyleSheet(
            """
            QMainWindow { background-color: #ece3d3; }
            QTextEdit {
                background: #f6f1e5;
                border: 1px solid #bda98a;
                color: #2d2418;
                font-family: 'Georgia';
                font-size: 15px;
                selection-background-color: #cdbb9a;
            }
            QPushButton {
                background: #d8c6a8;
                border: 1px solid #927a58;
                padding: 6px;
            }
            QLabel { color: #4f3f2b; font-weight: bold; }
            QListWidget, QLineEdit {
                background: #f8f3e9;
                border: 1px solid #bfa98d;
            }
            """
        )

    def change_language(self, code: str):
        self.tool = language_tool_python.LanguageToolPublicAPI(code)

    def typewriter_effect(self):
        txt = self.editor.toPlainText()
        if len(txt) > len(self.current_text):
            self.statusBar().showMessage("Original Crown Mill hissi: harfler işleniyor…", 900)
        self.current_text = txt

    def toggle_cursor_glow(self):
        self.cursor_glow = not self.cursor_glow
        color = "#8b6f47" if self.cursor_glow else "#2d2418"
        self.editor.setStyleSheet(f"QTextEdit{{caret-color:{color};}}")

    def check_spelling(self):
        text = self.editor.toPlainText().strip()
        if not text:
            return
        matches = self.tool.check(text)
        if not matches:
            self.result.setPlainText("Yazım denetimi temiz ✅")
            return
        lines = []
        for m in matches[:120]:
            lines.append(f"- {m.message} | Öneri: {', '.join(m.replacements[:3])}")
        self.result.setPlainText("\n".join(lines))

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Aç", "", "Docs (*.md *.pdf *.docx)")
        if not path:
            return
        ext = Path(path).suffix.lower()
        content = ""
        if ext == ".md":
            content = Path(path).read_text(encoding="utf-8", errors="ignore")
        elif ext == ".pdf":
            reader = PdfReader(path)
            content = "\n".join(page.extract_text() or "" for page in reader.pages)
        elif ext == ".docx":
            doc = Document(path)
            content = "\n".join(p.text for p in doc.paragraphs)
        self.editor.setPlainText(content)

    def new_zettel(self):
        z_id = datetime.utcnow().strftime("%Y%m%d%H%M%S")
        text = self.editor.toPlainText()
        links = re.findall(r"\[\[(.*?)\]\]", text)
        data = {"id": z_id, "text": text, "links": links, "created": datetime.utcnow().isoformat()}
        (NOTES_DIR / f"{z_id}.json").write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        self.refresh_notes()

    def refresh_notes(self):
        self.note_list.clear()
        for f in sorted(NOTES_DIR.glob("*.json")):
            self.note_list.addItem(f.stem)

    def fetch_rss(self):
        url = self.rss_input.text().strip()
        if not url:
            return
        feed = feedparser.parse(url)
        ix = self.ensure_index()
        writer = ix.writer()
        for e in feed.entries[:50]:
            title = getattr(e, "title", "")
            link = getattr(e, "link", "")
            summary = getattr(e, "summary", "")
            writer.update_document(doc_id=link or title, title=title, body=summary)
            self.read_later.addItem(f"{title} | {link}")
        writer.commit()
        self.save_read_later()

    def ensure_index(self):
        schema = Schema(doc_id=ID(stored=True, unique=True), title=TEXT(stored=True), body=TEXT(stored=True))
        if not any(RSS_DIR.iterdir()):
            return create_in(str(RSS_DIR), schema)
        return open_dir(str(RSS_DIR))

    def load_read_later(self):
        if READ_LATER_FILE.exists():
            items = json.loads(READ_LATER_FILE.read_text(encoding="utf-8"))
            for i in items:
                self.read_later.addItem(i)

    def save_read_later(self):
        items = [self.read_later.item(i).text() for i in range(self.read_later.count())]
        READ_LATER_FILE.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")

    def sync_to_webdav(self):
        url = os.getenv("WEBDAV_URL")
        user = os.getenv("WEBDAV_USER")
        pwd = os.getenv("WEBDAV_PASS")
        if not all([url, user, pwd]):
            QMessageBox.warning(self, "Eksik", "WEBDAV_URL / USER / PASS env değişkenleri gerekli.")
            return
        for nf in NOTES_DIR.glob("*.json"):
            with open(nf, "rb") as fh:
                requests.put(f"{url.rstrip('/')}/{nf.name}", data=fh, auth=(user, pwd), timeout=30)
        QMessageBox.information(self, "Tamam", "Cloud sync tamamlandı.")

    def ai_improve(self):
        key = os.getenv("OPENAI_API_KEY")
        if not key or OpenAI is None:
            self.result.setPlainText("OPENAI_API_KEY veya OpenAI SDK bulunamadı.")
            return
        client = OpenAI(api_key=key)
        text = self.editor.toPlainText().strip()
        if not text:
            return
        rsp = client.responses.create(
            model="gpt-5-mini",
            input=f"Bu metni daha akıcı hale getir, stil önerileri sun:\n\n{text}",
        )
        self.result.setPlainText(rsp.output_text)


if __name__ == "__main__":
    app = QApplication([])
    w = VelinDesk()
    w.show()
    app.exec()
