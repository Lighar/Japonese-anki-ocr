"""
Japanese Dictionary Screen Capture Tool (PyQt6 + Wayland)

Left-click and drag to select a region of the screen.
On release, OCR extracts Japanese text and queries OpenRouter LLM
for dictionary-style word information displayed in a popup with Anki card saving.
"""

import sys
import os
import json
import re
import threading

import PIL.ImageGrab
import requests

from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout,
    QTextEdit, QScrollArea, QFrame, QPushButton,
)
from PyQt6.QtGui import QPainter, QColor, QPen, QPixmap, QFont, QCursor
from PyQt6.QtCore import Qt, QRect, QPoint, pyqtSignal, QObject

from dotenv import load_dotenv

load_dotenv()

# ── Configuration ──────────────────────────────────────────────────────────────
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "")
OPENROUTER_MODEL = "google/gemini-2.0-flash-001"
ANKI_CONNECT_URL = "http://localhost:8765"
ANKI_DECK_NAME = "Japanese"
ANKI_MODEL_NAME = None  # Auto-detected at runtime
SAVED_WORDS_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "saved_words.json"
)

# ── Lazy-loaded OCR model ─────────────────────────────────────────────────────
_ocr_model = None


def get_ocr():
    """Lazy-load manga-ocr (downloads ~400MB model on first run)."""
    global _ocr_model
    if _ocr_model is None:
        print("Loading manga-ocr model (first run downloads ~400MB)...")
        from manga_ocr import MangaOcr
        _ocr_model = MangaOcr()
        print("manga-ocr model loaded.")
    return _ocr_model


# ── OpenRouter LLM API ────────────────────────────────────────────────────────

def query_openrouter(text: str) -> dict:
    """Send extracted Japanese text to OpenRouter for analysis + vocab cards."""
    if not OPENROUTER_API_KEY:
        return {
            "analysis": "⚠️ OPENROUTER_API_KEY not set.\nAdd it to .env file.",
            "vocab": [],
        }

    prompt = f"""You are a Japanese language teacher and dictionary assistant.
Given the following Japanese text extracted from a screen capture, provide:

1. A FULL SENTENCE ANALYSIS — explain the entire sentence/phrase, not just one word.
   Break down every component: particles, verb conjugations, grammatical structures, etc.

2. A list of VOCABULARY CARDS for Anki — pick the most interesting/useful words from the text.

Text: 「{text}」

Your response MUST follow this exact format. First, provide the analysis in markdown.
Then, after the analysis, include a JSON block with vocabulary cards.

## Full Sentence Analysis

**Full text**: [the full text in Japanese]
**Reading**: [full reading in hiragana]
**Translation**: [English translation]

### Breakdown
[Break down each word/particle in the sentence. For each element explain:]
- what it is (particle, noun, verb form, copula, etc.)
- its meaning
- any grammar notes

### Context & Register
[Explain the overall register, politeness level, and when you'd see this kind of sentence]

---

VOCABULARY_JSON_START
[
  {{
    "word": "[kanji form]",
    "reading": "[hiragana reading]",
    "katakana": "[katakana reading]",
    "meaning": "[English meaning]",
    "pos": "[part of speech]",
    "frequency": "[one of: very common, common, uncommon, rare]",
    "frequency_rank": [estimated rank number in manga/anime frequency lists, e.g. 150],
    "example": "[short example sentence in Japanese]",
    "example_en": "[English translation of example]"
  }}
]
VOCABULARY_JSON_END

Include 2-5 vocabulary items. Prioritize words that are:
- Useful for manga/anime reading
- Not ultra-basic (skip は, が, です unless in a special usage)
- Interesting for a learner

The frequency_rank should be your best estimate of how often this word appears in manga/anime
(1 = most common, higher = less common). Be reasonably accurate."""

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": OPENROUTER_MODEL,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3,
    }

    try:
        resp = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers=headers, json=payload, timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        content = data["choices"][0]["message"]["content"]
        analysis, vocab = _parse_llm_response(content)
        return {"analysis": analysis, "vocab": vocab}
    except requests.exceptions.Timeout:
        return {"analysis": "⚠️ Request timed out. Try again.", "vocab": []}
    except requests.exceptions.RequestException as e:
        return {"analysis": f"⚠️ API error: {e}", "vocab": []}
    except (KeyError, IndexError):
        return {"analysis": "⚠️ Unexpected API response format.", "vocab": []}


def _parse_llm_response(content: str) -> tuple:
    """Split the LLM response into analysis text and vocabulary JSON."""
    vocab = []

    json_match = re.search(
        r'VOCABULARY_JSON_START\s*\n?(.*?)\n?\s*VOCABULARY_JSON_END',
        content, re.DOTALL,
    )
    if json_match:
        json_str = json_match.group(1).strip()
        analysis = content[:json_match.start()].strip()
        analysis = re.sub(r'\n-{3,}\s*$', '', analysis).strip()
        # Strip optional ```json ... ``` fences the LLM sometimes adds
        json_str = re.sub(r'^```(?:json)?\s*', '', json_str)
        json_str = re.sub(r'\s*```$', '', json_str).strip()
        try:
            vocab = json.loads(json_str)
        except json.JSONDecodeError:
            try:
                cleaned = re.sub(r',\s*([}\]])', r'\1', json_str)
                vocab = json.loads(cleaned)
            except json.JSONDecodeError:
                print(f"Failed to parse vocab JSON: {json_str[:200]}")
    else:
        analysis = content
        json_array_match = re.search(
            r'(\[\s*\{.*?\}\s*\])\s*$', content, re.DOTALL
        )
        if json_array_match:
            try:
                vocab = json.loads(json_array_match.group(1))
                analysis = content[:json_array_match.start()].strip()
            except json.JSONDecodeError:
                pass

    return analysis, vocab


# ── AnkiConnect Integration ───────────────────────────────────────────────────

def anki_connect_request(action, **params):
    """Send a request to AnkiConnect."""
    payload = {"action": action, "version": 6, "params": params}
    try:
        resp = requests.post(ANKI_CONNECT_URL, json=payload, timeout=5)
        resp.raise_for_status()
        result = resp.json()
        if result.get("error"):
            return False, result["error"]
        return True, result.get("result")
    except requests.exceptions.ConnectionError:
        return False, "AnkiConnect not running. Open Anki with AnkiConnect plugin."
    except Exception as e:
        return False, str(e)


def _get_anki_model_name():
    """Auto-detect a usable Anki model with Front/Back fields, or create one."""
    global ANKI_MODEL_NAME
    if ANKI_MODEL_NAME is not None:
        return ANKI_MODEL_NAME

    success, models = anki_connect_request("modelNames")
    if not success or not models:
        ANKI_MODEL_NAME = "Basic"
        return ANKI_MODEL_NAME

    for model in models:
        ok, fields = anki_connect_request("modelFieldNames", modelName=model)
        if ok and fields and "Front" in fields and "Back" in fields:
            ANKI_MODEL_NAME = model
            print(f"Using Anki model: {model}")
            return ANKI_MODEL_NAME

    custom_name = "JapCapture"
    success, _ = anki_connect_request(
        "createModel",
        modelName=custom_name,
        inOrderFields=["Front", "Back"],
        cardTemplates=[{
            "Name": "Card 1",
            "Front": "{{Front}}",
            "Back": "{{FrontSide}}<hr id=answer>{{Back}}",
        }],
    )
    if success:
        ANKI_MODEL_NAME = custom_name
        print(f"Created Anki model: {custom_name}")
    else:
        ANKI_MODEL_NAME = models[0] if models else "Basic"
        print(f"Falling back to Anki model: {ANKI_MODEL_NAME}")

    return ANKI_MODEL_NAME


def add_to_anki(word, reading, meaning, example=""):
    """Add a card to Anki via AnkiConnect."""
    front = f"{word}\n{reading}"
    back = meaning
    if example:
        back += f"\n\n{example}"

    anki_connect_request("createDeck", deck=ANKI_DECK_NAME)
    model_name = _get_anki_model_name()

    return anki_connect_request(
        "addNote",
        note={
            "deckName": ANKI_DECK_NAME,
            "modelName": model_name,
            "fields": {"Front": front, "Back": back},
            "options": {"allowDuplicate": False},
            "tags": ["jap-capture"],
        },
    )


# ── Local Word Storage ────────────────────────────────────────────────────────

def load_saved_words():
    if os.path.exists(SAVED_WORDS_FILE):
        try:
            with open(SAVED_WORDS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return []
    return []


def save_word_locally(word_data):
    words = load_saved_words()
    for existing in words:
        if (existing.get("word") == word_data.get("word")
                and existing.get("reading") == word_data.get("reading")):
            return False, "Already saved"

    words.append(word_data)
    try:
        with open(SAVED_WORDS_FILE, "w", encoding="utf-8") as f:
            json.dump(words, f, ensure_ascii=False, indent=2)
        return True, f"Saved ({len(words)} total)"
    except IOError as e:
        return False, str(e)


def clear_saved_words():
    if os.path.exists(SAVED_WORDS_FILE):
        os.remove(SAVED_WORDS_FILE)


def send_all_to_anki():
    words = load_saved_words()
    if not words:
        return 0, 0, ["No saved words to send."]

    added, skipped, errors = 0, 0, []
    for w in words:
        success, result = add_to_anki(
            w.get("word", ""), w.get("reading", ""),
            w.get("meaning", ""), w.get("example", ""),
        )
        if success:
            added += 1
        else:
            msg = str(result)
            if "duplicate" in msg.lower():
                skipped += 1
            else:
                errors.append(f"{w.get('word', '?')}: {msg}")

    if not errors:
        clear_saved_words()
    return added, skipped, errors


# ── Markdown → HTML converter ─────────────────────────────────────────────────

def markdown_to_html(md: str) -> str:
    """Convert simplified markdown to styled HTML for QTextEdit."""
    lines = md.split("\n")
    html_lines = []

    for line in lines:
        stripped = line.strip()

        if re.match(r'^[-*_]{3,}$', stripped):
            html_lines.append('<hr style="border: 1px solid #333355;">')
            continue

        h_match = re.match(r'^(#{1,3})\s+(.+)$', stripped)
        if h_match:
            level = len(h_match.group(1))
            sizes = {1: 16, 2: 14, 3: 12}
            colors = {1: "#00ff88", 2: "#00cc70", 3: "#00aa58"}
            text = _inline_md_to_html(h_match.group(2))
            html_lines.append(
                f'<h{level} style="color:{colors[level]};font-size:{sizes[level]}px;'
                f'margin:6px 0 3px 0;">{text}</h{level}>'
            )
            continue

        b_match = re.match(r'^[\-*+]\s+(.+)$', stripped)
        if b_match:
            text = _inline_md_to_html(b_match.group(1))
            html_lines.append(
                f'<p style="margin:2px 0 2px 16px;color:#e0e0e0;">'
                f'<span style="color:#00ff88;">•</span> {text}</p>'
            )
            continue

        n_match = re.match(r'^(\d+)\.\s+(.+)$', stripped)
        if n_match:
            text = _inline_md_to_html(n_match.group(2))
            html_lines.append(
                f'<p style="margin:2px 0 2px 16px;color:#e0e0e0;">'
                f'{n_match.group(1)}. {text}</p>'
            )
            continue

        if stripped == "":
            html_lines.append("<br>")
            continue

        text = _inline_md_to_html(stripped)
        html_lines.append(f'<p style="margin:2px 0;color:#e0e0e0;">{text}</p>')

    return "\n".join(html_lines)


def _inline_md_to_html(text: str) -> str:
    """Convert inline markdown (bold, italic, code) to HTML."""
    text = re.sub(
        r'\*\*\*(.+?)\*\*\*',
        r'<b><i style="color:#ffffff;">\1</i></b>', text,
    )
    text = re.sub(r'\*\*(.+?)\*\*', r'<b style="color:#ffffff;">\1</b>', text)
    text = re.sub(r'\*(.+?)\*', r'<i style="color:#cccccc;">\1</i>', text)
    text = re.sub(
        r'`(.+?)`',
        r'<code style="background:#2a2a3e;color:#ffcc66;padding:1px 4px;'
        r'border-radius:3px;">\1</code>',
        text,
    )
    return text


# ── Styled colors ─────────────────────────────────────────────────────────────

_BG = "#1a1a2e"
_BG2 = "#16213e"
_ACCENT = "#00ff88"
_GOLD = "#ffcc66"
_DIM = "#888888"
_DIM2 = "#666666"
_CARD_BG = "#16213e"
_BLUE = "#0f3460"


# ── Thread → main thread signals ──────────────────────────────────────────────

class _Signals(QObject):
    """Bridge for worker threads to deliver results to the Qt main thread."""
    result_ready = pyqtSignal(str, dict, object)  # ocr_text, result, QPoint
    anki_done = pyqtSignal(int, int, list, object)  # added, skipped, errors, btn


# Singleton signals instance (created once, shared)
_signals = _Signals()


# ── Result Popup (PyQt6) ──────────────────────────────────────────────────────

class ResultPopup(QWidget):
    """Frameless popup showing OCR analysis + vocabulary cards."""

    def __init__(self, ocr_text: str, result: dict, parent_pos=None):
        super().__init__()
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.setStyleSheet(f"background-color: {_BG};")
        self.setMinimumWidth(560)
        self.setMaximumWidth(700)

        self._drag_pos = None

        analysis = result.get("analysis", "")
        vocab = result.get("vocab", [])

        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(6)

        # ── Title bar ──
        title_bar = QHBoxLayout()
        title_label = QLabel("📖 Japanese Dictionary")
        title_label.setStyleSheet(
            f"color:{_ACCENT};font-size:14px;font-weight:bold;"
        )
        title_bar.addWidget(title_label)
        title_bar.addStretch()

        close_btn = QPushButton("✕")
        close_btn.setFixedSize(28, 28)
        close_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        close_btn.setStyleSheet(f"""
            QPushButton {{
                background: transparent; color: {_DIM2};
                font-size: 14px; border: none; border-radius: 4px;
            }}
            QPushButton:hover {{ background: #2a2a4e; color: #ff6666; }}
        """)
        close_btn.clicked.connect(self.close)
        title_bar.addWidget(close_btn)
        layout.addLayout(title_bar)

        layout.addWidget(self._separator())

        # ── OCR text banner ──
        if ocr_text:
            ocr_label = QLabel(f"OCR:  {ocr_text}")
            ocr_label.setWordWrap(True)
            ocr_label.setStyleSheet(
                f"background:{_BLUE};color:#e0e0ff;font-size:13px;"
                f"padding:8px 12px;border-radius:6px;"
            )
            layout.addWidget(ocr_label)

        # ── Analysis (scrollable rich text) ──
        analysis_view = QTextEdit()
        analysis_view.setReadOnly(True)
        analysis_view.setHtml(markdown_to_html(analysis))
        analysis_view.setStyleSheet(f"""
            QTextEdit {{
                background: {_BG}; border: none; padding: 4px;
                selection-background-color: {_BLUE};
            }}
            QScrollBar:vertical {{
                background: {_BG}; width: 8px; border: none;
            }}
            QScrollBar::handle:vertical {{
                background: #333355; border-radius: 4px; min-height: 30px;
            }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
                height: 0px;
            }}
        """)
        analysis_view.setMinimumHeight(200)
        analysis_view.setMaximumHeight(350)
        layout.addWidget(analysis_view)

        # ── Vocabulary Cards ──
        if vocab:
            layout.addWidget(self._separator())

            vocab_header = QHBoxLayout()
            vl = QLabel("🃏 Vocabulary Cards")
            vl.setStyleSheet(
                f"color:{_GOLD};font-size:13px;font-weight:bold;"
            )
            vocab_header.addWidget(vl)
            vocab_header.addStretch()
            hint = QLabel("Click ＋ to save locally")
            hint.setStyleSheet(f"color:{_DIM2};font-size:10px;")
            vocab_header.addWidget(hint)
            layout.addLayout(vocab_header)

            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setHorizontalScrollBarPolicy(
                Qt.ScrollBarPolicy.ScrollBarAsNeeded
            )
            scroll.setVerticalScrollBarPolicy(
                Qt.ScrollBarPolicy.ScrollBarAlwaysOff
            )
            scroll.setStyleSheet(f"""
                QScrollArea {{ background: {_BG}; border: none; }}
                QScrollBar:horizontal {{
                    background: {_BG}; height: 8px; border: none;
                }}
                QScrollBar::handle:horizontal {{
                    background: #333355; border-radius: 4px; min-width: 30px;
                }}
                QScrollBar::add-line:horizontal,
                QScrollBar::sub-line:horizontal {{ width: 0px; }}
            """)
            scroll.setFixedHeight(210)

            cards_container = QWidget()
            cards_layout = QHBoxLayout(cards_container)
            cards_layout.setContentsMargins(0, 0, 0, 0)
            cards_layout.setSpacing(10)

            saved_words = load_saved_words()
            for word_data in vocab:
                card = self._make_card(word_data, saved_words)
                cards_layout.addWidget(card)
            cards_layout.addStretch()

            scroll.setWidget(cards_container)
            layout.addWidget(scroll)

        # ── Send All to Anki ──
        layout.addWidget(self._separator())

        send_row = QHBoxLayout()
        saved_count = len(load_saved_words())
        self._count_label = QLabel(
            f"📦 {saved_count} word{'s' if saved_count != 1 else ''} saved"
        )
        self._count_label.setStyleSheet(f"color:{_DIM};font-size:11px;")
        send_row.addWidget(self._count_label)
        send_row.addStretch()

        self._send_btn = QPushButton("📤 Send all words to Anki")
        self._send_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self._send_btn.setStyleSheet(f"""
            QPushButton {{
                background: {_BLUE}; color: {_GOLD};
                font-size: 12px; font-weight: bold;
                padding: 6px 16px; border: none; border-radius: 5px;
            }}
            QPushButton:hover {{ background: #1a5276; color: #ffe066; }}
        """)
        self._send_btn.clicked.connect(self._handle_send_all)
        send_row.addWidget(self._send_btn)
        layout.addLayout(send_row)

        # Connect the anki_done signal
        _signals.anki_done.connect(self._on_anki_done)

        # ── Footer ──
        footer = QLabel("Press Escape or click ✕ to close")
        footer.setStyleSheet("color:#555555;font-size:9px;")
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(footer)

        # Position near the selection
        self.adjustSize()
        if parent_pos:
            screen = QApplication.primaryScreen().geometry()
            px = min(parent_pos.x() + 10, screen.width() - self.width() - 20)
            py = max(parent_pos.y() - self.height() - 10, 10)
            if py + self.height() > screen.height():
                py = screen.height() - self.height() - 40
            self.move(px, py)

    # ── Helpers ──

    @staticmethod
    def _separator():
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"color:{_BG2};")
        sep.setFixedHeight(2)
        return sep

    def _make_card(self, word_data: dict, saved_words: list) -> QFrame:
        word = word_data.get("word", "?")
        reading = word_data.get("reading", "")
        katakana = word_data.get("katakana", "")
        meaning = word_data.get("meaning", "")
        pos = word_data.get("pos", "")
        freq = word_data.get("frequency", "")
        freq_rank = word_data.get("frequency_rank", 0)
        example = word_data.get("example", "")
        example_en = word_data.get("example_en", "")

        card = QFrame()
        card.setFixedWidth(240)
        card.setStyleSheet(
            f"QFrame {{ background:{_CARD_BG}; border:1px solid #222244;"
            f"border-radius:8px; padding:10px; }}"
        )
        cl = QVBoxLayout(card)
        cl.setContentsMargins(10, 8, 10, 8)
        cl.setSpacing(3)

        # Word
        wl = QLabel(word)
        wl.setStyleSheet("color:#ffffff;font-size:18px;font-weight:bold;")
        cl.addWidget(wl)

        # Reading
        r_text = reading
        if katakana and katakana != reading:
            r_text += f"  【{katakana}】"
        rl = QLabel(r_text)
        rl.setStyleSheet(f"color:{_ACCENT};font-size:12px;")
        cl.addWidget(rl)

        # POS + meaning
        ml = QLabel(f"{pos} — {meaning}")
        ml.setWordWrap(True)
        ml.setStyleSheet("color:#cccccc;font-size:11px;")
        cl.addWidget(ml)

        # Frequency
        freq_colors = {
            "very common": _ACCENT, "common": "#66bb6a",
            "uncommon": _GOLD, "rare": "#ff6666",
        }
        fc = freq_colors.get(freq.lower(), _DIM) if freq else _DIM
        freq_text = f"📊 {freq}"
        if freq_rank:
            freq_text += f" (#{freq_rank})"
        fl = QLabel(freq_text)
        fl.setStyleSheet(f"color:{fc};font-size:10px;font-weight:bold;")
        cl.addWidget(fl)

        # Example
        if example:
            ex = example
            if example_en:
                ex += f"\n{example_en}"
            el = QLabel(ex)
            el.setWordWrap(True)
            el.setStyleSheet("color:#999999;font-size:10px;")
            cl.addWidget(el)

        # Save button
        already_saved = any(
            w.get("word") == word and w.get("reading") == reading
            for w in saved_words
        )

        save_btn = QPushButton("✓ Saved" if already_saved else "＋ Save")
        save_btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        if already_saved:
            save_btn.setStyleSheet(
                f"QPushButton {{ background:#1a4d1a;color:{_ACCENT};"
                f"font-size:11px;font-weight:bold;padding:4px;border:none;"
                f"border-radius:4px; }}"
            )
            save_btn.setEnabled(False)
        else:
            save_btn.setStyleSheet(f"""
                QPushButton {{
                    background:{_BLUE};color:{_ACCENT};
                    font-size:11px;font-weight:bold;padding:4px;
                    border:none;border-radius:4px;
                }}
                QPushButton:hover {{ background:#1a5276;color:#44ffaa; }}
            """)
            save_btn.clicked.connect(
                lambda _, b=save_btn, d=word_data: self._save_word(b, d)
            )

        cl.addWidget(save_btn)
        return card

    def _save_word(self, btn: QPushButton, word_data: dict):
        success, msg = save_word_locally(word_data)
        if success:
            btn.setText("✓ Saved")
            btn.setStyleSheet(
                f"QPushButton {{ background:#1a4d1a;color:{_ACCENT};"
                f"font-size:11px;font-weight:bold;padding:4px;border:none;"
                f"border-radius:4px; }}"
            )
            btn.setEnabled(False)
            self._refresh_count()
        elif "Already saved" in msg:
            btn.setText("✓ Already saved")
            btn.setStyleSheet(
                "QPushButton { background:#4d3a1a;color:#ffcc66;"
                "font-size:11px;font-weight:bold;padding:4px;border:none;"
                "border-radius:4px; }"
            )
            btn.setEnabled(False)
        else:
            btn.setText(f"✗ {msg[:30]}")
            btn.setStyleSheet(
                "QPushButton { background:#4d1a1a;color:#ff6666;"
                "font-size:11px;font-weight:bold;padding:4px;border:none;"
                "border-radius:4px; }"
            )

    def _refresh_count(self):
        count = len(load_saved_words())
        self._count_label.setText(
            f"📦 {count} word{'s' if count != 1 else ''} saved"
        )

    def _handle_send_all(self):
        btn = self._send_btn
        btn.setText("⏳ Sending...")
        btn.setEnabled(False)

        def worker():
            added, skipped, errors = send_all_to_anki()
            _signals.anki_done.emit(added, skipped, errors, btn)

        threading.Thread(target=worker, daemon=True).start()

    def _on_anki_done(self, added, skipped, errors, btn):
        btn.setEnabled(True)
        if errors and errors != ["No saved words to send."]:
            btn.setText(f"✗ {len(errors)} error(s)")
            for e in errors:
                print(f"Anki error: {e}")
        elif added == 0 and skipped == 0:
            btn.setText("No words to send")
        else:
            msg = f"✓ {added} added"
            if skipped:
                msg += f", {skipped} duplicates"
            btn.setText(msg)
            self._refresh_count()

    # ── Dragging ──

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = (
                event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            )

    def mouseMoveEvent(self, event):
        if self._drag_pos and event.buttons() & Qt.MouseButton.LeftButton:
            self.move(event.globalPosition().toPoint() - self._drag_pos)

    def mouseReleaseEvent(self, event):
        self._drag_pos = None

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.close()


# ── Loading Popup ─────────────────────────────────────────────────────────────

class LoadingPopup(QWidget):
    def __init__(self, pos=None):
        super().__init__()
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
        )
        self.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
        self.setStyleSheet(f"background-color:{_BG};border-radius:10px;")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 18, 24, 18)

        title = QLabel("🔍 Analyzing…")
        title.setStyleSheet(
            f"color:{_ACCENT};font-size:15px;font-weight:bold;"
        )
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        sub = QLabel("Capturing & running OCR + LLM")
        sub.setStyleSheet(f"color:{_DIM};font-size:11px;")
        sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(sub)

        self.adjustSize()
        if pos:
            screen = QApplication.primaryScreen().geometry()
            px = min(pos.x() + 10, screen.width() - self.width() - 20)
            py = max(pos.y() - self.height() - 10, 10)
            self.move(px, py)


# ── Snipping Tool (Wayland-compatible) ────────────────────────────────────────

class SnippingTool(QWidget):
    def __init__(self, bg_image_path):
        super().__init__()
        self.bg_image_path = bg_image_path

        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self.showFullScreen()
        self.setCursor(Qt.CursorShape.CrossCursor)

        self.background = QPixmap(self.bg_image_path)
        self.begin = None
        self.end = None

        self._loading_popup = None
        self._result_popup = None

        # Connect the worker signal
        _signals.result_ready.connect(self._on_result_ready)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.drawPixmap(self.rect(), self.background)

        if self.begin and self.end:
            rect = QRect(self.begin, self.end).normalized()
            overlay_color = QColor(0, 0, 0, 120)

            # Top
            painter.fillRect(0, 0, self.width(), rect.top(), overlay_color)
            # Bottom
            painter.fillRect(
                0, rect.bottom() + 1, self.width(),
                self.height() - rect.bottom(), overlay_color,
            )
            # Left
            painter.fillRect(
                0, rect.top(), rect.left(), rect.height() + 1, overlay_color,
            )
            # Right
            painter.fillRect(
                rect.right() + 1, rect.top(),
                self.width() - rect.right(), rect.height() + 1, overlay_color,
            )

            pen = QPen(QColor(0, 255, 136))
            pen.setWidth(2)
            painter.setPen(pen)
            painter.drawRect(rect)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.begin = event.pos()
            self.end = self.begin
            self.update()

    def mouseMoveEvent(self, event):
        if self.begin:
            self.end = event.pos()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.end = event.pos()
            self._capture_and_process()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Escape:
            self.close()

    def _capture_and_process(self):
        rect = QRect(self.begin, self.end).normalized()
        if rect.width() < 10 or rect.height() < 10:
            self.close()
            return

        # Map widget coords to screen coords for popup placement
        top_left = self.mapToGlobal(rect.topLeft())
        popup_pos = QPoint(top_left.x(), top_left.y())

        # Crop from the background pixmap (already a full screenshot)
        cropped_pixmap = self.background.copy(rect)
        temp_crop_path = "/tmp/jap_ocr_crop.png"
        cropped_pixmap.save(temp_crop_path)

        # Close the fullscreen overlay
        self.hide()

        # Show loading popup
        self._loading_popup = LoadingPopup(popup_pos)
        self._loading_popup.show()

        # Run OCR + LLM in background thread
        def worker():
            try:
                from PIL import Image
                img = Image.open(temp_crop_path)
                ocr = get_ocr()
                text = ocr(img)
                print(f"OCR result: {text}")

                if not text or text.strip() == "":
                    result = {
                        "analysis": "⚠️ No text detected in the selected region.",
                        "vocab": [],
                    }
                else:
                    result = query_openrouter(text.strip())

                _signals.result_ready.emit(text or "", result, popup_pos)

            except Exception as e:
                _signals.result_ready.emit(
                    "", {"analysis": f"⚠️ Error: {e}", "vocab": []}, popup_pos,
                )

        threading.Thread(target=worker, daemon=True).start()

    def _on_result_ready(self, ocr_text: str, result: dict, pos):
        # Close loading popup
        if self._loading_popup:
            self._loading_popup.close()
            self._loading_popup = None

        # Show result popup
        self._result_popup = ResultPopup(ocr_text, result, pos)
        self._result_popup.show()


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if not OPENROUTER_API_KEY or OPENROUTER_API_KEY == "your_api_key_here":
        print("⚠️  Please set your OPENROUTER_API_KEY in .env file!")
        print("   Copy .env.example to .env and add your key.")
        print()

    # Force PyQt to use the native Wayland backend
    os.environ["QT_QPA_PLATFORM"] = "wayland"

    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 10))

    temp_path = "/tmp/wayland_full_temp.png"

    # 1. Grab the underlying screen
    try:
        im = PIL.ImageGrab.grab()
        im.save(temp_path)
    except Exception as e:
        print(f"Error: PIL failed to take the screenshot. Details: {e}")
        sys.exit(1)

    # 2. Launch the snipping overlay
    window = SnippingTool(temp_path)
    window.show()

    print("✅ Japanese Dictionary Capture is running.")
    print("   Left-click and drag to select a region.")
    print("   Press Ctrl+C in this terminal to quit.")

    try:
        app.exec()
    except KeyboardInterrupt:
        print("\nShutting down.")

    # 3. Clean up
    if os.path.exists(temp_path):
        os.remove(temp_path)