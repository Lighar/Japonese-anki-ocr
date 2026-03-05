"""
Japanese Dictionary Screen Capture Tool

Hold right-click and drag to select a region of the screen.
On release, OCR extracts Japanese text and queries OpenRouter LLM
for dictionary-style word information displayed in a popup with Anki cards.
"""

import os
import sys
import math
import re
import json
import threading
import tkinter as tk
from tkinter import font as tkfont
import requests
from PIL import ImageGrab
from pynput import mouse
from dotenv import load_dotenv

load_dotenv()

# ── Configuration ──────────────────────────────────────────────────────────────
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "")
OPENROUTER_MODEL = "google/gemini-2.0-flash-001"
MIN_DRAG_DISTANCE = 50  # pixels
ANKI_CONNECT_URL = "http://localhost:8765"
ANKI_DECK_NAME = "Japanese"
ANKI_MODEL_NAME = None  # Auto-detected at runtime
SAVED_WORDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "saved_words.json")

# ── Lazy-loaded OCR model ──────────────────────────────────────────────────────
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


# ── Markdown-to-Tkinter Renderer ──────────────────────────────────────────────

def render_markdown_to_text_widget(text_widget, markdown_text):
    """Parse markdown and insert styled text into a tk.Text widget."""
    text_widget.tag_configure("h1", font=("Segoe UI", 14, "bold"), foreground="#00ff88",
                              spacing3=6)
    text_widget.tag_configure("h2", font=("Segoe UI", 12, "bold"), foreground="#00cc70",
                              spacing3=4)
    text_widget.tag_configure("h3", font=("Segoe UI", 11, "bold"), foreground="#00aa58",
                              spacing3=3)
    text_widget.tag_configure("bold", font=("Segoe UI", 10, "bold"), foreground="#ffffff")
    text_widget.tag_configure("italic", font=("Segoe UI", 10, "italic"), foreground="#cccccc")
    text_widget.tag_configure("bold_italic", font=("Segoe UI", 10, "bold italic"),
                              foreground="#ffffff")
    text_widget.tag_configure("code", font=("Consolas", 10), foreground="#ffcc66",
                              background="#2a2a3e")
    text_widget.tag_configure("bullet", foreground="#00ff88", lmargin1=10, lmargin2=24)
    text_widget.tag_configure("normal", font=("Segoe UI", 10), foreground="#e0e0e0")
    text_widget.tag_configure("hr", foreground="#333355")

    lines = markdown_text.split("\n")

    for line in lines:
        stripped = line.strip()

        if re.match(r'^[-*_]{3,}$', stripped):
            text_widget.insert(tk.END, "─" * 50 + "\n", "hr")
            continue

        header_match = re.match(r'^(#{1,3})\s+(.+)$', stripped)
        if header_match:
            level = len(header_match.group(1))
            text_widget.insert(tk.END, header_match.group(2) + "\n", f"h{level}")
            continue

        bullet_match = re.match(r'^[\-*+]\s+(.+)$', stripped)
        if bullet_match:
            _insert_inline_markdown(text_widget, "  • " + bullet_match.group(1), "bullet")
            text_widget.insert(tk.END, "\n")
            continue

        num_match = re.match(r'^(\d+)\.\s+(.+)$', stripped)
        if num_match:
            _insert_inline_markdown(text_widget, num_match.group(1) + ". " + num_match.group(2), "normal")
            text_widget.insert(tk.END, "\n")
            continue

        if stripped == "":
            text_widget.insert(tk.END, "\n")
            continue

        _insert_inline_markdown(text_widget, stripped, "normal")
        text_widget.insert(tk.END, "\n")


def _insert_inline_markdown(text_widget, text, base_tag):
    """Parse inline markdown (bold, italic, code) and insert with tags."""
    pattern = re.compile(
        r'(\*\*\*(.+?)\*\*\*)'
        r'|(\*\*(.+?)\*\*)'
        r'|(\*(.+?)\*)'
        r'|(`(.+?)`)'
    )

    last_end = 0
    for match in pattern.finditer(text):
        if match.start() > last_end:
            text_widget.insert(tk.END, text[last_end:match.start()], base_tag)

        if match.group(2):
            text_widget.insert(tk.END, match.group(2), "bold_italic")
        elif match.group(4):
            text_widget.insert(tk.END, match.group(4), "bold")
        elif match.group(6):
            text_widget.insert(tk.END, match.group(6), "italic")
        elif match.group(8):
            text_widget.insert(tk.END, match.group(8), "code")

        last_end = match.end()

    if last_end < len(text):
        text_widget.insert(tk.END, text[last_end:], base_tag)


# ── AnkiConnect Integration ──────────────────────────────────────────────────

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

    # Get all model names
    success, models = anki_connect_request("modelNames")
    if not success or not models:
        ANKI_MODEL_NAME = "Basic"
        return ANKI_MODEL_NAME

    # Check each model for Front/Back fields
    for model in models:
        ok, fields = anki_connect_request("modelFieldNames", modelName=model)
        if ok and fields and "Front" in fields and "Back" in fields:
            ANKI_MODEL_NAME = model
            print(f"Using Anki model: {model}")
            return ANKI_MODEL_NAME

    # No suitable model found — create a custom one
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
        # Last resort fallback
        ANKI_MODEL_NAME = models[0] if models else "Basic"
        print(f"Falling back to Anki model: {ANKI_MODEL_NAME}")

    return ANKI_MODEL_NAME


def add_to_anki(word, reading, meaning, example=""):
    """Add a card to Anki via AnkiConnect."""
    front = f"{word}\n{reading}"
    back = f"{meaning}"
    if example:
        back += f"\n\n{example}"

    # Ensure the deck exists
    anki_connect_request("createDeck", deck=ANKI_DECK_NAME)

    model_name = _get_anki_model_name()

    success, result = anki_connect_request(
        "addNote",
        note={
            "deckName": ANKI_DECK_NAME,
            "modelName": model_name,
            "fields": {
                "Front": front,
                "Back": back,
            },
            "options": {"allowDuplicate": False},
            "tags": ["jap-capture"],
        }
    )
    return success, result


# ── Local Word Storage ───────────────────────────────────────────────────────

def load_saved_words():
    """Load saved words from the local JSON file."""
    if os.path.exists(SAVED_WORDS_FILE):
        try:
            with open(SAVED_WORDS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            return []
    return []


def save_word_locally(word_data):
    """Append a word to the local JSON file. Returns (success, message)."""
    words = load_saved_words()

    # Check for duplicates by word + reading
    for existing in words:
        if existing.get("word") == word_data.get("word") and \
           existing.get("reading") == word_data.get("reading"):
            return False, "Already saved"

    words.append(word_data)
    try:
        with open(SAVED_WORDS_FILE, "w", encoding="utf-8") as f:
            json.dump(words, f, ensure_ascii=False, indent=2)
        return True, f"Saved ({len(words)} total)"
    except IOError as e:
        return False, str(e)


def clear_saved_words():
    """Clear the local saved words file."""
    if os.path.exists(SAVED_WORDS_FILE):
        os.remove(SAVED_WORDS_FILE)


def send_all_to_anki():
    """Send all saved words to Anki and clear the file on success.
    Returns (added_count, skipped_count, errors)."""
    words = load_saved_words()
    if not words:
        return 0, 0, ["No saved words to send."]

    added = 0
    skipped = 0
    errors = []

    for w in words:
        word = w.get("word", "")
        reading = w.get("reading", "")
        meaning = w.get("meaning", "")
        example = w.get("example", "")

        success, result = add_to_anki(word, reading, meaning, example)
        if success:
            added += 1
        else:
            msg = str(result)
            if "duplicate" in msg.lower():
                skipped += 1
            else:
                errors.append(f"{word}: {msg}")

    # Clear file if everything was processed (even if some were duplicates)
    if not errors:
        clear_saved_words()

    return added, skipped, errors


# ── OpenRouter LLM API ─────────────────────────────────────────────────────────

def query_openrouter(text: str) -> dict:
    """Send extracted Japanese text to OpenRouter for full sentence analysis + vocab cards.
    
    Returns a dict with 'analysis' (markdown string) and 'vocab' (list of word dicts).
    """
    if not OPENROUTER_API_KEY:
        return {
            "analysis": "⚠️ OPENROUTER_API_KEY not set.\nAdd it to .env file.",
            "vocab": []
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
            headers=headers,
            json=payload,
            timeout=30,
        )
        resp.raise_for_status()
        data = resp.json()
        content = data["choices"][0]["message"]["content"]

        # Parse out the vocab JSON from the response
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

    # Try to extract JSON between markers
    json_match = re.search(
        r'VOCABULARY_JSON_START\s*\n?(.*?)\n?\s*VOCABULARY_JSON_END',
        content, re.DOTALL
    )

    if json_match:
        json_str = json_match.group(1).strip()
        # Remove the JSON block from the analysis text
        analysis = content[:json_match.start()].strip()
        # Clean trailing --- or whitespace
        analysis = re.sub(r'\n-{3,}\s*$', '', analysis).strip()

        try:
            vocab = json.loads(json_str)
        except json.JSONDecodeError:
            # Try to fix common JSON issues
            try:
                # Sometimes LLM adds trailing commas
                cleaned = re.sub(r',\s*([}\]])', r'\1', json_str)
                vocab = json.loads(cleaned)
            except json.JSONDecodeError:
                print(f"Failed to parse vocab JSON: {json_str[:200]}")
                vocab = []
    else:
        # No markers found — try to find a JSON array at the end
        analysis = content
        json_array_match = re.search(r'(\[\s*\{.*?\}\s*\])\s*$', content, re.DOTALL)
        if json_array_match:
            try:
                vocab = json.loads(json_array_match.group(1))
                analysis = content[:json_array_match.start()].strip()
            except json.JSONDecodeError:
                pass

    return analysis, vocab


# ── Tkinter UI (overlay + popup) ──────────────────────────────────────────────

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()

        # State
        self.dragging = False
        self.start_x = 0
        self.start_y = 0
        self.current_x = 0
        self.current_y = 0
        self.overlay = None
        self.canvas = None
        self.rect_id = None
        self.popup_win = None

        # Start the global mouse listener in a daemon thread
        self.listener = mouse.Listener(
            on_click=self._on_click,
            on_move=self._on_move,
        )
        self.listener.daemon = True
        self.listener.start()

        print("✅ Japanese Dictionary Capture is running.")
        print("   Right-click and drag to select a region.")
        print("   Press Ctrl+C in this terminal to quit.")

        self.root.mainloop()

    # ── Mouse callbacks ───────────────────────────────────────────────────

    def _on_click(self, x, y, button, pressed):
        if button == mouse.Button.right:
            if pressed:
                self.start_x = x
                self.start_y = y
                self.current_x = x
                self.current_y = y
                self.dragging = True
                self.root.after(0, self._show_overlay)
            else:
                if self.dragging:
                    self.dragging = False
                    self.current_x = x
                    self.current_y = y
                    dist = math.hypot(x - self.start_x, y - self.start_y)
                    if dist >= MIN_DRAG_DISTANCE:
                        self.root.after(0, self._on_release)
                    else:
                        self.root.after(0, self._hide_overlay)

    def _on_move(self, x, y):
        if self.dragging:
            self.current_x = x
            self.current_y = y
            self.root.after(0, self._update_rect)

    # ── Overlay ───────────────────────────────────────────────────────────

    def _show_overlay(self):
        if self.overlay is not None:
            return

        self.overlay = tk.Toplevel(self.root)
        self.overlay.attributes("-fullscreen", True)
        self.overlay.attributes("-topmost", True)
        self.overlay.attributes("-alpha", 0.25)
        self.overlay.configure(bg="black")
        self.overlay.overrideredirect(True)

        self.canvas = tk.Canvas(
            self.overlay, bg="black", highlightthickness=0,
            cursor="crosshair"
        )
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.rect_id = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline="#00ff88", width=2, dash=(6, 4),
        )

    def _update_rect(self):
        if self.canvas and self.rect_id:
            self.canvas.coords(
                self.rect_id,
                self.start_x, self.start_y,
                self.current_x, self.current_y,
            )

    def _hide_overlay(self):
        if self.overlay:
            self.overlay.destroy()
            self.overlay = None
            self.canvas = None
            self.rect_id = None

    # ── Capture + OCR + LLM ───────────────────────────────────────────────

    def _on_release(self):
        x1 = min(self.start_x, self.current_x)
        y1 = min(self.start_y, self.current_y)
        x2 = max(self.start_x, self.current_x)
        y2 = max(self.start_y, self.current_y)

        self._hide_overlay()
        self.root.after(150, lambda: self._capture_and_process(x1, y1, x2, y2))

    def _capture_and_process(self, x1, y1, x2, y2):
        self._show_loading_popup(x1, y1)

        def worker():
            try:
                img = ImageGrab.grab(bbox=(x1, y1, x2, y2))
                ocr = get_ocr()
                text = ocr(img)
                print(f"OCR result: {text}")

                if not text or text.strip() == "":
                    result = {
                        "analysis": "⚠️ No text detected in the selected region.",
                        "vocab": []
                    }
                else:
                    result = query_openrouter(text.strip())

                self.root.after(0, lambda: self._show_result_popup(
                    x1, y1, text, result
                ))

            except Exception as e:
                err_result = {
                    "analysis": f"⚠️ Error: {e}",
                    "vocab": []
                }
                self.root.after(0, lambda: self._show_result_popup(
                    x1, y1, "", err_result
                ))

        threading.Thread(target=worker, daemon=True).start()

    # ── Popup windows ────────────────────────────────────────────────────

    def _show_loading_popup(self, x, y):
        self._close_popup()

        win = tk.Toplevel(self.root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.configure(bg="#1a1a2e")

        frame = tk.Frame(win, bg="#1a1a2e", padx=20, pady=15)
        frame.pack()

        tk.Label(
            frame, text="🔍 Analyzing…",
            bg="#1a1a2e", fg="#00ff88",
            font=("Segoe UI", 13, "bold"),
        ).pack()

        tk.Label(
            frame, text="Capturing & running OCR + LLM",
            bg="#1a1a2e", fg="#888888",
            font=("Segoe UI", 9),
        ).pack(pady=(4, 0))

        win.update_idletasks()
        w = win.winfo_width()
        h = win.winfo_height()
        screen_w = win.winfo_screenwidth()
        px = min(x + 10, screen_w - w - 20)
        py = max(y - h - 10, 10)
        win.geometry(f"+{px}+{py}")

        self.popup_win = win

    def _show_result_popup(self, x, y, ocr_text, result):
        """Show the final result popup with analysis + Anki vocab cards."""
        self._close_popup()

        analysis = result.get("analysis", "")
        vocab = result.get("vocab", [])

        win = tk.Toplevel(self.root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        win.configure(bg="#1a1a2e")

        # Main frame
        outer = tk.Frame(win, bg="#16213e", bd=1, relief=tk.SOLID)
        outer.pack(padx=1, pady=1)

        frame = tk.Frame(outer, bg="#1a1a2e", padx=16, pady=12)
        frame.pack(fill=tk.BOTH, expand=True)

        # ── Title bar ──
        title_frame = tk.Frame(frame, bg="#1a1a2e")
        title_frame.pack(fill=tk.X, pady=(0, 8))

        tk.Label(
            title_frame, text="📖 Japanese Dictionary",
            bg="#1a1a2e", fg="#00ff88",
            font=("Segoe UI", 12, "bold"),
        ).pack(side=tk.LEFT)

        close_btn = tk.Label(
            title_frame, text="✕", bg="#1a1a2e", fg="#666666",
            font=("Segoe UI", 12), cursor="hand2",
        )
        close_btn.pack(side=tk.RIGHT)
        close_btn.bind("<Button-1>", lambda e: self._close_popup())

        tk.Frame(frame, bg="#16213e", height=1).pack(fill=tk.X, pady=(0, 8))

        # ── OCR text banner ──
        if ocr_text:
            ocr_frame = tk.Frame(frame, bg="#0f3460", padx=10, pady=6)
            ocr_frame.pack(fill=tk.X, pady=(0, 8))
            tk.Label(
                ocr_frame, text=f"OCR: {ocr_text}",
                bg="#0f3460", fg="#e0e0ff",
                font=("Segoe UI", 11),
                wraplength=500, justify=tk.LEFT,
            ).pack(anchor=tk.W)

        # ── Analysis text (scrollable) ──
        text_frame = tk.Frame(frame, bg="#1a1a2e")
        text_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text_widget = tk.Text(
            text_frame,
            bg="#1a1a2e", fg="#e0e0e0",
            font=("Segoe UI", 10),
            wrap=tk.WORD,
            width=65, height=16,
            bd=0, highlightthickness=0,
            yscrollcommand=scrollbar.set,
            selectbackground="#0f3460",
            cursor="arrow",
            padx=4, pady=4,
        )
        render_markdown_to_text_widget(text_widget, analysis)
        text_widget.config(state=tk.DISABLED)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=text_widget.yview)

        # ── Vocabulary Cards Section ──
        if vocab:
            tk.Frame(frame, bg="#16213e", height=1).pack(fill=tk.X, pady=(10, 6))

            vocab_header = tk.Frame(frame, bg="#1a1a2e")
            vocab_header.pack(fill=tk.X, pady=(0, 6))
            tk.Label(
                vocab_header, text="🃏 Vocabulary Cards",
                bg="#1a1a2e", fg="#ffcc66",
                font=("Segoe UI", 11, "bold"),
            ).pack(side=tk.LEFT)
            tk.Label(
                vocab_header, text="Click to save locally",
                bg="#1a1a2e", fg="#666666",
                font=("Segoe UI", 8),
            ).pack(side=tk.RIGHT)

            # Scrollable cards container
            cards_canvas = tk.Canvas(
                frame, bg="#1a1a2e", highlightthickness=0, height=160,
            )
            cards_scrollbar = tk.Scrollbar(
                frame, orient=tk.HORIZONTAL, command=cards_canvas.xview,
            )
            cards_inner = tk.Frame(cards_canvas, bg="#1a1a2e")

            cards_canvas.configure(xscrollcommand=cards_scrollbar.set)
            cards_scrollbar.pack(fill=tk.X, side=tk.BOTTOM)
            cards_canvas.pack(fill=tk.X, side=tk.TOP)

            cards_window = cards_canvas.create_window(
                (0, 0), window=cards_inner, anchor="nw"
            )

            for i, word_data in enumerate(vocab):
                self._create_vocab_card(cards_inner, word_data, i)

            # Update scroll region after cards are rendered
            def _on_cards_configure(event):
                cards_canvas.configure(scrollregion=cards_canvas.bbox("all"))
                cards_canvas.configure(height=cards_inner.winfo_reqheight())

            cards_inner.bind("<Configure>", _on_cards_configure)

            # Mouse wheel horizontal scroll on cards
            def _on_mousewheel(event):
                cards_canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
            cards_canvas.bind("<MouseWheel>", _on_mousewheel)

        # ── Send All to Anki button ──
        tk.Frame(frame, bg="#16213e", height=1).pack(fill=tk.X, pady=(10, 6))

        send_frame = tk.Frame(frame, bg="#1a1a2e")
        send_frame.pack(fill=tk.X, pady=(0, 4))

        # Word count badge
        saved_count = len(load_saved_words())
        count_text = f"📦 {saved_count} word{'s' if saved_count != 1 else ''} saved"
        self._saved_count_label = tk.Label(
            send_frame, text=count_text,
            bg="#1a1a2e", fg="#888888",
            font=("Segoe UI", 9),
        )
        self._saved_count_label.pack(side=tk.LEFT)

        send_btn = tk.Label(
            send_frame, text="📤 Send all words to Anki",
            bg="#0f3460", fg="#ffcc66",
            font=("Segoe UI", 10, "bold"),
            padx=14, pady=5,
            cursor="hand2",
        )
        send_btn.pack(side=tk.RIGHT)

        def _on_send_click(event):
            self._handle_send_all_to_anki(send_btn)
        send_btn.bind("<Button-1>", _on_send_click)

        def _on_send_enter(event):
            send_btn.configure(bg="#1a5276", fg="#ffe066")
        def _on_send_leave(event):
            send_btn.configure(bg="#0f3460", fg="#ffcc66")
        send_btn.bind("<Enter>", _on_send_enter)
        send_btn.bind("<Leave>", _on_send_leave)

        # ── Hint ──
        tk.Label(
            frame, text="Press Escape or click ✕ to close",
            bg="#1a1a2e", fg="#555555",
            font=("Segoe UI", 8),
        ).pack(pady=(8, 0))

        # Position near selection area
        win.update_idletasks()
        w = win.winfo_reqwidth()
        h = win.winfo_reqheight()
        screen_w = win.winfo_screenwidth()
        screen_h = win.winfo_screenheight()
        px = min(x + 10, screen_w - w - 20)
        py = max(y - h - 10, 10)
        if py + h > screen_h:
            py = screen_h - h - 40
        win.geometry(f"+{px}+{py}")

        win.bind("<Escape>", lambda e: self._close_popup())
        win.focus_force()
        self._make_draggable(win, title_frame)

        self.popup_win = win

    def _create_vocab_card(self, parent, word_data, index):
        """Create a single vocabulary card widget."""
        word = word_data.get("word", "?")
        reading = word_data.get("reading", "")
        katakana = word_data.get("katakana", "")
        meaning = word_data.get("meaning", "")
        pos = word_data.get("pos", "")
        freq = word_data.get("frequency", "")
        freq_rank = word_data.get("frequency_rank", 0)
        example = word_data.get("example", "")
        example_en = word_data.get("example_en", "")

        # Card frame
        card = tk.Frame(parent, bg="#16213e", padx=12, pady=8, bd=1, relief=tk.SOLID)
        card.pack(side=tk.LEFT, padx=(0, 8), pady=2)

        # Word (large)
        tk.Label(
            card, text=word,
            bg="#16213e", fg="#ffffff",
            font=("Segoe UI", 16, "bold"),
        ).pack(anchor=tk.W)

        # Reading
        reading_text = reading
        if katakana and katakana != reading:
            reading_text += f"  【{katakana}】"
        tk.Label(
            card, text=reading_text,
            bg="#16213e", fg="#00ff88",
            font=("Segoe UI", 10),
        ).pack(anchor=tk.W)

        # Part of speech + meaning
        tk.Label(
            card, text=f"{pos} — {meaning}",
            bg="#16213e", fg="#cccccc",
            font=("Segoe UI", 9),
            wraplength=220, justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=(2, 0))

        # Frequency badge
        freq_frame = tk.Frame(card, bg="#16213e")
        freq_frame.pack(anchor=tk.W, pady=(4, 0))

        # Color based on frequency
        freq_colors = {
            "very common": "#00ff88",
            "common": "#66bb6a",
            "uncommon": "#ffcc66",
            "rare": "#ff6666",
        }
        freq_color = freq_colors.get(freq.lower(), "#888888") if freq else "#888888"

        freq_label = f"📊 {freq}"
        if freq_rank:
            freq_label += f" (#{freq_rank})"

        tk.Label(
            freq_frame, text=freq_label,
            bg="#16213e", fg=freq_color,
            font=("Segoe UI", 8, "bold"),
        ).pack(side=tk.LEFT)

        # Example (if available)
        if example:
            ex_text = example
            if example_en:
                ex_text += f"\n{example_en}"
            tk.Label(
                card, text=ex_text,
                bg="#16213e", fg="#999999",
                font=("Segoe UI", 8),
                wraplength=220, justify=tk.LEFT,
            ).pack(anchor=tk.W, pady=(3, 0))

        # ── Save button ──
        btn_frame = tk.Frame(card, bg="#16213e")
        btn_frame.pack(fill=tk.X, pady=(6, 0))

        save_btn = tk.Label(
            btn_frame, text="＋ Save",
            bg="#0f3460", fg="#00ff88",
            font=("Segoe UI", 9, "bold"),
            padx=10, pady=3,
            cursor="hand2",
        )
        save_btn.pack(fill=tk.X)

        # Check if already saved
        saved_words = load_saved_words()
        already_saved = any(
            w.get("word") == word and w.get("reading") == reading
            for w in saved_words
        )
        if already_saved:
            save_btn.configure(text="✓ Saved", bg="#1a4d1a", fg="#00ff88", cursor="arrow")
        else:
            # Bind click
            def _on_save_click(event, data=word_data, btn=save_btn):
                self._handle_save_word(btn, data)
            save_btn.bind("<Button-1>", _on_save_click)

            # Hover effects
            def _on_enter(event, btn=save_btn):
                btn.configure(bg="#1a5276", fg="#44ffaa")
            def _on_leave(event, btn=save_btn):
                btn.configure(bg="#0f3460", fg="#00ff88")
            save_btn.bind("<Enter>", _on_enter)
            save_btn.bind("<Leave>", _on_leave)

    def _handle_save_word(self, btn, word_data):
        """Save a word to the local JSON file."""
        success, msg = save_word_locally(word_data)
        if success:
            btn.configure(text="✓ Saved", bg="#1a4d1a", fg="#00ff88", cursor="arrow")
            btn.unbind("<Button-1>")
            btn.unbind("<Enter>")
            btn.unbind("<Leave>")
            # Update the saved count label
            self._update_saved_count()
        else:
            if "Already saved" in msg:
                btn.configure(text="✓ Already saved", bg="#4d3a1a", fg="#ffcc66", cursor="arrow")
                btn.unbind("<Button-1>")
            else:
                btn.configure(text=f"✗ {msg[:30]}", bg="#4d1a1a", fg="#ff6666")
                self.root.after(2000, lambda: btn.configure(
                    text="＋ Save", bg="#0f3460", fg="#00ff88",
                ))

    def _update_saved_count(self):
        """Update the saved word count label."""
        if hasattr(self, '_saved_count_label') and self._saved_count_label.winfo_exists():
            count = len(load_saved_words())
            self._saved_count_label.configure(
                text=f"📦 {count} word{'s' if count != 1 else ''} saved"
            )

    def _handle_send_all_to_anki(self, btn):
        """Send all saved words to Anki."""
        btn.configure(text="⏳ Sending...", fg="#ffcc66")
        btn.update()

        def worker():
            added, skipped, errors = send_all_to_anki()

            if errors and errors != ["No saved words to send."]:
                msg = f"✗ {len(errors)} error(s)"
                self.root.after(0, lambda: btn.configure(
                    text=msg, bg="#4d1a1a", fg="#ff6666"
                ))
                for e in errors:
                    print(f"Anki error: {e}")
                self.root.after(3000, lambda: btn.configure(
                    text="📤 Send all words to Anki", bg="#0f3460", fg="#ffcc66"
                ))
            elif added == 0 and skipped == 0:
                self.root.after(0, lambda: btn.configure(
                    text="No words to send", bg="#4d3a1a", fg="#ffcc66"
                ))
                self.root.after(2000, lambda: btn.configure(
                    text="📤 Send all words to Anki", bg="#0f3460", fg="#ffcc66"
                ))
            else:
                msg = f"✓ {added} added"
                if skipped:
                    msg += f", {skipped} duplicates"
                self.root.after(0, lambda: btn.configure(
                    text=msg, bg="#1a4d1a", fg="#00ff88"
                ))
                self.root.after(0, self._update_saved_count)
                self.root.after(3000, lambda: btn.configure(
                    text="📤 Send all words to Anki", bg="#0f3460", fg="#ffcc66"
                ))

        threading.Thread(target=worker, daemon=True).start()

    def _make_draggable(self, window, handle):
        state = {"x": 0, "y": 0}

        def on_press(event):
            state["x"] = event.x
            state["y"] = event.y

        def on_drag(event):
            dx = event.x - state["x"]
            dy = event.y - state["y"]
            x = window.winfo_x() + dx
            y = window.winfo_y() + dy
            window.geometry(f"+{x}+{y}")

        handle.bind("<ButtonPress-1>", on_press)
        handle.bind("<B1-Motion>", on_drag)

    def _close_popup(self):
        if self.popup_win:
            self.popup_win.destroy()
            self.popup_win = None


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if not OPENROUTER_API_KEY or OPENROUTER_API_KEY == "your_api_key_here":
        print("⚠️  Please set your OPENROUTER_API_KEY in .env file!")
        print("   Copy .env.example to .env and add your key.")
        print()

    try:
        App()
    except KeyboardInterrupt:
        print("\nShutting down.")
        sys.exit(0)
