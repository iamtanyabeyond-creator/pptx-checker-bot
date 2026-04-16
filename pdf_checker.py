"""
pdf_checker.py — Проверка висячих предлогов через PDF
Читает координаты слов из PDF, группирует по строкам,
проверяет последнее слово каждой строки.
"""

import re
import pdfplumber
from collections import defaultdict


# ─────────────────────────────────────────────
# НАСТРОЙКИ
# ─────────────────────────────────────────────

# Допуск по вертикали для группировки слов в одну строку (в pt)
LINE_Y_TOLERANCE = 3.0

# Минимальная длина строки — игнорируем слишком короткие (номера страниц и т.п.)
MIN_LINE_CHARS = 5

# Предлоги и союзы
PREPOSITIONS = (
    "в", "на", "с", "к", "по", "о", "и", "а", "но",
    "из", "до", "от", "за", "об", "под", "над", "при",
    "про", "для", "без", "или", "не", "то", "бы", "же",
    "со", "во", "ко",
)
HANGING_RE = re.compile(
    r'^(' + '|'.join(PREPOSITIONS) + r')$',
    re.IGNORECASE | re.UNICODE
)


# ─────────────────────────────────────────────
# ОСНОВНАЯ ЛОГИКА
# ─────────────────────────────────────────────

def group_words_into_lines(words: list, y_tolerance: float = LINE_Y_TOLERANCE) -> list:
    """
    Группирует слова в строки по близости координаты y (верхнего края).
    Возвращает список строк, каждая строка — список слов, отсортированных по x.
    """
    if not words:
        return []

    # Сортируем сначала по y, потом по x
    sorted_words = sorted(words, key=lambda w: (w["top"], w["x0"]))

    lines = []
    current_line = [sorted_words[0]]
    current_y = sorted_words[0]["top"]

    for word in sorted_words[1:]:
        if abs(word["top"] - current_y) <= y_tolerance:
            current_line.append(word)
        else:
            lines.append(sorted(current_line, key=lambda w: w["x0"]))
            current_line = [word]
            current_y = word["top"]

    if current_line:
        lines.append(sorted(current_line, key=lambda w: w["x0"]))

    return lines


def check_pdf_hanging_prepositions(filepath: str) -> dict:
    """
    Проверяет висячие предлоги в PDF.
    Возвращает: {page_num: [{"line": str, "word": str, "context": str}]}
    """
    results = defaultdict(list)

    with pdfplumber.open(filepath) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            page_num = page_idx + 1

            words = page.extract_words(
                x_tolerance=3,
                y_tolerance=LINE_Y_TOLERANCE,
                keep_blank_chars=False,
            )

            if not words:
                continue

            lines = group_words_into_lines(words, y_tolerance=LINE_Y_TOLERANCE)

            for line_words in lines:
                if not line_words:
                    continue

                line_text = " ".join(w["text"] for w in line_words)

                # Пропускаем очень короткие строки
                if len(line_text.strip()) < MIN_LINE_CHARS:
                    continue

                last_word = line_words[-1]["text"].strip(".,!?:;()[]")

                if HANGING_RE.match(last_word):
                    # Контекст: предыдущее слово + висячий предлог
                    if len(line_words) >= 2:
                        prev_word = line_words[-2]["text"]
                        context = f"…{prev_word} {last_word}"
                    else:
                        context = last_word

                    results[page_num].append({
                        "line": line_text[:80],
                        "word": last_word,
                        "context": context,
                    })

    return dict(results)


# ─────────────────────────────────────────────
# ФОРМАТИРОВАНИЕ ОТЧЁТА
# ─────────────────────────────────────────────

def format_pdf_report(results: dict, total_pages: int) -> str:
    lines = []
    lines.append("📄 Проверка висячих предлогов (PDF)")
    lines.append(f"{total_pages} слайдов  ·  найдено: {sum(len(v) for v in results.values())}")

    if not results:
        lines.append("\n✅ Висячих предлогов не найдено")
        return "\n".join(lines)

    lines.append("\nГде встречается — исправьте неразрывным пробелом (Ctrl+Shift+Пробел перед предлогом):")

    for page_num in sorted(results.keys()):
        issues = results[page_num]
        if not issues:
            continue

        lines.append(f"\n*Слайд {page_num}*")
        lines.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

        seen = set()
        for issue in issues:
            key = issue["context"]
            if key in seen:
                continue
            seen.add(key)
            word = issue["word"]
            line_preview = issue["line"][:60].replace("*", "\\*").replace("_", "\\_")
            lines.append(f"⚠️ Висячий предлог «{word}»")
            lines.append(f"   ↳ _{line_preview}_")
            lines.append("")

    return "\n".join(lines)


def split_message(text: str, limit: int = 4000) -> list:
    parts, current, cur_len = [], [], 0
    for line in text.split("\n"):
        if cur_len + len(line) + 1 > limit and current:
            parts.append("\n".join(current))
            current, cur_len = [], 0
        current.append(line)
        cur_len += len(line) + 1
    if current:
        parts.append("\n".join(current))
    return parts
