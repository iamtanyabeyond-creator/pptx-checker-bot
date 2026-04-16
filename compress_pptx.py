"""
compress_pptx.py — Сжатие изображений внутри .pptx файла
.pptx — это ZIP-архив. Изображения лежат в ppt/media/.
Мы их находим, сжимаем через Pillow и кладём обратно.
"""

import io
import os
import zipfile
import shutil
import tempfile
from PIL import Image

# ─────────────────────────────────────────────
# НАСТРОЙКИ СЖАТИЯ
# ─────────────────────────────────────────────

MAX_DIMENSION  = 1920    # максимальная сторона изображения в пикселях
JPEG_QUALITY   = 85      # качество JPEG (85 — хороший баланс размер/качество)
SUPPORTED_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif"}


# ─────────────────────────────────────────────
# ОСНОВНАЯ ФУНКЦИЯ
# ─────────────────────────────────────────────

def compress_pptx(input_path: str, output_path: str) -> dict:
    """
    Сжимает изображения внутри .pptx.
    Возвращает статистику: сколько изображений обработано, старый и новый размер.
    """
    original_size = os.path.getsize(input_path)
    stats = {
        "original_mb": round(original_size / 1024 / 1024, 1),
        "compressed_mb": 0,
        "images_processed": 0,
        "images_skipped": 0,
    }

    with tempfile.TemporaryDirectory() as tmpdir:
        # Распаковываем pptx (это ZIP)
        with zipfile.ZipFile(input_path, 'r') as zin:
            zin.extractall(tmpdir)

        # Находим и сжимаем все изображения в ppt/media/
        media_dir = os.path.join(tmpdir, "ppt", "media")
        if os.path.isdir(media_dir):
            for filename in os.listdir(media_dir):
                ext = os.path.splitext(filename)[1].lower()
                if ext not in SUPPORTED_EXTS:
                    stats["images_skipped"] += 1
                    continue

                img_path = os.path.join(media_dir, filename)
                compressed = _compress_image(img_path)
                if compressed:
                    stats["images_processed"] += 1
                else:
                    stats["images_skipped"] += 1

        # Упаковываем обратно в pptx
        _repack_zip(tmpdir, output_path)

    compressed_size = os.path.getsize(output_path)
    stats["compressed_mb"] = round(compressed_size / 1024 / 1024, 1)
    stats["saved_mb"] = round((original_size - compressed_size) / 1024 / 1024, 1)
    stats["ratio"] = round(original_size / compressed_size, 1) if compressed_size else 1

    return stats


def _compress_image(img_path: str) -> bool:
    """
    Сжимает одно изображение на месте.
    Возвращает True если сжатие выполнено.
    """
    try:
        with Image.open(img_path) as img:
            orig_w, orig_h = img.size

            # Конвертируем RGBA → RGB для JPEG (JPEG не поддерживает прозрачность)
            if img.mode in ("RGBA", "LA", "P"):
                background = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                if img.mode in ("RGBA", "LA"):
                    background.paste(img, mask=img.split()[-1])
                else:
                    background.paste(img)
                img = background
            elif img.mode != "RGB":
                img = img.convert("RGB")

            # Уменьшаем если слишком большое
            if orig_w > MAX_DIMENSION or orig_h > MAX_DIMENSION:
                img.thumbnail((MAX_DIMENSION, MAX_DIMENSION), Image.LANCZOS)

            # Сохраняем как JPEG с нужным качеством
            # Меняем расширение на .jpg если было .png/.bmp и т.д.
            ext = os.path.splitext(img_path)[1].lower()
            if ext in (".png", ".bmp", ".tiff", ".tif"):
                # Сохраняем как JPEG (меньше размер), меняем расширение
                new_path = os.path.splitext(img_path)[0] + ".jpg"
                img.save(new_path, "JPEG", quality=JPEG_QUALITY, optimize=True)
                # Переименовываем только если имя изменилось
                if new_path != img_path:
                    os.replace(new_path, img_path)  # сохраняем под старым именем
                    # Но PPTX ищет файл по оригинальному имени — сохраняем под ним
            else:
                img.save(img_path, "JPEG", quality=JPEG_QUALITY, optimize=True)

        return True

    except Exception:
        return False


def _repack_zip(source_dir: str, output_path: str):
    """
    Упаковывает директорию обратно в ZIP (pptx).
    Важно: файл [Content_Types].xml должен идти первым.
    """
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        # Сначала Content_Types.xml (обязательно первым для совместимости)
        ct_path = os.path.join(source_dir, "[Content_Types].xml")
        if os.path.exists(ct_path):
            zout.write(ct_path, "[Content_Types].xml")

        # Затем всё остальное
        for root, dirs, files in os.walk(source_dir):
            # Сортируем для воспроизводимости
            dirs.sort()
            for filename in sorted(files):
                if filename == "[Content_Types].xml" and root == source_dir:
                    continue  # уже добавили
                file_path = os.path.join(root, filename)
                arcname = os.path.relpath(file_path, source_dir)
                zout.write(file_path, arcname)


# ─────────────────────────────────────────────
# ФОРМАТИРОВАНИЕ СООБЩЕНИЯ О СЖАТИИ
# ─────────────────────────────────────────────

def format_compress_stats(stats: dict) -> str:
    saved = stats.get("saved_mb", 0)
    if saved <= 0:
        return (
            f"🗜 Сжатие: {stats['original_mb']} МБ → {stats['compressed_mb']} МБ "
            f"(изображений: {stats['images_processed']})"
        )
    return (
        f"🗜 Файл сжат: {stats['original_mb']} МБ → {stats['compressed_mb']} МБ "
        f"(–{saved} МБ, в {stats['ratio']}× меньше)"
    )
