"""
ai_editor.py — Редакторская проверка текста презентации через GigaChat
"""

import os
from gigachat import GigaChat
from gigachat.models import Chat, Messages, MessagesRole


EDITOR_SYSTEM_PROMPT = """Ты — опытный редактор презентаций.
Твоя задача — разобрать текст со слайдов и дать конкретные рекомендации.

Что нужно найти:
- Канцелярит: «в целях осуществления», «является одним из», «на сегодняшний день»
- Слабые формулировки: «очень важно», «достаточно хорошие», «в целом неплохо»
- Слишком длинные и сложные предложения для презентации
- Повторы слов и мыслей между слайдами
- Заголовок не отражает суть слайда
- Пассивный залог там, где лучше активный

Формат — строго по слайдам:
Слайд N: [конкретная проблема] → [конкретное исправление]

Если на слайде всё хорошо — не упоминай его.
Не пиши общих советов — только конкретные замены.
Отвечай на русском языке."""


def build_prompt(slides_text: str) -> str:
    return (
        f"Текст из презентации по слайдам:\n\n{slides_text}\n\n"
        "Проведи редакторский разбор. Найди речевые ошибки, канцелярит, слабые формулировки. "
        "Дай конкретные рекомендации по каждому проблемному слайду."
    )


def run_ai_editor(slides_text: str) -> str:
    credentials = os.getenv("GIGACHAT_CREDENTIALS")
    if not credentials:
        return "❌ GigaChat не настроен. Проверь GIGACHAT_CREDENTIALS в файле .env"

    try:
        with GigaChat(credentials=credentials, verify_ssl_certs=False) as giga:
            response = giga.chat(
                Chat(
                    messages=[
                        Messages(role=MessagesRole.SYSTEM, content=EDITOR_SYSTEM_PROMPT),
                        Messages(role=MessagesRole.USER, content=build_prompt(slides_text)),
                    ]
                )
            )
            return response.choices[0].message.content
    except Exception as e:
        return f"❌ Ошибка при обращении к GigaChat: {str(e)}"


def format_ai_report(ai_response: str) -> str:
    return (
        "✍️ *Редакторский разбор от ИИ*\n"
        "━━━━━━━━━━━━━━━━━━\n\n"
        + ai_response
    )
