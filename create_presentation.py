from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

prs = Presentation()

def add_slide(title_text, items, title_color=RGBColor(0x1B, 0x3A, 0x5C)):
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    title = slide.shapes.title
    title.text = title_text
    title.text_frame.paragraphs[0].font.size = Pt(28)
    title.text_frame.paragraphs[0].font.bold = True
    title.text_frame.paragraphs[0].font.color.rgb = title_color

    body = slide.placeholders[1]
    tf = body.text_frame
    tf.clear()
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.space_after = Pt(8)
        if item.startswith("🔹") or item.startswith("✔") or item.startswith("👉"):
            p.font.size = Pt(18); p.font.bold = True
        elif item.startswith("•") or item.startswith("-"):
            p.font.size = Pt(16)
        elif item.startswith("> "):
            p.text = item[2:]
            p.font.size = Pt(14); p.font.italic = True
            p.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
        else:
            p.font.size = Pt(18)
    return slide

# --- Слайд 1. Титульный ---
slide1 = prs.slides.add_slide(prs.slide_layouts[0])
title1 = slide1.shapes.title
subtitle1 = slide1.placeholders[1]
title1.text = "Интеграция сервисов Mail\nв VK Мессенджер"
subtitle1.text = "Потенциал аудитории, прогноз выручки и MVP\nАнализ на основе данных VK, март 2026"
title1.text_frame.paragraphs[0].font.size = Pt(36)
title1.text_frame.paragraphs[0].font.bold = True
subtitle1.text_frame.paragraphs[0].font.size = Pt(20)

# --- Слайд 2. Какие сервисы Mail запускаем первыми ---
add_slide("Какие сервисы Mail запускаем первыми", [
    "🔹 6 продуктов: Почта, Облако, Задачи, Календарь, Заметки, VK Звонки",
    "✔ Первый эшелон (минимальный порог входа):",
    "   • Задачи — виджет в чате, создание из сообщения",
    "   • Календарь — встречи прямо из переписки",
    "   • Заметки — аналог «Избранного» Telegram",
    "   • VK Звонки — расширение текущих звонков",
    "✔ Второй эшелон: Облако (выгодные тарифы, но сложнее внедрить)",
    "✔ Третий эшелон: Почта (слишком громоздкий функционал)",
    "> Почему: привычные паттерны, быстрое влияние на retention и DAU"
])

# --- Слайд 3. Потенциальная аудитория и выручка ---
add_slide("Потенциальная аудитория и годовая выручка", [
    "🔹 Исходные данные:",
    "   • MAU VK MAX = 107 млн",
    "   • ARPU продукта ≈ 570 ₽/год",
    "",
    "🔹 Расчёт по продуктам:",
    "   • Задачи     (20%): 21,4 млн чел. → 12,2 млрд ₽/год",
    "   • Календарь  (25%): 26,75 млн чел. → 15,2 млрд ₽/год",
    "   • Заметки    (20%): 21,4 млн чел. → 12,2 млрд ₽/год",
    "   • VK Звонки  (50%): 53,5 млн чел. → 30,5 млрд ₽/год",
    "   • Облако     (30%): 32,1 млн чел. → 18,3 млрд ₽/год",
    "",
    "👉 Суммарный потенциал – более 88 млрд ₽/год",
    "> Первыми внедряем Задачи + Заметки (низкая сложность, быстрый эффект)"
])

# --- Слайд 4. Конкурентный ландшафт ---
add_slide("Конкурентный ландшафт (Россия)", [
    "🔹 Ключевые конкуренты: Telegram, ТамТам, TenChat",
    "",
    "✔ Базовые функции (чаты, звонки, боты, группы) – у всех ✅",
    "✔ Кошелёк / переводы – есть у MAX и Telegram, у остальных ❌",
    "✔ Избранное (заметки) – у MAX, Telegram, TenChat (ТамТам ❌)",
    "✔ Вложенные каналы – только у MAX и Telegram",
    "",
    "👉 Уникальное преимущество MAX:",
    "   • Глубокая интеграция Задач, Календаря, Облака прямо в мессенджер",
    "   • Этого нет ни у одного конкурента"
])

# --- Слайд 5. Ключевые гипотезы роста и методы проверки ---
add_slide("Ключевые гипотезы роста и методы проверки", [
    "🔹 Высокий приоритет:",
    "   • Заметки + Задачи → +5% retention, +20% DAU",
    "   • Календарь → +10 пунктов scroll depth",
    "   • VK Звонки → +5% ежедневный retention",
    "   ✔ Проверка: A/B-тест 2 недели (50/50%)",
    "",
    "🔹 Discovery до разработки:",
    "   • Кастдев-интервью 10–15 пользователей",
    "   • Всплывающий опрос для 1000+ активных",
    "   • Аналитика точек выхода из приложения"
])

# --- Слайд 6. MVP и что появится первым ---
add_slide("MVP: Заметки + Задачи как единый виджет", [
    "👉 Первый релиз — виджет внутри чата:",
    "   • Создание заметки/задачи из сообщения (одна кнопка)",
    "   • Дедлайн и напоминание",
    "   • Вкладка «Все задачи»",
    "   • Отправка задачи другому пользователю в чате",
    "",
    "✔ Почему они:",
    "   • Минимальная сложность разработки",
    "   • Нет отдельной авторизации",
    "   • Закрывают главную боль: закреплённые сообщения вместо заметок",
    "   • Быстро измеряемый эффект на retention",
    "",
    "После проверки гипотез → Добавляем Календарь и VK Звонки"
])

# Сохраняем файл
prs.save("Интеграция_Mail_в_VK.pptx")
print("Готово: Интеграция_Mail_в_VK.pptx")
