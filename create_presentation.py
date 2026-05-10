from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import copy

# Цветовая палитра
DARK_BLUE = RGBColor(0x1B, 0x3A, 0x5C)      # Тёмно-синий (заголовки, акценты)
ACCENT_BLUE = RGBColor(0x2E, 0x86, 0xC1)    # Голубой
LIGHT_BG = RGBColor(0xF0, 0xF4, 0xFA)       # Светло-голубой фон
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)

prs = Presentation()
prs.slide_width = Inches(13.333)   # 16:9
prs.slide_height = Inches(7.5)

# Функция установки фона слайда
def set_background(slide, color):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color

# Функция добавления декоративной полосы сверху
def add_top_bar(slide, color, height=Inches(0.08)):
    """Тонкая цветная полоса в верхней части слайда"""
    left = Inches(0)
    top = Inches(0)
    width = prs.slide_width
    shape = slide.shapes.add_shape(
        1, left, top, width, height   # 1 = MSO_SHAPE.RECTANGLE
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()

# Функция создания контентного слайда с иконками
def add_content_slide(title_text, items, title_icon=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Пустой макет
    set_background(slide, LIGHT_BG)
    add_top_bar(slide, ACCENT_BLUE)

    # Заголовок
    left = Inches(0.8)
    top = Inches(0.5)
    width = Inches(11.5)
    height = Inches(1.0)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = f"{title_icon}  {title_text}" if title_icon else title_text
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = DARK_BLUE

    # Контент (список)
    left = Inches(1.0)
    top = Inches(1.7)
    width = Inches(11.2)
    height = Inches(5.0)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        # Каждый пункт начинается с иконки, если не указано своё начало
        if not item.startswith("📌") and not item.startswith("🔹") and not item.startswith("✔") and not item.startswith("👉") and not item.startswith(">>"):
            # Добавляем маркер-буллит
            p.text = f"●  {item}"
        else:
            p.text = item
        p.space_after = Pt(10)
        p.font.size = Pt(18)
        p.font.color.rgb = DARK_GRAY
        # Подпункты выделяем отступом и меньшим шрифтом
        if item.startswith("   "):
            p.level = 1
            p.font.size = Pt(16)
    return slide

# ==================== Титульный слайд ====================
slide1 = prs.slides.add_slide(prs.slide_layouts[6])
set_background(slide1, DARK_BLUE)

# Большая декоративная полоса внизу заголовка
shape = slide1.shapes.add_shape(
    1, Inches(0), Inches(3.2), prs.slide_width, Inches(0.1)
)
shape.fill.solid()
shape.fill.fore_color.rgb = ACCENT_BLUE
shape.line.fill.background()

# Заголовок
txBox = slide1.shapes.add_textbox(Inches(1.5), Inches(1.5), Inches(10.3), Inches(1.8))
tf = txBox.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "Интеграция сервисов Mail в VK Мессенджер"
p.font.size = Pt(40)
p.font.bold = True
p.font.color.rgb = WHITE
p.alignment = PP_ALIGN.LEFT

# Подзаголовок
txBox2 = slide1.shapes.add_textbox(Inches(1.5), Inches(3.6), Inches(10.3), Inches(1.2))
tf2 = txBox2.text_frame
p2 = tf2.paragraphs[0]
p2.text = "Потенциал аудитории, прогноз выручки и MVP\nАнализ на основе данных VK, март 2026 г."
p2.font.size = Pt(24)
p2.font.color.rgb = RGBColor(0xBB, 0xCC, 0xDD)
p2.alignment = PP_ALIGN.LEFT

# Иконка презентации (эмодзи)
txBox3 = slide1.shapes.add_textbox(Inches(11.5), Inches(0.3), Inches(1.5), Inches(1.0))
tf3 = txBox3.text_frame
p3 = tf3.paragraphs[0]
p3.text = "📊"
p3.font.size = Pt(36)

# ==================== Слайд 2. Приоритезация сервисов ====================
add_content_slide(
    "Какие сервисы Mail запускаем первыми",
    [
        "📌 6 продуктов: Почта, Облако, Задачи, Календарь, Заметки, VK Звонки",
        "",
        "✔ Первый эшелон (минимальный порог входа):",
        "   ● Задачи — виджет в чате, создание из сообщения",
        "   ● Календарь — встречи прямо из переписки",
        "   ● Заметки — аналог «Избранного» Telegram",
        "   ● VK Звонки — расширение текущих звонков",
        "",
        "✔ Второй эшелон: Облако (выгодные тарифы, но сложнее)",
        "✔ Третий эшелон: Почта (громоздкий функционал)",
        "",
        ">> Почему: привычные паттерны, быстрое влияние на retention и DAU"
    ],
    title_icon="🚀"
)

# ==================== Слайд 3. Аудитория и выручка ====================
add_content_slide(
    "Потенциальная аудитория и годовая выручка",
    [
        "📊 Исходные данные: MAU VK MAX = 107 млн, ARPU ≈ 570 ₽/год",
        "",
        "🔹 Задачи      (20%): 21,4 млн чел. → 12,2 млрд ₽",
        "🔹 Календарь   (25%): 26,75 млн чел. → 15,2 млрд ₽",
        "🔹 Заметки     (20%): 21,4 млн чел. → 12,2 млрд ₽",
        "🔹 VK Звонки   (50%): 53,5 млн чел. → 30,5 млрд ₽",
        "🔹 Облако      (30%): 32,1 млн чел. → 18,3 млрд ₽",
        "",
        "💰 Суммарный потенциал – более 88 млрд ₽/год",
        ">> Первыми внедряем Задачи + Заметки (быстрый эффект, низкая сложность)"
    ],
    title_icon="📈"
)

# ==================== Слайд 4. Конкуренты ====================
add_content_slide(
    "Конкурентный ландшафт (Россия)",
    [
        "🔍 Ключевые конкуренты: Telegram, ТамТам, TenChat",
        "",
        "✅ Базовые функции (чаты, звонки, боты, группы) – у всех",
        "✅ Кошелёк / переводы – у MAX и Telegram (остальные ❌)",
        "✅ Избранное (заметки) – у MAX, Telegram, TenChat",
        "✅ Вложенные каналы – только у MAX и Telegram",
        "",
        "💎 Уникальное преимущество MAX:",
        "   ● Интеграция Задач, Календаря, Облака прямо в мессенджер",
        "   ● Ни у кого из конкурентов такого нет"
    ],
    title_icon="🔎"
)

# ==================== Слайд 5. Гипотезы ====================
add_content_slide(
    "Ключевые гипотезы роста и методы проверки",
    [
        "🚀 Высокий приоритет:",
        "   ● Заметки + Задачи → +5% retention, +20% DAU",
        "   ● Календарь → +10 пунктов scroll depth",
        "   ● VK Звонки → +5% ежедневный retention",
        "   ● A/B-тест 2 недели (50/50%)",
        "",
        "🔬 Discovery до разработки:",
        "   ● 10–15 кастдев-интервью",
        "   ● Опрос 1000+ активных пользователей",
        "   ● Аналитика точек выхода",
        "",
        "💡 Дополнительно: упрощение API и ботов, приватные каналы, кошелёк"
    ],
    title_icon="💡"
)

# ==================== Слайд 6. MVP ====================
add_content_slide(
    "MVP: Заметки + Задачи как единый виджет",
    [
        "🎯 Первый релиз — виджет внутри чата:",
        "   ● Создание заметки/задачи из сообщения (одна кнопка)",
        "   ● Дедлайн и напоминание",
        "   ● Вкладка «Все задачи»",
        "   ● Шеринг задачи в чате",
        "",
        "⚡ Почему они:",
        "   ● Минимальная сложность разработки, без доп. авторизации",
        "   ● Закрывают главную боль: закреплённые сообщения вместо заметок",
        "   ● Быстрый измеряемый эффект на retention",
        "",
        "👉 После проверки гипотез → Календарь и VK Звонки"
    ],
    title_icon="🎯"
)

# Сохранение
prs.save("Интеграция_Mail_в_VK.pptx")
print("✅ Готова стильная презентация: Интеграция_Mail_в_VK.pptx")
