# -*- coding: utf-8 -*-
"""НТС, работа 3: мультигипотезы для модели оценки рисков внедрения ИИ,
количественные данные Eurostat, код проверки, скриншоты — презентация."""
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

HERE = Path(__file__).parent
OUT = HERE.parent / "НТС3_Мультигипотезы_Рослов.pptx"
FONT = "Bahnschrift Light"
DARK = RGBColor(0x22, 0x2B, 0x36)
BLUE = RGBColor(0x2E, 0x6F, 0xB7)
ORANGE = RGBColor(0xE0, 0x8A, 0x1E)
GREEN = RGBColor(0x3E, 0x9E, 0x52)
RED = RGBColor(0xB0, 0x3A, 0x3A)
W, H = Inches(13.333), Inches(7.5)

prs = Presentation()
prs.slide_width, prs.slide_height = W, H
BLANK = prs.slide_layouts[6]


def text(slide, x, y, w, h, paras, size=16, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP):
    """paras: список строк или (строка, {bold, color, size})."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for i, p in enumerate(paras):
        t, o = (p, {}) if isinstance(p, str) else p
        par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        par.alignment = o.get("align", align)
        par.space_after = Pt(o.get("after", 4))
        for j, chunk in enumerate(t.split("**")):
            if not chunk:
                continue
            r = par.add_run()
            r.text = chunk
            r.font.name = FONT
            r.font.size = Pt(o.get("size", size))
            r.font.bold = o.get("bold", False) or j % 2 == 1
            r.font.color.rgb = o.get("color", DARK)
    return tb


def title(slide, t, sub=None):
    text(slide, Inches(0.4), Inches(0.25), W - Inches(0.8), Inches(0.8),
         [(t.upper(), {"size": 30, "align": PP_ALIGN.CENTER})])
    if sub:
        text(slide, Inches(0.4), Inches(0.95), W - Inches(0.8), Inches(0.5),
             [(sub, {"size": 15, "align": PP_ALIGN.CENTER,
                     "color": RGBColor(0x55, 0x5F, 0x6B)})])


def picture(slide, path, x, y, w=None, h=None):
    """Картинка, вписанная в прямоугольник w×h с сохранением пропорций."""
    iw, ih = Image.open(HERE / path).size
    if w and h:
        k = min(w / iw, h / ih)
        pw, ph = int(iw * k), int(ih * k)
        x, y = x + (w - pw) // 2, y + (h - ph) // 2
        return slide.shapes.add_picture(str(HERE / path), x, y, pw, ph)
    return slide.shapes.add_picture(str(HERE / path), x, y, w, h)


def box(slide, x, y, w, h, t, color, size=15):
    sh = slide.shapes.add_shape(1, x, y, w, h)
    sh.fill.solid()
    sh.fill.fore_color.rgb = color
    sh.line.fill.background()
    tf = sh.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.text = t
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        for r in p.runs:
            r.font.name, r.font.size = FONT, Pt(size)
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    return sh


def table(slide, x, y, w, rows, widths, size=12, h_row=0.4):
    shp = slide.shapes.add_table(len(rows), len(rows[0]), x, y, w,
                                 Inches(h_row * len(rows)))
    tb = shp.table
    for j, cw in enumerate(widths):
        tb.columns[j].width = Inches(cw)
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = tb.cell(i, j)
            c.text = ""
            p = c.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = str(val)
            r.font.name, r.font.size = FONT, Pt(size)
            r.font.bold = i == 0
            r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF) if i == 0 else DARK
            c.fill.solid()
            c.fill.fore_color.rgb = (BLUE if i == 0 else
                                     RGBColor(0xF3, 0xF6, 0xFA) if i % 2 else
                                     RGBColor(0xFF, 0xFF, 0xFF))
            c.margin_top = c.margin_bottom = Inches(0.04)
    return tb


def slide():
    return prs.slides.add_slide(BLANK)


X0, CW = Inches(0.6), W - Inches(1.2)

# 1 --- титул
s = slide()
text(s, X0, Inches(2.0), CW, Inches(2),
     [("МУЛЬТИГИПОТЕЗЫ ДЛЯ МОДЕЛИ ОЦЕНКИ РИСКОВ ВНЕДРЕНИЯ ТЕХНОЛОГИЙ "
       "ИСКУССТВЕННОГО ИНТЕЛЛЕКТА В IT-КОМПАНИИ НА ОСНОВЕ НЕЧЁТКОЙ ЛОГИКИ",
       {"size": 30, "align": PP_ALIGN.CENTER})])
text(s, X0, Inches(4.3), CW, Inches(1.2),
     [("Феномен → гипотезы → количественные данные → код → подтверждение",
       {"size": 18, "align": PP_ALIGN.CENTER, "color": BLUE})])
text(s, X0, Inches(5.8), CW, Inches(1),
     [("Выполнил: Рослов Д. С., гр. М558М", {"size": 16,
                                             "align": PP_ALIGN.RIGHT}),
      ("Санкт-Петербург, 2026", {"size": 16, "align": PP_ALIGN.RIGHT})])

# 2 --- феномен
s = slide()
title(s, "Феномен")
text(s, X0, Inches(1.0), CW, Inches(1.4), [
    ("IT-компании сами создают технологии ИИ, но **каждая третья** из них "
     "ИИ не использует, а **каждая девятая** рассматривала внедрение и "
     "**отказалась**. При этом в IT-секторе ИИ используют в 3,3 раза чаще, "
     "чем в экономике в целом, и доля растёт на 24 % в год.",
     {"size": 18, "align": PP_ALIGN.CENTER})])
picture(s, "fig_phenomenon.png", X0, Inches(2.45), CW, Inches(4.7))
text(s, X0, Inches(7.05), CW, Inches(0.4), [
    ("Источник: Eurostat, isoc_eb_ain2, NACE J62–J63, ЕС-27, 2021–2025",
     {"size": 11, "color": RGBColor(0x77, 0x77, 0x77)})])

# 3 --- феномен: границы
s = slide()
title(s, "Феномен", "границы и характеристика")
y = Inches(1.9)
for i, (t, c) in enumerate([("Решение\nо запуске ИИ", BLUE),
                            ("Пилот", BLUE),
                            ("Барьеры:\nданные, право,\nкомпетенции", ORANGE),
                            ("Отказ\n(11 % IT-компаний)", RED),
                            ("Внедрение\n(66 %)", GREEN)]):
    box(s, Inches(0.6 + i * 2.5), y, Inches(2.2), Inches(1.3), t, c, 16)
text(s, X0, Inches(3.4), CW, Inches(0.5), [
    ("тело феномена — отрезок от решения о запуске до отказа: именно здесь "
     "работает модель оценки рисков", {"size": 15, "align": PP_ALIGN.CENTER,
                                       "color": RGBColor(0x55, 0x5F, 0x6B)})])
text(s, X0, Inches(4.1), CW, Inches(3), [
    "**Границы феномена:**",
    "• отрасль — IT-компании (NACE J62–J63: разработка ПО, IT-консалтинг, "
    "информационные услуги);",
    "• период — 2021–2025 гг., когда генеративный ИИ стал массовым;",
    "• объект наблюдения — предприятия от 10 занятых, которые рассматривали "
    "ИИ, но не внедрили его;",
    "• размер компании усиливает феномен: ИИ используют 17 % малых "
    "(10–49 чел.) и 55 % крупных (250+) предприятий.",
    "**Что тут происходит?** Почему отказываются те, у кого есть и "
    "технологии, и специалисты?"], size=17)

# 4 --- гипотеза 1
s = slide()
title(s, "Гипотеза 1", "это всё технологии…")
text(s, X0, Inches(1.5), CW, Inches(5.6), [
    "**1.1 ИИ бесполезен** = доказательство => модели галлюцинируют, "
    "результат нестабилен, выгода для бизнеса не видна, «хайп» не "
    "подтверждается практикой.",
    "**1.2 ИИ несовместим с ИС** = доказательство => унаследованные системы, "
    "нет API и MLOps-инфраструктуры, интеграция дороже самой модели.",
    "**1.3 Нет данных** = доказательство => данные разрознены, не размечены, "
    "непредставительны, обучать и проверять модель не на чем.",
    "",
    ("Проверка: если Г1 верна, технологические барьеры должны быть самыми "
     "частыми причинами отказа и чаще встречаться там, где ИИ используют "
     "меньше.", {"color": BLUE}),
    ("Связь с моделью: x₁ — качество данных, x₂ — новизна и сложность "
     "ИИ-решения.", {"color": BLUE})], size=18)

# 5 --- гипотеза 2
s = slide()
title(s, "Гипотеза 2", "это всё люди и среда…")
text(s, X0, Inches(1.5), CW, Inches(5.6), [
    "**2.1 Компетенции** = характеристика => нет специалистов на стыке ИИ, "
    "данных и предметной области; дефицит кадров на рынке.",
    "**2.2 Стоимость** = характеристика => вычисления, лицензии, специалисты "
    "дорогие, эффект неочевиден — бюджет не выделяют.",
    "**2.3 Правовая неясность** = характеристика => неясно, кто отвечает за "
    "ошибку ИИ; регулирование (AI Act, 152-ФЗ) меняется быстрее проектов.",
    "**2.4 Защита персональных данных** = характеристика => риск утечки "
    "данных клиентов через модели и облачные сервисы.",
    "**2.5 Этика** = характеристика => опасения предвзятости, репутационные "
    "риски.",
    ("Связь с моделью: x₃ — компетентность команды, x₄ — регуляторная и "
     "этическая чувствительность, x₆ — организационная готовность.",
     {"color": BLUE})], size=18)

# 6 --- гипотеза 3
s = slide()
title(s, "Гипотеза 3", "профиль рисков смещается от технологий к регуляторике")
text(s, X0, Inches(1.5), CW, Inches(5.6), [
    "**3.1** С развитием облачных ИИ-сервисов технологические барьеры "
    "(несовместимость, данные) ослабевают.",
    "**3.2** Одновременно растут регуляторные барьеры: правовая неясность "
    "и защита персональных данных (AI Act принят в 2024 г.).",
    "**3.3** Модель, настроенная на «технические» риски 2023 года, "
    "устаревает — веса факторов нужно регулярно пересматривать.",
    "**3.4** Следствие для СППР: фактору x₄ (регуляторная "
    "чувствительность) — больший вес в базе правил; калибровка функций "
    "принадлежности — ежегодно.",
    "",
    ("Проверка: сравнение долей отказавшихся по каждой причине в 2023 и "
     "2025 гг. по странам (парный критерий Уилкоксона).",
     {"color": BLUE})], size=18)

# 7 --- гипотеза 4
s = slide()
title(s, "Гипотеза 4",
      "если нельзя понять, какой барьер главный, — оцени все сразу")
for i, (t, c) in enumerate([("x₁ данные", BLUE), ("x₂ технология", BLUE),
                            ("x₃ компетенции", BLUE), ("x₄ регуляторика", BLUE),
                            ("x₅ безопасность", BLUE), ("x₆ организация", BLUE)]):
    box(s, Inches(0.6 + i * 2.05), Inches(1.6), Inches(1.85), Inches(0.8), t,
        c, 15)
box(s, Inches(3.9), Inches(2.9), Inches(5.5), Inches(1.0),
    "Нечёткий вывод Мамдани → интегральный риск R ∈ [0; 100]", GREEN, 17)
text(s, X0, Inches(4.3), CW, Inches(3), [
    "**4.1** Ни один барьер в отдельности не объясняет, внедрит ли "
    "компания ИИ: связь каждого барьера с использованием ИИ слабая.",
    "**4.2** Барьеры слабо связаны между собой — это разные, независимые "
    "факторы, их нельзя свести к одному показателю.",
    "**4.3** Значит, решение о запуске ИИ-проекта нужно принимать по "
    "интегральной оценке всех факторов одновременно — это и делает СППР "
    "на нечёткой логике.",
    ("Проверка: корреляции Спирмена «барьер — использование ИИ» и «барьер — "
     "барьер» по странам.", {"color": BLUE})], size=17)

# 8 --- количественные данные
s = slide()
title(s, "Количественные данные для проверки гипотез")
table(s, X0, Inches(1.2), CW, [
    ["Показатель Eurostat", "Что измеряет", "Ед.", "Годы", "Гипотеза",
     "Фактор модели"],
    ["E_AI_TANY", "Предприятия, использующие ИИ", "%", "2021–2025",
     "Феномен, Г4", "R (исход)"],
    ["E_AI_EC", "Рассматривали ИИ, но не используют", "%", "2023–2025",
     "Феномен", "—"],
    ["E_AI_BNU / BINC / BDDT", "Отказ: бесполезность / несовместимость / "
     "данные", "%", "2023–2025", "Г1, Г3", "x₁, x₂"],
    ["E_AI_BLE", "Отказ: нет компетенций", "%", "2023–2025", "Г2", "x₃"],
    ["E_AI_BCST", "Отказ: высокая стоимость", "%", "2023–2025", "Г2", "x₆"],
    ["E_AI_BLEG / BCDP / BEC", "Отказ: право / ПДн / этика", "%",
     "2023–2025", "Г2, Г3", "x₄, x₅"],
], [2.6, 3.9, 0.6, 1.3, 1.6, 2.1], size=13, h_row=0.55)
text(s, X0, Inches(5.3), CW, Inches(2), [
    "Разрезы: IT-сектор (NACE J62–J63) и вся экономика; ЕС-27 и 18–30 "
    "стран; размер предприятия (10–49, 50–249, 250+).",
    "Доли причин отказа рассчитаны от числа предприятий, которые "
    "рассматривали ИИ, но не внедрили (единица PC_ENT_AI_EC).",
    "Всего загружено 1 625 наблюдений из двух наборов данных через "
    "открытый API Eurostat."], size=16)

# 9 --- код: загрузка
s = slide()
title(s, "Код: загрузка данных", "fetch_data.py — API Eurostat, формат JSON-stat → CSV")
picture(s, "code_fetch.png", X0, Inches(1.5), CW, Inches(5.8))

# 10 --- код: Г1, Г2
s = slide()
title(s, "Код и результат: гипотезы 1 и 2",
      "hypotheses.py — критерий Уилкоксона и корреляция Спирмена")
picture(s, "code_g12.png", Inches(0.3), Inches(1.45), Inches(6.2), Inches(5.9))
picture(s, "out_g12.png", Inches(6.7), Inches(1.45), Inches(6.35), Inches(5.9))

# 11 --- результат Г1, Г2
s = slide()
title(s, "Гипотезы 1 и 2: результат")
picture(s, "fig_barriers.png", Inches(0.3), Inches(1.1), Inches(8.2),
        Inches(5.2))
text(s, Inches(8.7), Inches(1.2), Inches(4.3), Inches(6), [
    ("Г1 «технологии» — отвергнута", {"bold": True, "color": RED}),
    "Технологические причины в среднем называют 27,8 % отказавшихся, "
    "«бесполезность» ИИ — лишь 16,6 %.",
    ("Г2 «люди и среда» — подтверждена", {"bold": True, "color": GREEN}),
    "Людские и средовые причины — 39,0 %, разница значима: критерий "
    "Уилкоксона, n = 18, p = 0,0014.",
    "Лидеры: защита ПДн 53,2 %, правовая неясность 52,5 %, компетенции "
    "52,3 %.",
    "Единственный барьер, значимо связанный с уровнем использования ИИ "
    "по странам, — стоимость (rs = −0,54; p = 0,022)."], size=15)
text(s, Inches(0.3), Inches(6.5), Inches(8.2), Inches(0.9), [
    ("Серые столбцы — 2023 г., цветные — 2025 г. Источник: Eurostat "
     "isoc_eb_ain2.", {"size": 11, "color": RGBColor(0x77, 0x77, 0x77)})])

# 12 --- результат Г3
s = slide()
title(s, "Гипотеза 3: результат", "парный критерий Уилкоксона, 2023 → 2025, n = 17 стран")
picture(s, "fig_shift.png", Inches(0.3), Inches(1.4), W - Inches(0.6),
        Inches(3.9))
text(s, X0, Inches(5.4), CW, Inches(2), [
    ("Подтверждена: правовая неясность +14,4 п. п. (p = 0,017), защита ПДн "
     "+12,6 п. п. (p = 0,002) — рост значим.", {"color": GREEN}),
    "Несовместимость и проблемы с данными значимо не изменились "
    "(p = 0,78 и 0,85): технологические барьеры не растут, а "
    "регуляторные — растут.",
    ("Вывод для модели: вес фактора x₄ увеличить, базу правил "
     "перекалибровывать ежегодно.", {"color": BLUE})], size=16)

# 13 --- результат Г4
s = slide()
title(s, "Гипотеза 4: результат")
picture(s, "fig_g4.png", Inches(0.3), Inches(1.0), W - Inches(0.6),
        Inches(4.3))
text(s, X0, Inches(5.4), CW, Inches(2), [
    ("Подтверждена: сильнейший одиночный барьер (стоимость) объясняет лишь "
     "29 % вариации использования ИИ (rs² = 0,29); остальные семь связей "
     "не значимы.", {"color": GREEN}),
    "Барьеры между собой почти независимы: средний rs = 0,09, значимо "
    "связаны лишь 2 пары из 28.",
    ("Вывод: решение нужно принимать по интегральной оценке всех факторов "
     "— нечёткая модель с входами x₁…x₆ обоснована.", {"color": BLUE})],
    size=16)

# 14 --- код Г3, Г4 и вывод программы
s = slide()
title(s, "Код и результат: гипотезы 3 и 4")
picture(s, "code_g34.png", Inches(0.3), Inches(1.0), Inches(6.4), Inches(6.3))
picture(s, "out_g34.png", Inches(6.8), Inches(1.0), Inches(6.25), Inches(3.2))
picture(s, "out_phen.png", Inches(6.8), Inches(4.4), Inches(6.25), Inches(2.9))

# 15 --- итог
s = slide()
title(s, "Гипотеза + статистика + факты = подтверждение")
table(s, X0, Inches(1.2), CW, [
    ["Гипотеза", "Статистика", "Факт", "Вывод"],
    ["Г1 «это всё технологии»", "Доли причин отказа, Спирмен",
     "Технологические причины 27,8 %, связь с использованием не значима",
     "Отвергнута"],
    ["Г2 «это всё люди и среда»", "Уилкоксон, n = 18",
     "39,0 % против 27,8 %, p = 0,0014", "Подтверждена"],
    ["Г3 «сдвиг к регуляторике»", "Уилкоксон парный, n = 17",
     "Право +14,4 п. п. (p = 0,017), ПДн +12,6 п. п. (p = 0,002)",
     "Подтверждена"],
    ["Г4 «оцени все сразу»", "Корреляции Спирмена",
     "max rs² = 0,29; средний rs между барьерами 0,09",
     "Подтверждена"],
], [3.0, 2.8, 4.6, 1.73], size=14, h_row=0.7)
text(s, X0, Inches(5.0), CW, Inches(2.3), [
    "**Следствия для модели СППР:**",
    "• входы x₃ (компетенции) и x₄ (регуляторика и ПДн) — ключевые, им "
    "больший вес в базе продукционных правил;",
    "• в модель добавляется экономический фактор (стоимость) — единственный "
    "значимо связанный с отказом по странам;",
    "• интегральная нечёткая оценка всех факторов обоснована: ни один "
    "фактор в отдельности не решает."], size=16)

# 16 --- релевантность данных
s = slide()
title(s, "Релевантность данных и источники")
text(s, Inches(0.4), Inches(1.1), Inches(6.3), Inches(6.2), [
    "**Почему данным можно доверять:**",
    "• официальная статистика ЕС: ежегодное обследование использования "
    "ИКТ предприятиями по единой методологии во всех странах;",
    "• отраслевой срез совпадает с объектом модели — IT-компании "
    "(NACE J62–J63);",
    "• свежие данные: 2021–2025 гг., обновление Eurostat от 15.06.2026;",
    "• данные получены кодом через открытый API — расчёт воспроизводим.",
    "**Ограничения:**",
    "• это данные ЕС, а не России; аналогичные показатели по России "
    "публикует ВШЭ в сборнике «Индикаторы цифровой экономики»;",
    "• анализ по странам — агрегаты, n = 17–18 стран с полными данными;",
    "• причины отказа указывают только предприятия, рассматривавшие ИИ."],
    size=14)
text(s, Inches(6.9), Inches(1.1), Inches(6.1), Inches(6.2), [
    "**Источники:**",
    "1. Eurostat. Artificial intelligence by NACE Rev. 2 activity "
    "(isoc_eb_ain2). URL: https://ec.europa.eu/eurostat/databrowser/view/"
    "isoc_eb_ain2",
    "2. Eurostat. Artificial intelligence by size class of enterprise "
    "(isoc_eb_ai). URL: https://ec.europa.eu/eurostat/databrowser/view/"
    "isoc_eb_ai",
    "3. Eurostat. Statistics API (JSON-stat 2.0). URL: https://ec.europa.eu/"
    "eurostat/web/user-guides/data-browser/api-data-access",
    "4. НИУ ВШЭ. Статистические сборники. URL: https://www.hse.ru/"
    "primarydata/",
    "5. Google Dataset Search. URL: https://datasetsearch.research.google.com",
    "6. Regulation (EU) 2024/1689 (Artificial Intelligence Act).",
    "Код: fetch_data.py, hypotheses.py, figures.py (Python 3, SciPy, "
    "NumPy, Matplotlib)."], size=13)

prs.save(OUT)
print("saved", OUT)
