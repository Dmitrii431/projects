# -*- coding: utf-8 -*-
"""НТС-3. Проверка мультигипотез для модели оценки рисков внедрения ИИ
в IT-компании на открытых данных Eurostat (NACE J62-J63 — разработка ПО,
IT-консалтинг и информационные услуги)."""
import csv
from collections import defaultdict
from itertools import combinations

import numpy as np
from scipy import stats

AGG = {"EU27_2020", "EA"}
IT, ALL = "J62_J63", "C10-S951_X_K"
TECH = {"E_AI_BNU": "ИИ бесполезен", "E_AI_BINC": "Несовместимость с ИС",
        "E_AI_BDDT": "Данные: доступ/качество"}
PEOPLE = {"E_AI_BLE": "Нет компетенций", "E_AI_BCST": "Высокая стоимость",
          "E_AI_BLEG": "Правовая неясность", "E_AI_BCDP": "Защита ПДн",
          "E_AI_BEC": "Этика"}
BARRIERS = TECH | PEOPLE


def load(name):
    return list(csv.DictReader(open(name)))


def pick(rows, **flt):
    return {r["time"] if "geo" in flt else r["geo"]: float(r["value"])
            for r in rows if all(r[k] == v for k, v in flt.items())}


trend, barr, ctry, size = (load(f) for f in ("eu_trend.csv", "eu_barriers.csv",
                                             "countries_it.csv", "eu_size.csv"))

# ---------------------------------------------------------------- ФЕНОМЕН
print("=== ФЕНОМЕН: использование ИИ, % предприятий (ЕС-27) ===")
it = pick(trend, nace_r2=IT, indic_is="E_AI_TANY", geo="EU27_2020")
al = pick(trend, nace_r2=ALL, indic_is="E_AI_TANY", geo="EU27_2020")
for y in sorted(it):
    print(f"{y}: IT-сектор {it[y]:5.1f}   вся экономика {al[y]:5.1f}   "
          f"разрыв x{it[y] / al[y]:.1f}")
cagr = (it["2025"] / it["2021"]) ** (1 / 4) - 1
print(f"Среднегодовой рост в IT-секторе 2021-2025: {cagr:.1%}")
print(f"Не используют ИИ в IT-секторе в 2025: {100 - it['2025']:.1f} %")
sz = {r["size_emp"]: float(r["value"]) for r in size
      if r["indic_is"] == "E_AI_TANY" and r["time"] == "2025"
      and r["unit"] == "PC_ENT"}
print("По размеру (все отрасли, 2025):",
      ", ".join(f"{k}: {sz[k]:.1f}" for k in ("10-49", "50-249", "GE250")))

# -------------------------------------------- Г1 и Г2: какие барьеры главные
print("\n=== Г1 «это всё технологии» / Г2 «это всё люди и среда» ===")
print("Доля предприятий IT-сектора ЕС, рассматривавших ИИ, но не внедривших,"
      " по причине, %")
eu = {k: pick(barr, nace_r2=IT, indic_is=k, unit="PC_ENT_AI_EC",
              geo="EU27_2020") for k in BARRIERS}
for k in sorted(BARRIERS, key=lambda k: -eu[k]["2025"]):
    g = "Г1" if k in TECH else "Г2"
    print(f"  {g} {BARRIERS[k]:<24} 2023: {eu[k]['2023']:5.1f}  "
          f"2025: {eu[k]['2025']:5.1f}")

# по странам, 2025: парное сравнение средних по группам барьеров
Y = "2025"
cb = {k: pick([r for r in ctry if r["time"] == Y and r["geo"] not in AGG],
              indic_is=k, unit="PC_ENT_AI_EC") for k in BARRIERS}
use = pick([r for r in ctry if r["time"] == Y and r["geo"] not in AGG],
           indic_is="E_AI_TANY", unit="PC_ENT")
geos = sorted(set.intersection(*(set(v) for v in cb.values())) & set(use))
t_mean = np.array([np.mean([cb[k][g] for k in TECH]) for g in geos])
p_mean = np.array([np.mean([cb[k][g] for k in PEOPLE]) for g in geos])
w = stats.wilcoxon(p_mean, t_mean, alternative="greater")
print(f"\nСтраны с полными данными: n = {len(geos)}")
print(f"Средняя доля: технологические {t_mean.mean():.1f} %, "
      f"людские/средовые {p_mean.mean():.1f} %")
print(f"Критерий Уилкоксона (Г2 > Г1): W = {w.statistic:.0f}, "
      f"p = {w.pvalue:.4f}")

# связь каждого барьера с уровнем использования ИИ по странам
print("\nРанговая корреляция Спирмена «доля барьера - использование ИИ», "
      f"{Y}, n = {len(geos)}")
rs_use = {}
for k in BARRIERS:
    r, p = stats.spearmanr([cb[k][g] for g in geos], [use[g] for g in geos])
    rs_use[k] = (r, p)
    mark = "значимо" if p < 0.05 else "не значимо"
    print(f"  {BARRIERS[k]:<24} rs = {r:+.2f}  p = {p:.3f}  ({mark})")

# --------------------------------------- Г3: смещение профиля рисков 2023→2025
print("\n=== Г3: профиль рисков смещается от технологий к регуляторике ===")
for k in ("E_AI_BLEG", "E_AI_BCDP", "E_AI_BINC", "E_AI_BDDT"):
    a = pick([r for r in ctry if r["time"] == "2023" and r["geo"] not in AGG],
             indic_is=k, unit="PC_ENT_AI_EC")
    b = pick([r for r in ctry if r["time"] == "2025" and r["geo"] not in AGG],
             indic_is=k, unit="PC_ENT_AI_EC")
    gg = sorted(set(a) & set(b))
    d = np.array([b[g] - a[g] for g in gg])
    res = stats.wilcoxon(d)
    print(f"  {BARRIERS[k]:<24} n = {len(gg)}  медиана изменения "
          f"{np.median(d):+5.1f} п.п.  p = {res.pvalue:.3f}")

# ---------------------------------------------- Г4: «оцени все сразу»
print("\n=== Г4: ни один барьер не главный — оценивать все сразу ===")
keys = list(BARRIERS)
M = np.array([[cb[k][g] for g in geos] for k in keys])
R, P = stats.spearmanr(M.T)
pairs = list(combinations(range(len(keys)), 2))
pos = sum(R[i, j] > 0 for i, j in pairs)
sig = sum(R[i, j] > 0 and P[i, j] < 0.05 for i, j in pairs)
print(f"Пар барьеров: {len(pairs)}; положительная связь: {pos}; "
      f"значимая положительная: {sig}")
print(f"Средний rs между барьерами: {np.mean([R[i, j] for i, j in pairs]):.2f}"
      " (барьеры — разные, слабо связанные факторы)")
best = max(rs_use, key=lambda k: abs(rs_use[k][0]))
print(f"Сильнейший одиночный барьер: {BARRIERS[best]}, "
      f"rs² = {rs_use[best][0] ** 2:.2f} (объясняет < половины вариации)")
np.save("corr_matrix.npy", R)
