# -*- coding: utf-8 -*-
"""Загрузка открытых данных Eurostat (обследование использования ИКТ
предприятиями, наборы isoc_eb_ain2 и isoc_eb_ai) в CSV."""
import csv
import itertools
import json
import urllib.request

API = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/"
TECH = ["E_AI_TANY", "E_AI_EC"]
BARRIERS = ["E_AI_BCST", "E_AI_BLE", "E_AI_BINC", "E_AI_BDDT", "E_AI_BCDP",
            "E_AI_BLEG", "E_AI_BEC", "E_AI_BNU"]


def fetch(dataset, **flt):
    q = "&".join(f"{k}={v}" for k, vals in flt.items()
                 for v in (vals if isinstance(vals, list) else [vals]))
    url = f"{API}{dataset}?format=JSON&lang=en&{q}"
    d = json.load(urllib.request.urlopen(url, timeout=60))
    dims = d["id"]
    cats = [sorted(d["dimension"][k]["category"]["index"].items(),
                   key=lambda x: x[1]) for k in dims]
    sizes = d["size"]
    rows = []
    for combo in itertools.product(*[range(s) for s in sizes]):
        flat = 0
        for i, c in enumerate(combo):
            flat = flat * sizes[i] + c
        v = d["value"].get(str(flat))
        if v is None:
            continue
        rows.append({k: cats[i][combo[i]][0] for i, k in enumerate(dims)}
                    | {"value": v})
    return rows


def save(rows, name):
    with open(name, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0]))
        w.writeheader()
        w.writerows(rows)
    print(name, len(rows))


if __name__ == "__main__":
    nace = ["J62_J63", "C10-S951_X_K"]
    # 1) динамика использования ИИ: IT-сектор и экономика в целом, ЕС-27
    save(fetch("isoc_eb_ain2", geo="EU27_2020", nace_r2=nace,
               indic_is=TECH + ["E_AI_TX"], unit=["PC_ENT"]), "eu_trend.csv")
    # 2) барьеры: доля предприятий, рассматривавших ИИ, но не внедривших
    save(fetch("isoc_eb_ain2", geo="EU27_2020", nace_r2=nace,
               indic_is=BARRIERS, unit=["PC_ENT", "PC_ENT_AI_EC"]),
         "eu_barriers.csv")
    # 3) страны: использование ИИ и барьеры в IT-секторе, 2023-2025
    save(fetch("isoc_eb_ain2", nace_r2="J62_J63",
               time=["2023", "2024", "2025"],
               indic_is=TECH + BARRIERS, unit=["PC_ENT", "PC_ENT_AI_EC"]),
         "countries_it.csv")
    # 4) размер предприятия (все отрасли), ЕС-27
    save(fetch("isoc_eb_ai", geo="EU27_2020",
               indic_is=TECH + BARRIERS, unit=["PC_ENT", "PC_ENT_AI_EC"]),
         "eu_size.csv")
