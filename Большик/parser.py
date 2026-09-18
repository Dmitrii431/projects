#!/usr/bin/env python3
"""
Парсер файла начислений ЖКХ.

Формат строки (разделитель ';', строка заканчивается ';'):
    лицевой счёт ; ФИО ; адрес ; период ; сумма [; прибор ; показание]...

Адрес внутри себя разделён запятыми:
    населённый пункт, улица, дом, квартира(ы)
улицы может не быть, квартиры может не быть или быть несколько.

Корректные строки грузятся в SQLite, некорректные выгружаются в CSV
с номером строки и причиной отбраковки.

Запуск:
    python3 parser.py --input Testovye_dannye.txt --db data.db --errors errors.csv
"""

import argparse
import csv
import re
import sqlite3
import sys
import time
from pathlib import Path

RE_ACCOUNT = re.compile(r"^\d+$")
RE_FIO = re.compile(r"^[^\d;]+$", re.UNICODE)
RE_PERIOD = re.compile(r"^(\d{1,2})(\d{2})$")
RE_AMOUNT = re.compile(r"^-?\d+\.\d{2}$")
RE_VALUE = re.compile(r"^-?\d+(\.\d+)?$")
RE_WORD = re.compile(r"[^\W\d_]+", re.UNICODE)
RE_SERIAL = re.compile(r"^(\d{6,})\b")


class Rejected(Exception):
    """Строка не соответствует формату."""

    def __init__(self, message, kind=None):
        super().__init__(message)
        self.kind = kind or message.split(":")[0]


def is_house(part):
    """Похоже ли на номер дома: 12, 12а, '2 б', 12/1, 51-Б, '13 корп. 3', №18А.

    Отличается от названия улицы тем, что начинается с цифры (или №),
    короткое и не содержит длинных слов ('3 Интернационала ул' — улица).
    """
    if not part or len(part) > 14:
        return False
    if not (part[0].isdigit() or part[0] == "№"):
        return False
    return all(len(w) <= 5 for w in RE_WORD.findall(part))


def parse_address(raw):
    """'Весь, Полевая, 2, 2' -> (нас. пункт, улица|None, дом, [квартиры])

    Дом ищется как первая часть (кроме самой первой), похожая на номер дома:
    всё до неё — населённый пункт и улица, всё после — квартиры.
    Так разбираются и адреса без улицы ('Федоровка д, 81'), и адреса,
    где сам населённый пункт содержит запятую
    ('Собинка г, Б. Сокольники, Клязьменская, 3, 0').
    """
    parts = [p.strip() for p in raw.split(",")]
    parts = [p for p in parts if p != ""]
    if len(parts) < 2:
        raise Rejected("адрес: меньше двух частей")

    house_idx = next((i for i in range(1, len(parts)) if is_house(parts[i])), None)
    if house_idx is None:
        raise Rejected("адрес: не найден номер дома")

    head = parts[:house_idx]
    house = parts[house_idx]
    apartments = parts[house_idx + 1:]

    if len(head) == 1:
        locality, street = head[0], None
    else:
        locality, street = ", ".join(head[:-1]), head[-1]

    clean = []
    for apt in apartments:
        if apt == "0":  # квартиры нет
            continue
        if apt not in clean:
            clean.append(apt)
    return locality, street, house, clean


def parse_line(line):
    """Разбирает строку. Бросает Rejected, если формат нарушен."""
    if not line.endswith(";"):
        raise Rejected("строка не заканчивается символом ';'")

    fields = line.split(";")[:-1]

    if len(fields) < 5:
        raise Rejected(f"полей {len(fields)}, минимум 5", "меньше пяти полей")
    if len(fields) % 2 == 0:
        raise Rejected(
            f"полей {len(fields)}: чётное число, пары прибор/показание неполные",
            "чётное число полей (нет ФИО или неполная пара прибор/показание)",
        )

    account, fio, address, period, amount = (f.strip() for f in fields[:5])

    if not RE_ACCOUNT.match(account):
        raise Rejected("лицевой счёт: не число")
    if not fio or not RE_FIO.match(fio):
        raise Rejected("ФИО: пустое или содержит недопустимые символы")

    locality, street, house, apartments = parse_address(address)

    m = RE_PERIOD.match(period)
    if not m:
        raise Rejected(f"период '{period}': ожидается ММГГ", "некорректный период")
    month, year = int(m.group(1)), 2000 + int(m.group(2))
    if not 1 <= month <= 12:
        raise Rejected(f"период '{period}': месяц вне диапазона 1-12", "некорректный период")

    if not RE_AMOUNT.match(amount):
        raise Rejected(
            f"сумма '{amount}': ожидается число с двумя знаками после точки",
            "некорректная сумма начисления",
        )

    meters = []
    for i in range(5, len(fields), 2):
        title = fields[i].strip()
        value = fields[i + 1].strip()
        if not title:
            raise Rejected("прибор учёта: пустое название", "пустое название прибора учёта")
        if not RE_VALUE.match(value):
            raise Rejected(
                f"показание '{value}' прибора '{title}': не число",
                "показание прибора не число",
            )
        serial = RE_SERIAL.match(title)
        meters.append((serial.group(1) if serial else None, title, float(value)))

    return {
        "account": account,
        "fio": fio,
        "locality": locality,
        "street": street,
        "house": house,
        "apartments": apartments,
        "month": month,
        "year": year,
        "period_raw": period,
        "amount": float(amount),
        "meters": meters,
    }


class Loader:
    """Загрузка разобранных строк в SQLite с кэшем справочников."""

    def __init__(self, conn):
        self.conn = conn
        self.localities = {}
        self.streets = {}
        self.addresses = {}
        self.accounts = {}
        self.meters = {}

    def _locality(self, name):
        key = name
        if key not in self.localities:
            cur = self.conn.execute(
                "INSERT INTO localities(name) VALUES (?) ON CONFLICT(name) DO UPDATE SET name=name RETURNING id",
                (name,),
            )
            self.localities[key] = cur.fetchone()[0]
        return self.localities[key]

    def _street(self, locality_id, name):
        key = (locality_id, name)
        if key not in self.streets:
            cur = self.conn.execute(
                "INSERT INTO streets(locality_id, name) VALUES (?, ?) "
                "ON CONFLICT(locality_id, name) DO UPDATE SET name=name RETURNING id",
                key,
            )
            self.streets[key] = cur.fetchone()[0]
        return self.streets[key]

    def _address(self, locality_id, street_id, house):
        key = (locality_id, street_id, house)
        if key not in self.addresses:
            cur = self.conn.execute(
                "INSERT INTO addresses(locality_id, street_id, house) VALUES (?, ?, ?) "
                "ON CONFLICT(locality_id, street_id, house) DO UPDATE SET house=house RETURNING id",
                key,
            )
            self.addresses[key] = cur.fetchone()[0]
        return self.addresses[key]

    def _account(self, number, fio, address_id):
        if number not in self.accounts:
            cur = self.conn.execute(
                "INSERT INTO accounts(account_number, fio, address_id) VALUES (?, ?, ?) "
                "ON CONFLICT(account_number) DO UPDATE SET account_number=account_number RETURNING id",
                (number, fio, address_id),
            )
            self.accounts[number] = cur.fetchone()[0]
        return self.accounts[number]

    def _meter(self, account_id, serial, title):
        key = (account_id, title)
        if key not in self.meters:
            cur = self.conn.execute(
                "INSERT INTO meters(account_id, serial, title) VALUES (?, ?, ?) "
                "ON CONFLICT(account_id, title) DO UPDATE SET title=title RETURNING id",
                (account_id, serial, title),
            )
            self.meters[key] = cur.fetchone()[0]
        return self.meters[key]

    def load(self, rec):
        loc = self._locality(rec["locality"])
        street = self._street(loc, rec["street"]) if rec["street"] else None
        addr = self._address(loc, street, rec["house"])
        acc = self._account(rec["account"], rec["fio"], addr)

        for apt in rec["apartments"]:
            self.conn.execute(
                "INSERT OR IGNORE INTO account_apartments(account_id, apartment) VALUES (?, ?)",
                (acc, apt),
            )

        cur = self.conn.execute(
            "INSERT INTO charges(account_id, period_month, period_year, period_raw, amount) "
            "VALUES (?, ?, ?, ?, ?) ON CONFLICT(account_id, period_year, period_month) DO NOTHING "
            "RETURNING id",
            (acc, rec["month"], rec["year"], rec["period_raw"], rec["amount"]),
        )
        row = cur.fetchone()
        if row is None:
            raise Rejected("дубликат: начисление по этому счёту за этот период уже загружено")
        charge = row[0]

        for serial, title, value in rec["meters"]:
            meter = self._meter(acc, serial, title)
            self.conn.execute(
                "INSERT OR IGNORE INTO readings(charge_id, meter_id, value) VALUES (?, ?, ?)",
                (charge, meter, value),
            )


def main():
    ap = argparse.ArgumentParser(description="Парсер файла начислений ЖКХ")
    ap.add_argument("--input", required=True, help="исходный текстовый файл")
    ap.add_argument("--db", default="data.db", help="файл базы SQLite")
    ap.add_argument("--errors", default="errors.csv", help="CSV с отбракованными строками")
    ap.add_argument("--schema", default=str(Path(__file__).with_name("schema.sql")))
    ap.add_argument("--encoding", default="cp1251")
    args = ap.parse_args()

    db_path = Path(args.db)
    if db_path.exists():
        db_path.unlink()

    conn = sqlite3.connect(db_path)
    conn.executescript(Path(args.schema).read_text(encoding="utf-8"))
    conn.execute("PRAGMA journal_mode = OFF")
    conn.execute("PRAGMA synchronous = OFF")

    loader = Loader(conn)
    ok = bad = 0
    reasons = {}
    started = time.time()

    with open(args.input, encoding=args.encoding, errors="replace") as src, \
         open(args.errors, "w", encoding="utf-8-sig", newline="") as err:
        writer = csv.writer(err, delimiter=";", quoting=csv.QUOTE_MINIMAL)
        writer.writerow(["line_number", "reason", "raw_line"])

        for n, raw in enumerate(src, 1):
            line = raw.rstrip("\r\n")
            if not line.strip():
                continue
            try:
                loader.load(parse_line(line))
                ok += 1
            except Rejected as e:
                bad += 1
                reasons[e.kind] = reasons.get(e.kind, 0) + 1
                writer.writerow([n, str(e), line])
            if n % 50000 == 0:
                conn.commit()

    conn.commit()
    counts = {
        t: conn.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
        for t in ("localities", "streets", "addresses", "accounts",
                  "account_apartments", "charges", "meters", "readings")
    }
    conn.close()

    print(f"Обработано строк : {ok + bad}")
    print(f"Загружено в БД   : {ok}")
    print(f"Отбраковано      : {bad}  -> {args.errors}")
    print(f"Время            : {time.time() - started:.1f} c")
    print("\nПричины отбраковки:")
    for reason, cnt in sorted(reasons.items(), key=lambda x: -x[1]):
        print(f"  {cnt:>7}  {reason}")
    print("\nСтроки в таблицах:")
    for table, cnt in counts.items():
        print(f"  {table:<20} {cnt}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
