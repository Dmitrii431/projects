-- Схема БД для данных начислений ЖКХ (SQLite)
PRAGMA foreign_keys = ON;

DROP TABLE IF EXISTS readings;
DROP TABLE IF EXISTS meters;
DROP TABLE IF EXISTS charges;
DROP TABLE IF EXISTS account_apartments;
DROP TABLE IF EXISTS accounts;
DROP TABLE IF EXISTS addresses;
DROP TABLE IF EXISTS streets;
DROP TABLE IF EXISTS localities;

-- Населённый пункт
CREATE TABLE localities (
    id   INTEGER PRIMARY KEY,
    name TEXT NOT NULL UNIQUE
);

-- Улица (может отсутствовать: адрес вида "Федоровка д, 81")
CREATE TABLE streets (
    id          INTEGER PRIMARY KEY,
    locality_id INTEGER NOT NULL REFERENCES localities(id),
    name        TEXT NOT NULL,
    UNIQUE (locality_id, name)
);

-- Дом
CREATE TABLE addresses (
    id          INTEGER PRIMARY KEY,
    locality_id INTEGER NOT NULL REFERENCES localities(id),
    street_id   INTEGER REFERENCES streets(id),
    house       TEXT NOT NULL,
    UNIQUE (locality_id, street_id, house)
);

-- Лицевой счёт плательщика
CREATE TABLE accounts (
    id             INTEGER PRIMARY KEY,
    account_number TEXT NOT NULL,
    fio            TEXT NOT NULL,
    address_id     INTEGER NOT NULL REFERENCES addresses(id),
    UNIQUE (account_number)
);

-- Квартиры лицевого счёта: 0..n (в файле может не быть или быть несколько)
CREATE TABLE account_apartments (
    id         INTEGER PRIMARY KEY,
    account_id INTEGER NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    apartment  TEXT NOT NULL,
    UNIQUE (account_id, apartment)
);

-- Начисление за период
CREATE TABLE charges (
    id           INTEGER PRIMARY KEY,
    account_id   INTEGER NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    period_month INTEGER NOT NULL CHECK (period_month BETWEEN 1 AND 12),
    period_year  INTEGER NOT NULL,
    period_raw   TEXT NOT NULL,
    amount       NUMERIC NOT NULL,
    UNIQUE (account_id, period_year, period_month)
);

-- Прибор учёта
CREATE TABLE meters (
    id         INTEGER PRIMARY KEY,
    account_id INTEGER NOT NULL REFERENCES accounts(id) ON DELETE CASCADE,
    serial     TEXT,            -- номер прибора, если он выделен в начале поля
    title      TEXT NOT NULL,   -- поле 5(а) целиком
    UNIQUE (account_id, title)
);

-- Показание прибора за период начисления
CREATE TABLE readings (
    id        INTEGER PRIMARY KEY,
    charge_id INTEGER NOT NULL REFERENCES charges(id) ON DELETE CASCADE,
    meter_id  INTEGER NOT NULL REFERENCES meters(id) ON DELETE CASCADE,
    value     NUMERIC NOT NULL,
    UNIQUE (charge_id, meter_id)
);

CREATE INDEX idx_accounts_address ON accounts(address_id);
CREATE INDEX idx_charges_account  ON charges(account_id);
CREATE INDEX idx_readings_meter   ON readings(meter_id);
