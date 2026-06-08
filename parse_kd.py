# -*- coding: utf-8 -*-
"""
Парсер «Коллективные договора.xlsx» → kd_knowledge.json
Берёт два сводных листа (Свод и Свод сумма) и упаковывает в компактный JSON
для AI-ассистента на странице merger.kz.

Запуск:
    python parse_kd.py

После запуска:
    git add kd_knowledge.json
    git commit -m "update KD knowledge base"
    git push --force
"""
import json
import os
from pathlib import Path

import pandas as pd

# ── Пути ────────────────────────────────────────────────────────────────────
SRC_XLSX = Path(
    r"C:\Users\Daryn\SharePoint\Департамент тарифной политики - Инф\Daryn\Дарын"
    r"\PBI\Коллектирный договор\Коллективные договора.xlsx"
)
OUT_JSON = Path(__file__).parent / "kd_knowledge.json"

# Порядок ДЗО в столбцах (как в Excel; пробелы в названиях нормализуются)
DZO_ORDER = [
    "ОМГ", "ММГ", "КБМ", "ЭМГ", "ОСК", "ОТК", "ОКК", "МЭМ",
    "МТК", "ОМС", "ККС", "УДТВ", "Кэтеринг", "КазГПЗ", "Каспий Битум",
]

# Маппинг имени листа в каноническое имя ДЗО (для парсинга 16 листов ДЗО)
SHEET_TO_DZO = {
    "ОМГ": "ОМГ", "ММГ": "ММГ", "КБМ": "КБМ", "ЭМГ": "ЭМГ",
    "ОСК": "ОСК", "ОТК": "ОТК", "ОКК": "ОКК", "МЭМ": "МЭМ",
    "МТК": "МТК", "ОМС ": "ОМС", "ОМС": "ОМС",
    "ККС": "ККС", "УДТВ": "УДТВ", "Кэтеринг": "Кэтеринг",
    "КАЗГПЗ": "КазГПЗ", "КазГПЗ": "КазГПЗ",
    "Кбитум": "Каспий Битум", "Каспий Битум": "Каспий Битум",
    "КМГ секьюрити": "КМГ Секьюрити",
}


_META_COLS = {"№/№", "Код (шифр)", "Наименование (расшифровка) кода", "Един изм"}


def sheet_to_wide_records(df: pd.DataFrame) -> list[dict]:
    """Старый wide-формат: одна запись = код × все ДЗО (для svod_summa)."""
    headers = [str(h).strip() if pd.notna(h) else "" for h in df.iloc[1].tolist()]
    records: list[dict] = []
    for idx in range(2, len(df)):
        row = df.iloc[idx]
        rec: dict = {}
        for col_idx, h in enumerate(headers):
            if not h:
                continue
            val = row[col_idx]
            if pd.isna(val):
                continue
            if isinstance(val, (int, float)):
                rec[h] = val
            else:
                s = str(val).strip()
                if s:
                    rec[h] = s
        if rec:
            records.append(rec)
    return records


def sheet_to_long_records(df: pd.DataFrame) -> list[dict]:
    """Конвертирует широкую таблицу (один код × 15 ДЗО в столбцах)
    в длинный формат: одна запись = (код × ОДНО ДЗО).

    Это критически важно: модель не сможет «съехать» по столбцам,
    т.к. для каждого ДЗО запись содержит ТОЛЬКО его значение.
    Пустые ячейки порождают записи с value=null — модель видит явный «пусто».
    """
    headers = [str(h).strip() if pd.notna(h) else "" for h in df.iloc[1].tolist()]

    records: list[dict] = []
    for idx in range(2, len(df)):
        row = df.iloc[idx]
        code = row[2] if len(row) > 2 else None
        name = row[3] if len(row) > 3 else None
        unit = row[4] if len(row) > 4 else None

        if pd.isna(code) and pd.isna(name):
            continue
        # Строки-заголовки разделов (Код есть, но это просто 'OPL'/'PEN' без номера)
        # — пропускаем, они только структурируют документ
        code_s = str(code).strip() if pd.notna(code) else ""
        name_s = str(name).strip() if pd.notna(name) else ""
        unit_s = str(unit).strip() if pd.notna(unit) else ""
        if code_s and " " not in code_s and pd.isna(unit) and name_s:
            # это header-строка раздела — пропускаем
            continue
        if not code_s:
            continue

        # Проходим по каждому ДЗО-столбцу
        for col_idx, h in enumerate(headers):
            if not h or h in _META_COLS:
                continue
            val = row[col_idx]
            if pd.isna(val):
                continue
            v = val if isinstance(val, (int, float)) else str(val).strip()
            if v == "" or v is None:
                continue
            # Для сумм пропускаем нули (в svod_summa 0 = «фактически не платили»)
            if isinstance(v, (int, float)) and v == 0:
                continue
            records.append({
                "code": code_s,
                "name": name_s,
                "unit": unit_s,
                "dzo": h,
                "value": v,
            })
    return records


def _detect_dzo_columns(df: pd.DataFrame) -> dict:
    """Авто-детекция колонок на листе ДЗО.
    Возвращает dict с ключами: sum, qty, code, size, name.
    Берёт ПРАВЫЕ совпадения для sum/qty — они относятся к 2025 году
    (на листах две группы колонок: 2024 и 2025 годы).
    """
    cols = {"sum": None, "qty": None, "code": None, "size": None, "name": None}
    for r in range(min(8, len(df))):
        for c in range(df.shape[1]):
            v = df.iat[r, c]
            if pd.isna(v):
                continue
            s = str(v)
            if "Сумма тыс" in s:
                cols["sum"] = c  # правое совпадение перезапишет левое = 2025
            elif "Количество единиц" in s:
                cols["qty"] = c
            elif s.strip() == "Код":
                cols["code"] = c
            elif s.strip() == "Размер":
                cols["size"] = c
            elif "Наименование пункта" in s and cols["name"] is None:
                cols["name"] = c
    return cols


def parse_dzo_sheet(df: pd.DataFrame, dzo_name: str) -> list[dict]:
    """Парсит лист одного ДЗО → список записей с численностью и суммой за 2025 год."""
    cols = _detect_dzo_columns(df)
    if cols["code"] is None or (cols["qty"] is None and cols["sum"] is None):
        return []

    records: list[dict] = []
    for i in range(len(df)):
        code_v = df.iat[i, cols["code"]]
        if pd.isna(code_v):
            continue
        code_s = str(code_v).strip()
        # Пропускаем общий шифр "KD" и строки без конкретного кода
        if not code_s or code_s == "KD" or " " not in code_s:
            continue

        qty = df.iat[i, cols["qty"]] if cols["qty"] is not None else None
        sum_ = df.iat[i, cols["sum"]] if cols["sum"] is not None else None
        size = df.iat[i, cols["size"]] if cols["size"] is not None else None
        name = df.iat[i, cols["name"]] if cols["name"] is not None else None

        # Запись имеет смысл только если есть хотя бы один из показателей
        if pd.isna(qty) and pd.isna(sum_):
            continue

        # Компактная запись: только то, чего нет в svod_razmer/summa.
        # name/size опускаем — они уже есть в свод-таблицах по тому же (code, dzo).
        rec: dict = {"code": code_s, "dzo": dzo_name}
        if pd.notna(qty):
            try:
                rec["headcount"] = int(qty)
            except (ValueError, TypeError):
                rec["headcount"] = qty
        if pd.notna(sum_) and isinstance(sum_, (int, float)):
            rec["sum_thousand_tenge"] = round(float(sum_), 1)
        records.append(rec)
    return records


def main() -> None:
    if not SRC_XLSX.exists():
        raise SystemExit(f"Не найден файл: {SRC_XLSX}")

    print(f"Читаю: {SRC_XLSX.name}")
    df_svod = pd.read_excel(SRC_XLSX, sheet_name="Свод", header=None)
    df_summa = pd.read_excel(SRC_XLSX, sheet_name="Свод сумма", header=None)

    # Обе таблицы — LONG формат. Это критично, чтобы модель не съезжала по столбцам ДЗО.
    # svod_summa тоже сжимается: пустых/нулевых ячеек ~85%, остаётся только реальный факт.
    svod = sheet_to_long_records(df_svod)
    summa = sheet_to_long_records(df_summa)

    # Парсим 16 листов ДЗО → численность + сумма за 2025 для каждого (code, dzo)
    xl = pd.ExcelFile(SRC_XLSX)
    dzo_details: list[dict] = []
    for sheet_name in xl.sheet_names:
        if sheet_name in ("Свод", "Свод сумма"):
            continue
        dzo_canonical = SHEET_TO_DZO.get(sheet_name)
        if not dzo_canonical:
            print(f"  ⚠ лист {sheet_name!r} не сопоставлен ДЗО — пропуск")
            continue
        df_dzo = pd.read_excel(SRC_XLSX, sheet_name=sheet_name, header=None)
        recs = parse_dzo_sheet(df_dzo, dzo_canonical)
        dzo_details.extend(recs)
        print(f"  list {sheet_name!r:25s} -> {len(recs)} records")

    result = {
        "meta": {
            "document": "Коллективные договора КМГ",
            "source_file": SRC_XLSX.name,
            "mrp_2026": 4325,
            "format": "long",  # одна запись = (code × dzo); пустые ячейки исключены
            "fields_svod_razmer": {
                "code":  "Код выплаты (OPL 1, MP 8, PEN 3 и т.д.)",
                "name":  "Наименование (расшифровка) кода",
                "unit":  "Единица измерения (МРП / ЧТС / МТС/МДО / % / тенге / кал дни)",
                "dzo":   "Дочерняя организация",
                "value": "Размер ПОЛОЖЕННОЙ выплаты — норматив (в единицах unit)",
            },
            "fields_svod_summa": {
                "code":  "Код выплаты",
                "name":  "Наименование",
                "unit":  "Поле справочное (норматив unit) — НЕ единица value",
                "dzo":   "Дочерняя организация",
                "value": "ФАКТ выплат за период в ТЫСЯЧАХ тенге (тыс.тнг). Умножай на 1000 для тенге.",
            },
            "fields_dzo_details": {
                "code": "Код выплаты (как в svod_razmer/summa)",
                "dzo": "Дочерняя организация",
                "headcount": "Численность работников, получивших эту выплату в 2025",
                "sum_thousand_tenge": "Фактическая сумма за 2025 в ТЫСЯЧАХ тенге (как в svod_summa)",
                "_note": "Данные за 2025 год. ЭМГ и КМГ Секьюрити отсутствуют — у них другая структура листа.",
            },
            "terms": {
                "МРП": "Месячный расчётный показатель (4 325 тенге в 2026 году)",
                "ЧТС": "Часовая тарифная ставка",
                "ТС": "Тарифная ставка / Должностной оклад (ДО)",
                "МТС": "Месячная тарифная ставка",
                "МДО": "Месячный должностной оклад",
                "ВУТ": "Вредные условия труда",
                "ДЗО": "Дочерние и зависимые организации КМГ",
            },
        },
        "dzo_list": DZO_ORDER,
        "svod_razmer": svod,    # норматив: что положено по КД (МРП / ЧТС / % / суммы)
        "svod_summa": summa,    # факт: сколько выплачено за период (в тыс.тнг)
        "dzo_details": dzo_details,  # численность + сумма из 16 листов ДЗО (за 2025)
    }

    OUT_JSON.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    size_kb = OUT_JSON.stat().st_size / 1024
    chars = len(json.dumps(result, ensure_ascii=False))
    print(f"OK -> {OUT_JSON}")
    print(f"  размер: {size_kb:,.1f} KB")
    print(f"  строк: svod_razmer={len(svod)}, svod_summa={len(summa)}, dzo_details={len(dzo_details)}")
    print(f"  ~токенов: {chars // 3:,}")


if __name__ == "__main__":
    main()
