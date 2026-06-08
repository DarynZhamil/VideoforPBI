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

    result = {
        "meta": {
            "document": "Коллективные договора КМГ",
            "source_file": SRC_XLSX.name,
            "mrp_2026": 4325,
            "currency": "тенге",
            "format": "long",  # одна запись = (code × dzo); пустые ячейки исключены
            "fields": {
                "code":  "Код выплаты (OPL 1, MP 8, PEN 3 и т.д.)",
                "name":  "Наименование (расшифровка) кода",
                "unit":  "Единица измерения (МРП / ЧТС / МТС/МДО / % / тенге / кал дни)",
                "dzo":   "Дочерняя организация",
                "value": "Размер выплаты (для svod_razmer) или сумма в тенге (для svod_summa)",
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
        "svod_razmer": svod,  # размер выплат (МРП / ЧТС / % / суммы)
        "svod_summa": summa,  # денежный эквивалент по факту
    }

    OUT_JSON.write_text(
        json.dumps(result, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    size_kb = OUT_JSON.stat().st_size / 1024
    chars = len(json.dumps(result, ensure_ascii=False))
    print(f"OK -> {OUT_JSON}")
    print(f"  размер: {size_kb:,.1f} KB")
    print(f"  строк: размер={len(svod)}, сумма={len(summa)}")
    print(f"  ~токенов: {chars // 3:,}")


if __name__ == "__main__":
    main()
