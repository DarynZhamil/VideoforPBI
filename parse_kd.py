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


def sheet_to_records(df: pd.DataFrame) -> list[dict]:
    """Конвертирует лист в список словарей (заголовки берём из строки 1)."""
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
            # числа оставляем числами, остальное в строку
            if isinstance(val, (int, float)):
                rec[h] = val
            else:
                s = str(val).strip()
                if s:
                    rec[h] = s
        if rec:
            records.append(rec)
    return records


def main() -> None:
    if not SRC_XLSX.exists():
        raise SystemExit(f"Не найден файл: {SRC_XLSX}")

    print(f"Читаю: {SRC_XLSX.name}")
    df_svod = pd.read_excel(SRC_XLSX, sheet_name="Свод", header=None)
    df_summa = pd.read_excel(SRC_XLSX, sheet_name="Свод сумма", header=None)

    svod = sheet_to_records(df_svod)
    summa = sheet_to_records(df_summa)

    result = {
        "meta": {
            "document": "Коллективные договора КМГ",
            "source_file": SRC_XLSX.name,
            "mrp_2026": 4325,
            "currency": "тенге",
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
