from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Dict

import pandas as pd


class QuestionType(Enum):
    SCALE = "Шкальна (1–5)"
    BINARY = "Дихотомічна (Так/Ні)"
    CATEGORICAL = "Категоріальна"
    OPEN = "Відкрита / текстова"
    TECHNICAL = "Технічне поле"


BINARY_SET = {"так", "ні", "не знаю", "yes", "no", "don't know", "dont know"}


@dataclass
class QuestionInfo:
    code: str
    text: str
    qtype: QuestionType


def detect_type(series: pd.Series) -> QuestionType:
    """
    Евристичне визначення типу питання за розподілом відповідей.
    :param series: стовпець із відповідями.
    """
    v = series.dropna()
    if v.empty:
        return QuestionType.OPEN

    v_str = v.astype(str).str.strip()
    uniq = set(v_str.unique())
    n_unique = len(uniq)
    n = len(v_str)

    # 1) Шкала 1–5 (Лайкерт)
    if uniq.issubset({"1", "2", "3", "4", "5"}):
        return QuestionType.SCALE

    # 2) Дихотомічні запитання
    low = {x.lower() for x in uniq}
    if low.issubset(BINARY_SET):
        return QuestionType.BINARY

    # 3) Категоріальні варіанти (обмежений набір повторюваних відповідей)
    if n_unique <= 15 and n_unique / max(n, 1) <= 0.7:
        return QuestionType.CATEGORICAL

    # 4) Все інше – відкриті відповіді
    return QuestionType.OPEN


def classify_questions(
    df: pd.DataFrame,
    technical_columns: int = 1,
) -> Dict[str, QuestionInfo]:
    """
    Класифікує всі стовпці таблиці (крім технічних) за типом питання.

    :param df: DataFrame із відповідями.
    :param technical_columns: скільки перших стовпців вважаємо технічними
                              (наприклад, 'Позначка часу').
    :return: словник {ім'я стовпця: QuestionInfo}.
    """
    result: Dict[str, QuestionInfo] = {}

    for idx, col in enumerate(df.columns):
        text = str(col).strip()
        if idx < technical_columns:
            qtype = QuestionType.TECHNICAL
        else:
            qtype = detect_type(df[col])

        code = f"Q{idx - technical_columns + 1}" if idx >= technical_columns else "-"
        result[col] = QuestionInfo(code=code, text=text, qtype=qtype)

    return result
