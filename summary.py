from __future__ import annotations
from dataclasses import dataclass
from typing import Dict, List
import pandas as pd
from classification import QuestionInfo, QuestionType

@dataclass
class QuestionSummary:
    question: QuestionInfo
    table: pd.DataFrame  # колонки: ["Варіант відповіді", "Кількість", "%"]


def _build_summary_for_series(
    series: pd.Series, question: QuestionInfo
) -> QuestionSummary:
    """
    Формує підсумкову таблицю для одного питання.
    """
    v = series.dropna()

    if v.empty or question.qtype in (QuestionType.OPEN, QuestionType.TECHNICAL):
        table = pd.DataFrame(columns=["Варіант відповіді", "Кількість", "%"])
        return QuestionSummary(question=question, table=table)

    counts = v.astype(str).str.strip().value_counts().sort_index()
    total = counts.sum()
    perc = (counts / total * 100).round(1)

    table = pd.DataFrame(
        {
            "Варіант відповіді": counts.index,
            "Кількість": counts.values,
            "%": perc.values,
        }
    )

    return QuestionSummary(question=question, table=table)


def build_all_summaries(
    df: pd.DataFrame,
    qinfo: Dict[str, QuestionInfo],
) -> List[QuestionSummary]:
    """
    Формує спискок зведених таблиць для всіх релевантних питань.
    """
    summaries: List[QuestionSummary] = []

    for col, info in qinfo.items():
        if info.qtype in (QuestionType.OPEN, QuestionType.TECHNICAL):
            continue
        summaries.append(_build_summary_for_series(df[col], info))

    return summaries
