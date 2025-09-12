# -*- coding: utf-8 -*-
from __future__ import annotations
from typing import List, Dict, Optional


def tri(v: Optional[bool]) -> str:
    """Return "Да"/"Нет"/"—" depending on bool value.

    Parameters
    ----------
    v: Optional[bool]
        Value to format.
    """
    return "Да" if v is True else "Нет" if v is False else "—"


def recommendation(status: List[str], mapping: Dict[str, str]) -> Optional[str]:
    """Convert status codes to a human readable recommendation.

    Parameters
    ----------
    status: List[str]
        List of status codes.
    mapping: Dict[str, str]
        Mapping from status code to recommendation text.
    """
    recs = [mapping.get(s) for s in status if mapping.get(s)]
    return "; ".join(recs) if recs else None
