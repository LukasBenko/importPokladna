from __future__ import annotations

from dataclasses import dataclass, field
from typing import List


@dataclass
class Item:
    skratka_typu_ppd: str
    suma_ppd: str
    poznamka_ppd: str = ""
    skratka_os: str = ""
    skratka_eo: str = ""


@dataclass
class Doc:
    skratka_pk: str
    druh_pd: str
    datum_pd: str
    ucel_pd: str
    komu_od: str = ""
    items: List[Item] = field(default_factory=list)
