from dataclasses import dataclass
from typing import List, Tuple

@dataclass
class ReportData:
    body_info: List[str]
    deflections: List[Tuple[str, int, int]]
    operating_forces: List[Tuple[float, float]]
    air_data: Tuple[str, str]
    water: Tuple[int, str]
    ultimate: Tuple[int, int, List[str]]
    deflection_val: int
    filename: str