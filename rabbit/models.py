from dataclasses import dataclass
from typing import List

@dataclass
class MatchCandidate():
    suggested: str
    score: float
    date: str 

@dataclass
class MultiMatchResult():
    original: str
    candidates: List[MatchCandidate]