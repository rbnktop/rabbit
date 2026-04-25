import pandas as pd
from rapidfuzz import fuzz, process
from typing import List, Optional, Callable
from models import MatchCandidate, MultiMatchResult
import unicodedata


def normalize_text(text: str) -> str:
    """Removes accents, converts to lowercase, and strips extra whitespace."""
    
    if not isinstance(text, str):
        text = str(text)
    text = text.lower().strip()
    nksfd = unicodedata.normalize('NFKD', text)
    return "".join([c for c in nksfd if not unicodedata.combining(c)])

def get_match_score(original, candidate):
    clean_org = normalize_text(original)
    clean_cand = normalize_text(candidate)

    score = fuzz.token_sort_ratio(clean_org, clean_cand) / 100.0
    return score

def load_excel_data(path: str) -> pd.DataFrame:
    return pd.read_excel(path)

def parse_search_items(text: str) -> List[str]:
    return [line.strip() for line in text.splitlines() if line.strip()]

def build_excel_pool(df: pd.DataFrame) -> dict:
    """Creates a dictionary {Name: Most_Recent_Date}."""
    pool_map = {}
    
    for col in df.columns:
        try:
            recent_date = pd.to_datetime(col)
        except:
            continue

        names_in_col = df[col].dropna().unique()
        
        for name in names_in_col:
            clean_name = normalize_text(str(name))
            if clean_name not in pool_map or recent_date > pool_map[clean_name]:
                pool_map[clean_name] = recent_date
                
    return pool_map


def find_smart_matches(items, pool_map, threshold=0.49, auto_approve=0.98, progress_callback=None):
    results = []
    pool= list(pool_map.keys())

    for i, item in enumerate(items):
        best_candidates = []
        
        for candidate in pool:
            score = get_match_score(item, candidate)
            raw_date = pool_map.get(candidate)
            formatted_date = raw_date.strftime('%d/%m/%Y') if raw_date else "N/A"

            if score >= auto_approve:
                best_candidates = [MatchCandidate(suggested=candidate, score=score, date=formatted_date)]
                break

            if score >= threshold:
                best_candidates.append(MatchCandidate(suggested=candidate, score=score, date=formatted_date))

        best_candidates.sort(key=lambda x: x.score, reverse=True)
        candidates_list = best_candidates[:5] 
        results.append(MultiMatchResult(original=item, candidates=candidates_list))
        
        if progress_callback:
            progress_callback(i + 1, len(items))

    return results

def extract_dates_for_match(df: pd.DataFrame, match_value: str) -> str:
    if not match_value or match_value == "NONE":
        return "FOTO"
    
    match_str = str(match_value).strip()
    found_dates = []
    
    for col_name in df.columns:
        column_data_clean = df[col_name].astype(str).apply(normalize_text)
        
        if column_data_clean.eq(match_str).any():
            try: 
                found_dates.append(pd.to_datetime(col_name))
            except: 
                continue
    
    if found_dates:
        return max(found_dates).strftime("%d/%m/%Y")
    
    return "FOTO"