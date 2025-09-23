from collections import Counter
from typing import List, Dict, Any
import re


INSIGHT_COLS = [
	"Access Insights",
	"Patient Management / Care Insights",
	"Clinical Development Insights",
	"Competitive Insights",
	"Product Insights (Drug Science)",
	"Education",
	"Logistics",
	"Other",
	"Adverse Event (AE) Insights"
]

TIERS = [1,2,3]

def pie_insight_category_counts(rows: List[Dict]) -> Dict[str, int]:
	"""
	Tallies category-hits across all rows.
	Each '1' in a category column contributes one hit.
	Overlaps are expected (a single row can increment multiple categories).
	"""
	hits = Counter()
	for r in rows:
		for col in INSIGHT_COLS:
			val = r.get(col, 0)
			try:
				val = int(val)
			except Exception:
				val = 0
			if val == 1:
				hits[col] += 1
	return dict(hits)

def pie_insight_category_counts_raw(rows: List[Dict]) -> Dict[str, int]:
	"""
	Returns raw category counts (no normalization).
	Each '1' in a category column contributes one hit.
	"""
	counts = pie_insight_category_counts(rows)
	return {k: counts.get(k, 0) for k in INSIGHT_COLS}

_TIER_LABELS = {1: "Tier 1", 2: "Tier 2", 3: "Tier 3"}

def kol_tier_counts_pretty(rows: List[Dict[str, Any]]) -> Dict[str, int]:
  """
  Count KOL Tier values (1/2/3) from the 'KOL Tier' column and return pretty labels.
	Accepts values like 1, "1", "Tier 1", "T1".
	"""
  hits = Counter({1: 0, 2: 0, 3: 0})
  for r in rows:
    raw = r.get("KOL Tier", None)
    if raw is None:
      continue
    s = str(raw).strip()
    # Extract first digit 1/2/3 if present
    m = re.search(r"\b([123])\b", s)
    if not m:
      # Also allow T1/T2/T3 or 'Tier1' (no space)
      m = re.search(r"\bT(?:ier)?\s*([123])\b", s, re.IGNORECASE)
    if m:
      tier = int(m.group(1))
      if tier in hits:
        hits[tier] += 1
  # Convert to pretty labels
  return {_TIER_LABELS[k]: v for k, v in hits.items() if v > 0}