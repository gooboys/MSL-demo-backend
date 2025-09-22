from collections import Counter
from typing import List, Dict

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
