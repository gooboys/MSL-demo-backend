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

def pie_insight_category_percentages(rows: List[Dict]) -> Dict[str, float]:
	"""
	Converts counts to normalized percentages that sum to ~100%.
	"""
	counts = pie_insight_category_counts(rows)
	total = sum(counts.values())
	if total == 0:
		return {k: 0.0 for k in INSIGHT_COLS}
	return {k: (counts.get(k, 0) / total) * 100.0 for k in INSIGHT_COLS}
