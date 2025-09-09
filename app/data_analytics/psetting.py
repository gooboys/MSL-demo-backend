from collections import Counter
from typing import List, Dict, Tuple, Optional

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

def _clean_name(s: Optional[str]) -> Optional[str]:
	if not isinstance(s, str):
		return s
	s = s.replace('\\" \\"', " ").replace('\" \"', " ")
	s = s.replace('\\"', '"').strip('"').strip()
	return " ".join(s.split())


def pie_practice_setting_by_interaction(rows: List[Dict]) -> Dict[str, int]:
	"""
	Counts interactions by KOL Practice Setting using unique IDs.
	If multiple rows share the same ID, they count once.
	"""
	id_to_setting: Dict[str, str] = {}
	for r in rows:
		id_val = str(r.get("ID", "")).strip()
		if not id_val:
			# if no ID, treat the row as its own interaction with a synthetic key
			id_val = f"_row_{id(r)}"
		if id_val in id_to_setting:
			continue
		setting = (r.get("KOL Practice Setting") or "").strip() or "Unknown"
		id_to_setting[id_val] = setting
	return dict(Counter(id_to_setting.values()))