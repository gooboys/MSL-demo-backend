from typing import List, Dict

def count_unique_interactions(rows: List[Dict]) -> int:
	ids = set()
	missing = 0
	for r in rows:
		id_val = str(r.get("ID", "")).strip()
		if id_val:
			ids.add(id_val)
		else:
			# count rows without IDs as distinct interactions
			missing += 1
	return len(ids) + missing
