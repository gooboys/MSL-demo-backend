from typing import List, Dict

print("[congresses] importing...")
__all__ = ["list_unique_congresses"]

def _get_congress(row: Dict) -> str:
	if "Congress Name (if applic.)" in row:
		val = row.get("Congress Name (if applic.)")
	elif "Congress Name (if applic" in row:
		val = row.get("Congress Name (if applic")
	else:
		return ""
	if isinstance(val, dict):
		return next(iter(val.values()), "") or ""
	return str(val or "").strip()

def list_unique_congresses(rows: List[Dict]) -> List[str]:
	seen = set()
	out = []
	for r in rows:
		name = _get_congress(r)
		if not name:
			continue
		if name not in seen:
			seen.add(name)
			out.append(name)
	print("[congresses] list_unique_congresses ->", len(out))
	return sorted(out)

print("[congresses] import OK, exports:", __all__)