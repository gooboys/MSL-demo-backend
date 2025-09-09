from typing import List, Dict
def _clean_name(s):
  if not isinstance(s, str):
      return s
  s = s.replace('\\" \\"', " ").replace('\" \"', " ")
  s = s.replace('\\"', '"').strip('"').strip()
  return " ".join(s.split())

def list_unique_msls(rows: List[Dict]) -> List[str]:
	seen = set()
	out = []
	for r in rows:
		name = _clean_name(r.get("MSL Name"))
		if not name:
			continue
		if name not in seen:
			seen.add(name)
			out.append(name)
	# sort for stable output
	return sorted(out)
