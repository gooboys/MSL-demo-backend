import json
from app.data_analytics.congresses import list_unique_congresses
from app.data_analytics.HCPinteractions import count_unique_interactions
from app.data_analytics.icategories import pie_insight_category_counts, pie_insight_category_percentages
from app.data_analytics.psetting import pie_practice_setting_by_interaction
from app.data_analytics.uniqueMSLs import list_unique_msls


def _extract_rows(content):
  """
  Accepts many possible shapes and returns List[Dict].
  Handles:
  - content is already a list of dicts
  - content is a dict with 'data': [...]
  - content is a dict with 'items': [{'json': {...}}, ...] (common in n8n)
  - content is a JSON string of any of the above
  """
  # If it's a string, try to parse as JSON first
  if isinstance(content, str):
    try:
      content = json.loads(content)
    except Exception:
      # if it's not JSON, treat as a single 'text' record
      return [{"text": content}]

  # If it's already a list of dicts, we’re done
  if isinstance(content, list):
    return content

  # If it's a dict, check common wrappers
  if isinstance(content, dict):
    if "data" in content and isinstance(content["data"], list):
      return content["data"]

    if "items" in content and isinstance(content["items"], list):
      # n8n often wraps each row under {"json": {...}}
      items = content["items"]
      rows = []
      for it in items:
        if isinstance(it, dict) and "json" in it and isinstance(it["json"], dict):
          rows.append(it["json"])
        else:
          rows.append(it)
      return rows

    # if dict but none of the above, treat the dict itself as a single row
    return [content]

  # Anything else → wrap as single row
  return [ {"value": content} ]

def _normalize_fields_inplace(rows):
  # fix names like Raj\" \"Singh -> Raj Singh
  def _clean_name(s):
    if not isinstance(s, str):
      return s
    s = s.replace('\\" \\"', " ").replace('\" \"', " ")
    s = s.replace('\\"', '"').strip('"').strip()
    return " ".join(s.split())

  # fix congress header/value mangling:
  # you printed: 'Congress Name (if applic': {')': 'ASCO 2025'}
  # normalize to a single key: "Congress Name (if applic.)": "ASCO 2025"
  for r in rows:
    # names
    if "KOL Name" in r:
      r["KOL Name"] = _clean_name(r["KOL Name"])
    if "MSL Name" in r:
      r["MSL Name"] = _clean_name(r["MSL Name"])

    # congress
    if "Congress Name (if applic.)" in r:
      val = r["Congress Name (if applic.)"]
      if isinstance(val, dict):
        r["Congress Name (if applic.)"] = next(iter(val.values()), "")
    elif "Congress Name (if applic" in r:
      val = r["Congress Name (if applic"]
      if isinstance(val, dict):
        val = next(iter(val.values()), "")
      r["Congress Name (if applic.)"] = val
      # optionally remove the malformed key
      try:
        del r["Congress Name (if applic"]
      except Exception:
        pass

def get_powerpoint(data):
	# unwrap to rows
	content = data.get("content", data)
	rows = _extract_rows(content)
	_normalize_fields_inplace(rows)

	# --- Now use the helpers ---

	# Pie chart: practice setting
	practice_counts = pie_practice_setting_by_interaction(rows)
	print("Practice setting counts:", practice_counts)

	# Pie chart: insight categories
	category_counts = pie_insight_category_counts(rows)
	category_percentages = pie_insight_category_percentages(rows)
	print("Insight category counts:", category_counts)
	print("Insight category %:", category_percentages)

	# List of congresses
	congresses = list_unique_congresses(rows)
	print("Unique congresses:", congresses)

	# Number of interactions
	n_interactions = count_unique_interactions(rows)
	print("Number of interactions:", n_interactions)

	# Unique MSLs
	msls = list_unique_msls(rows)
	print("Unique MSLs:", msls)

	# after this point you can pass these dicts/lists into your chart/pptx builder
	return {
		"practice_counts": practice_counts,
		"category_counts": category_counts,
		"category_percentages": category_percentages,
		"congresses": congresses,
		"n_interactions": n_interactions,
		"msls": msls
	}