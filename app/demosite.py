import json
import io
import base64
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from app.data_analytics.congresses import list_unique_congresses
from app.data_analytics.hcp_interactions import count_unique_interactions
from app.data_analytics.icategories import pie_insight_category_counts, pie_insight_category_percentages
from app.data_analytics.psetting import pie_practice_setting_by_interaction
from app.data_analytics.unique_msls import list_unique_msls
import traceback
print("[demosite] importing congresses...")

try:
	import app.data_analytics.congresses as _c
	print("[demosite] congresses.__file__:", getattr(_c, "__file__", None))
	print("[demosite] has list_unique_congresses?:", hasattr(_c, "list_unique_congresses"))
	print("[demosite] dir sample:", [n for n in dir(_c) if "congress" in n.lower() or "list" in n.lower()])
	# keep this line last; it will still raise if the symbol truly isn't there
	from app.data_analytics.congresses import list_unique_congresses
	print("[demosite] direct import succeeded")
except Exception as e:
	print("[demosite] import failed:", repr(e))
	traceback.print_exc()
	raise

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

def _png_b64(png_bytes: bytes) -> str:
	"""
	Return a base64-encoded PNG string (no data URI prefix) for JSON-safe transport.
	"""
	return base64.b64encode(png_bytes).decode("utf-8")

def _create_pie_chart(data: dict[str, int], title: str) -> bytes:
	"""
	Create a pie chart from raw counts and return PNG bytes.
	Zero-value keys are dropped; empty -> 'No Data'.
	"""
	# filter zeros
	data = {k: v for k, v in (data or {}).items() if v > 0}
	if not data:
		data = {"No Data": 1}

	labels = list(data.keys())
	values = list(data.values())

	fig, ax = plt.subplots(figsize=(6, 4))
	ax.set_title(title)
	ax.pie(
		values,
		labels=labels,
		autopct=lambda pct: f"{pct:.0f}%",
		startangle=90
	)
	ax.axis("equal")

	buf = io.BytesIO()
	plt.savefig(buf, format="png", bbox_inches="tight", dpi=200)
	plt.close(fig)
	return buf.getvalue()

def get_powerpoint(data):
	# unwrap to rows
	content = data.get("content", data)
	rows = _extract_rows(content)
	_normalize_fields_inplace(rows)

	# --- Extracted metrics ---

	# Pie chart: practice setting (by unique interaction/ID)
	practice_counts = pie_practice_setting_by_interaction(rows)
	print("Practice setting counts:", practice_counts)

	# Pie chart: insight categories (raw category-hits, overlaps allowed)
	category_counts = pie_insight_category_counts(rows)
	print("Insight category counts:", category_counts)

	# List of congresses
	congresses = list_unique_congresses(rows)
	print("Unique congresses:", congresses)

	# Number of interactions
	n_interactions = count_unique_interactions(rows)
	print("Number of interactions:", n_interactions)

	# Unique MSLs
	msls = list_unique_msls(rows)
	print("Unique MSLs:", msls)

	# --- Build PNG pies (raw counts) ---
	practice_pie_png = _create_pie_chart(practice_counts, "HCP Practice Setting")
	category_pie_png = _create_pie_chart(category_counts, "Insight Categories")

	# Base64 for n8n (JSON-safe)
	practice_pie_b64 = _png_b64(practice_pie_png)
	category_pie_b64 = _png_b64(category_pie_png)

	# Return payload for n8n
	return {
		"practice_counts": practice_counts,
		"category_counts": category_counts,
		"congresses": congresses,
		"n_interactions": n_interactions,
		"msls": msls,
		"practice_pie_png_b64": practice_pie_b64,
		"category_pie_png_b64": category_pie_b64,
		"_meta": {
			"practice_pie_title": "HCP Practice Setting",
			"category_pie_title": "Insight Categories",
			"images_format": "png",
			"images_encoding": "base64"
		}
	}