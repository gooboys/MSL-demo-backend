import json
import io
import base64
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from app.data_analytics.congresses import list_unique_congresses
from app.data_analytics.hcp_interactions import count_unique_interactions
from app.data_analytics.icategories import pie_insight_category_counts
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

def _create_pie_chart(data: dict[str, int]) -> bytes:
  """
  Create a pie chart (PNG bytes) with:
    - Legend at the bottom (color-coded key)
    - Slice labels hidden; counts shown on/near slices
    - Small slices get count labels outside with leader lines
  """
  # Decides color scheme
  palette = [
    "#08306B",  # very dark navy blue
    "#08519C",  # strong blue
    "#2171B5",  # medium blue
    "#41B6C4",  # teal
    "#7FCDBB",   # aquamarine
    "#4292C6",  # lighter blue
    "#6BAED6",  # sky blue
    "#9ECAE1"  # pale blue
  ]
  # Filter zeros & handle empty
  data = {k: int(v) for k, v in (data or {}).items() if int(v) > 0}
  if not data:
    data = {"No Data": 1}

  labels = list(data.keys())
  values = np.array(list(data.values()), dtype=float)
  total = values.sum()

  # Threshold below which labels go outside (as a % of total)
  OUTSIDE_THRESHOLD = 6.0  # percent

  fig, ax = plt.subplots(figsize=(5.5, 4))

  # Draw pie without labels (legend will handle labels)
  wedges, _ = ax.pie(
    values,
    labels=None,
    startangle=90,
    wedgeprops=dict(linewidth=0.5, edgecolor="white"),
    colors=palette[:len(values)]
  )
  ax.axis("equal")

  # Add count labels inside or outside depending on size
  for wedge, val in zip(wedges, values):
    pct = (val / total) * 100.0
    theta = (wedge.theta2 + wedge.theta1) / 2.0
    theta_rad = np.deg2rad(theta)

    # Default: inside
    r = 0.7
    x = r * np.cos(theta_rad)
    y = r * np.sin(theta_rad)

    if pct < OUTSIDE_THRESHOLD:
      # Outside with a leader line
      r_out = 1.15
      x_out = r_out * np.cos(theta_rad)
      y_out = r_out * np.sin(theta_rad)

      ax.annotate(
        f"{int(val):,}",
        xy=(np.cos(theta_rad), np.sin(theta_rad)),  # anchor at unit circle
        xytext=(x_out, y_out),
        ha="left" if x_out >= 0 else "right",
        va="center",
        arrowprops=dict(arrowstyle="-", lw=0.8, shrinkA=0, shrinkB=0),
      )
    else:
      ax.text(x, y, f"{int(val):,}", ha="center", va="center", fontsize=10)

  # Legend at the bottom, multi-column if many labels
  ncol = 2 if len(labels) <= 6 else 3
  ax.legend(
    wedges,
    labels,
    loc="upper center",
    bbox_to_anchor=(0.5, -0.08),
    ncol=2,              # ✅ always 2 items per row
    frameon=False,
    handlelength=1.0,
    handletextpad=0.6,
    columnspacing=1.4,
    fontsize=9,
  )

  buf = io.BytesIO()
  plt.savefig(buf, format="png", bbox_inches="tight", dpi=200)
  plt.close(fig)
  return buf.getvalue()

def data_preprocess(data):
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
	practice_pie_png = _create_pie_chart(practice_counts)
	category_pie_png = _create_pie_chart(category_counts)

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
		"practice_pie_png_b64": practice_pie_png,
		"category_pie_png_b64": category_pie_png,
    "insight_count": len(rows),
		"_meta": {
			"practice_pie_title": "HCP Practice Setting",
			"category_pie_title": "Insight Categories",
			"images_format": "png",
			"images_encoding": "base64"
		}
	}

def second_process(data):
  settings = data["practice_counts"]
  academic_count = settings.get('Academic Center', 0)
  other_count = sum(v for k, v in settings.items() if k != 'Academic Center')
  stats = {
    'graph1': data["category_pie_png_b64"], # Done
    'graph2': data["practice_pie_png_b64"], # Done
    'deployedMSLS': len(data["msls"]), # Done
    'totalInteractions': data["n_interactions"], # Done
    'AcademicSettings': academic_count, # Done
    'CommunitySettings': other_count, # Done
    'InsightCount': data["insight_count"], # Done
    'Congresses': data["congresses"], # Done
    'Reporting_Dates':'June-September 2025'
  }
  return stats

