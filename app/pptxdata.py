from pptx import Presentation
from pptx.util import Pt, Inches, Emu
from pptx.dml.color import RGBColor
from pptx.shapes.group import GroupShape
from io import BytesIO
import os
from typing import Dict, Any, Tuple, List, Set

# ======================
# EMU conversions
# ======================
EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_PT = 12700
EMU_PER_PX = 9525  # assumes 96 DPI

def _to_emu_units(val, units="in"):
  if isinstance(val, (Inches, Pt, Emu)):
    return int(val)
  if isinstance(val, (int, float)):
    u = units.lower()
    if u in ("in", "inch", "inches"):
      return int(val * EMU_PER_INCH)
    if u in ("cm",):
      return int(val * EMU_PER_CM)
    if u in ("pt", "point", "points"):
      return int(val * EMU_PER_PT)
    if u in ("px", "pixel", "pixels"):
      return int(val * EMU_PER_PX)
  raise TypeError("Position/size must be a number with units in {'in','cm','pt','px'} or a pptx unit (Inches/Pt/Emu).")

def _fit_size(nw_emu, nh_emu, max_w_emu, max_h_emu):
  if nw_emu <= 0 or nh_emu <= 0:
    return max_w_emu, max_h_emu
  r = min(max_w_emu / float(nw_emu), max_h_emu / float(nh_emu))
  return int(nw_emu * r), int(nh_emu * r)

# ======================
# Color & text utils
# ======================
def hex_to_rgb(hex_code: str) -> tuple[int, int, int]:
  s = hex_code.lstrip('#')
  if len(s) != 6:
    raise ValueError(f"hex must be 6 chars, got '{hex_code}'")
  return tuple(int(s[i:i+2], 16) for i in (0, 2, 4))

def _textify_categories(val) -> str:
  if val is None:
    return ""
  if isinstance(val, (list, tuple, set)):
    return ", ".join(map(str, val))
  return str(val)

# ======================
# Shape discovery
# ======================
def _iter_shapes_recursive(container):
  for shp in container.shapes:
    yield shp
    if isinstance(shp, GroupShape):
      for inner in _iter_shapes_recursive(shp):
        yield inner

def _index_shapes_by_name(prs: Presentation):
  """
  Returns:
    by_name: Dict[str, Tuple[int, object]]  -> first occurrence of name
    dups:    Dict[str, List[Tuple[int, object]]] -> all occurrences (when >1)
  """
  by_name: Dict[str, Tuple[int, Any]] = {}
  dups: Dict[str, List[Tuple[int, Any]]] = {}
  for si, slide in enumerate(prs.slides):
    for shp in _iter_shapes_recursive(slide):
      name = getattr(shp, "name", None)
      if not name:
        continue
      if name in by_name:
        dups.setdefault(name, [by_name[name]]).append((si, shp))
      else:
        by_name[name] = (si, shp)
  return by_name, dups

def dump_shape_map(prs: Presentation):
  """
  Debug helper: prints slide -> shape names
  """
  for si, slide in enumerate(prs.slides):
    print(f"[pptx] slide {si}:")
    for shp in _iter_shapes_recursive(slide):
      nm = getattr(shp, "name", None)
      tp = getattr(getattr(shp, "shape_type", None), "name", "?")
      txt = getattr(getattr(shp, "text_frame", None), "text", "")
      if txt:
        txt = txt.replace("\r", " ").replace("\n", " ")
        if len(txt) > 48:
          txt = txt[:45] + "..."
      print(f"  - type={tp:<12} name='{nm}' text='{txt}'")

# ======================
# Text setters (robust)
# ======================
def _set_text_simple(shp, text: str, font="Century Gothic", size=14, color=(40, 36, 111)):
  tf = getattr(shp, "text_frame", None)
  if tf is None:
    raise ValueError(f"no text_frame on {getattr(shp,'name','?')}")
  tf.text = text or ""
  p0 = tf.paragraphs[0]
  if not p0.runs:
    run = p0.add_run()
    run.text = text or ""
  run = p0.runs[0]
  if font:
    run.font.name = font
  if size:
    run.font.size = Pt(size)
  if color:
    run.font.color.rgb = RGBColor(*color)

def _overwrite_shape_text(
  shp,
  text: str,
  font_name: str = "Calibri",
  font_size: int | float = 20,
  font_color: tuple[int, int, int] = (0, 0, 0),
  bold: bool | None = None,
  italic: bool | None = None,
):
  tf = getattr(shp, "text_frame", None)
  if tf is None:
    raise ValueError(f"Shape name={getattr(shp, 'name', '?')} has no text frame")
  tf.clear()
  p = tf.paragraphs[0]
  run = p.add_run()
  run.text = text or ""
  if font_name:
    run.font.name = font_name
  if font_size:
    run.font.size = Pt(font_size)
  if font_color:
    run.font.color.rgb = RGBColor(*font_color)
  if bold is not None:
    run.font.bold = bold
  if italic is not None:
    run.font.italic = italic

# ======================
# Image placement
# ======================
def insert_image_fit_units(
  prs,
  slide_idx: int,
  image_bytes: bytes,
  box_w, box_h,
  pos_x, pos_y,
  units: str = "in"
):
  """
  Place an image (bytes) on slide `slide_idx`, scaled to FIT inside a box of (box_w x box_h)
  whose top-left corner is at (pos_x, pos_y), all in the chosen `units`.
  Returns the picture shape.
  """
  max_w_emu = _to_emu_units(box_w, units)
  max_h_emu = _to_emu_units(box_h, units)
  left_emu  = _to_emu_units(pos_x, units)
  top_emu   = _to_emu_units(pos_y, units)

  slide = prs.slides[slide_idx]
  stream = BytesIO(image_bytes)
  pic = slide.shapes.add_picture(stream, left_emu, top_emu)

  native_w = pic.width
  native_h = pic.height
  fit_w, fit_h = _fit_size(native_w, native_h, max_w_emu, max_h_emu)

  pic.width = fit_w
  pic.height = fit_h
  pic.left = left_emu
  pic.top = top_emu
  return pic

# ======================
# Core editor (name-first, slide-agnostic)
# ======================
def editPPTX(pres, ref: Dict[int, Dict], items: Dict[int, Tuple[str, int]], debug: bool = False):
  """
  Edits shapes by NAME; slide indices in `items` are hints only.
  Survives slide reorders and most template edits as long as names remain stable.

  Args:
    pres  : Presentation
    ref   : dict[int, dict]  -> per-id config (contains text/font/etc.)
    items : dict[int, (shape_name, slide_idx_hint)]
    debug : print duplicate & missing name diagnostics
  """
  # headings using simple setter
  special_ids: Set[int] = {220, 229, 238, 249, 258, 267, 278, 287, 296}
  special_names = { items[i][0] for i in special_ids if i in items }

  # name-keyed config from (ref, items)
  ref_by_name: Dict[str, Dict] = {}
  for shape_id, cfg in ref.items():
    info = items.get(shape_id)
    if not info:
      continue
    shape_name, _ = info
    # normalize categories fields to text when present
    if "categories" in cfg and "text" not in cfg:
      cfg["text"] = _textify_categories(cfg["categories"])
    ref_by_name[shape_name] = cfg

  by_name, dups = _index_shapes_by_name(pres)

  if debug and dups:
    print("[pptx] Duplicate shape names detected (will use the first occurrence):")
    for nm, locs in dups.items():
      print(f"  - {nm}: " + ", ".join([f"slide {si}" for si, _ in locs]))

  missing: List[str] = []
  for name, cfg in ref_by_name.items():
    entry = by_name.get(name)
    if entry is None:
      missing.append(name)
      continue
    slide_idx, shp = entry
    try:
      if name in special_names:
        _set_text_simple(
          shp,
          text=cfg.get("text", ""),
          font=cfg.get("font", "Calibri"),
          size=cfg.get("font_size", 12),
          color=cfg.get("font_color", (0, 0, 0)),
        )
      else:
        _overwrite_shape_text(
          shp,
          text=cfg.get("text", ""),
          font_name=cfg.get("font", "Calibri"),
          font_size=cfg.get("font_size", 12),
          font_color=cfg.get("font_color", (0, 0, 0)),
          bold=cfg.get("bold", None),
          italic=cfg.get("italic", None),
        )
    except Exception as e:
      if debug:
        print(f"[pptx] Failed writing '{name}' on slide {slide_idx}: {e}")

  if debug and missing:
    print("[pptx] Names not found in current template (removed/renamed?):")
    for nm in missing:
      print(f"  - {nm}")

# ======================
# Public entry
# ======================
def true_replacement(
  stats: Dict[str, Any],
  patient: List[Dict[str, Any]],
  education: List[Dict[str, Any]],
  competitive: List[Dict[str, Any]],
  single: List[Dict[str, Any]],
  template_path: str | None = None,
  debug: bool = False
) -> bytes:
  """
  Loads the template, injects text+graphs, returns PPTX bytes.
  """
  base_dir = os.path.dirname(os.path.abspath(__file__))
  template_path = template_path or os.path.join(base_dir, "New Acquis Template.pptx")

  if not os.path.exists(template_path):
    raise FileNotFoundError(f"Template not found at: {template_path}")

  prs = Presentation(template_path)

  def safe_get(lst, idx, key, default=""):
    return (lst[idx].get(key) if 0 <= idx < len(lst) and isinstance(lst[idx], dict) else default)

  # Category order (kept for reference)
  order = [
    "Access Insights",
    "Patient Management / Care Insights",
    "Clinical Development Insights",
    "Competitive Insights",
    "Product Insights (Drug Science)",
    "Education",
    "Logistics",
    "Adverse Event (AE) Insights",
    "Other"
  ]
  catcount = [stats.get('category_count', {}).get(key, 0) for key in order]

  # Build ref (id -> cfg). NOTE: categories fields normalized to text via editor.
  ref: Dict[int, Dict[str, Any]] = {
    146: {"text": safe_get(single, 0, "Raw CRM Input (from MSL)"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f"), "bold": True},
    148: {"text": safe_get(single, 0, "idea"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f"), "bold": True},
    149: {"text": safe_get(single, 1, "idea"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f"), "bold": True},
    150: {"text": safe_get(single, 3, "idea"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f"), "bold": True},
    151: {"text": safe_get(single, 4, "idea"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f"), "bold": True},
    152: {"text": safe_get(single, 2, "idea"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f"), "bold": True},

    159: {"text": safe_get(single, 0, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    160: {"text": safe_get(single, 1, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    161: {"text": safe_get(single, 3, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    162: {"text": safe_get(single, 4, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    163: {"text": safe_get(single, 2, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},

    # categories: accept list or str
    165: {"text": _textify_categories(safe_get(single, 0, "categories")),
          "font": "Century Gothic", "font_size": 10, "font_color": (255, 255, 255)},
    166: {"text": _textify_categories(safe_get(single, 1, "categories")),
          "font": "Century Gothic", "font_size": 10, "font_color": (255, 255, 255)},

    169: {"text": safe_get(single, 0, "categorization_rationale"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    170: {"text": safe_get(single, 1, "categorization_rationale"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    172: {"text": safe_get(single, 2, "categorization_rationale"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},

    # Slide 2
    187: {"text": stats.get("Reporting_Dates", ""),
          "font": "Century Gothic", "font_size": 16, "font_color": hex_to_rgb("28246f")},
    191: {"text": str(stats.get("deployedMSLS", "")),
          "font": "Century Gothic", "font_size": 16, "font_color": hex_to_rgb("28246f")},
    193: {"text": "Total: " + str(stats.get("totalInteractions", "")),
          "font": "Century Gothic", "font_size": 16, "font_color": hex_to_rgb("28246f")},
    189: {"text": "\n".join(stats.get('Congresses', [])),
          "font": "Century Gothic", "font_size": 16, "font_color": hex_to_rgb("28246f")},
    195: {"text": str(stats.get('InsightCount', "")),
          "font": "Century Gothic", "font_size": 16, "font_color": hex_to_rgb("28246f")},

    197: {"text": str(catcount[0]), "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    199: {"text": str(catcount[1]), "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    201: {"text": str(catcount[2]), "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    203: {"text": str(catcount[3]), "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    205: {"text": str(catcount[4]), "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    207: {"text": str(catcount[5]), "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    209: {"text": str(catcount[6]), "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    211: {"text": str(catcount[7]), "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    213: {"text": str(catcount[8]), "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},

    # Slide 3: Patient
    220: {"text": f"Theme 1 (n={len(patient[0].get('other_sources', [])) + 3})",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    229: {"text": f"Theme 2 (n={len(patient[1].get('other_sources', [])) + 3})",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    238: {"text": f"Theme 3 (n={len(patient[2].get('other_sources', [])) + 3})",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},

    222: {"text": patient[0].get("gap_definition", ""),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    231: {"text": patient[1].get("gap_definition", ""),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    240: {"text": patient[2].get("gap_definition", ""),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},

    224: {"text": "\n".join([f"id {q.get('id')}: '{q.get('quote','')}'" for q in patient[0].get("representative_quotes", [])]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    233: {"text": "\n".join([f"id {q.get('id')}: '{q.get('quote','')}'" for q in patient[1].get("representative_quotes", [])]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    242: {"text": "\n".join([f"id {q.get('id')}: '{q.get('quote','')}'" for q in patient[2].get("representative_quotes", [])]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},

    226: {"text": "1: " + (patient[0].get("root_cause_questions", ["",""])[0] or "") + "\n" +
                  "2: " + (patient[0].get("root_cause_questions", ["",""])[1] or ""),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    235: {"text": "1: " + (patient[1].get("root_cause_questions", ["",""])[0] or "") + "\n" +
                  "2: " + (patient[1].get("root_cause_questions", ["",""])[1] or ""),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    244: {"text": "1: " + (patient[2].get("root_cause_questions", ["",""])[0] or "") + "\n" +
                  "2: " + (patient[2].get("root_cause_questions", ["",""])[1] or ""),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},

    # Slide 4: Education
    249: {"text": f"Theme 1 (n={len(education[0].get('other_sources', [])) + 3})",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    258: {"text": f"Theme 2 (n={len(education[1].get('other_sources', [])) + 3})",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    267: {"text": f"Theme 3 (n={len(education[2].get('other_sources', [])) + 3})",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},

    251: {"text": education[0].get("gap_definition", ""),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    260: {"text": education[1].get("gap_definition", ""),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    269: {"text": education[2].get("gap_definition", ""),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},

    253: {"text": "\n".join([f"id {q.get('id')}: '{q.get('quote','')}'" for q in education[0].get("representative_quotes", [])]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    262: {"text": "\n".join([f"id {q.get('id')}: '{q.get('quote','')}'" for q in education[1].get("representative_quotes", [])]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    271: {"text": "\n".join([f"id {q.get('id')}: '{q.get('quote','')}'" for q in education[2].get("representative_quotes", [])]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},

    255: {"text": "1: " + (education[0].get("root_cause_questions", ["",""])[0] or "") + "\n" +
                  "2: " + (education[0].get("root_cause_questions", ["",""])[1] or ""),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    264: {"text": "1: " + (education[1].get("root_cause_questions", ["",""])[0] or "") + "\n" +
                  "2: " + (education[1].get("root_cause_questions", ["",""])[1] or ""),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    273: {"text": "1: " + (education[2].get("root_cause_questions", ["",""])[0] or "") + "\n" +
                  "2: " + (education[2].get("root_cause_questions", ["",""])[1] or ""),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},

    # Slide 5: Competitive
    278: {"text": f"Theme 1 (n={len(competitive[0].get('other_sources', [])) + 3})",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    287: {"text": f"Theme 2 (n={len(competitive[1].get('other_sources', [])) + 3})",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    296: {"text": f"Theme 3 (n={len(competitive[2].get('other_sources', [])) + 3})",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},

    280: {"text": competitive[0].get("gap_definition", ""),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    289: {"text": competitive[1].get("gap_definition", ""),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    298: {"text": competitive[2].get("gap_definition", ""),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},

    282: {"text": "\n".join([f"id {q.get('id')}: '{q.get('quote','')}'" for q in competitive[0].get("representative_quotes", [])]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    291: {"text": "\n".join([f"id {q.get('id')}: '{q.get('quote','')}'" for q in competitive[1].get("representative_quotes", [])]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    300: {"text": "\n".join([f"id {q.get('id')}: '{q.get('quote','')}'" for q in competitive[2].get("representative_quotes", [])]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},

    284: {"text": "1: " + (competitive[0].get("root_cause_questions", ["",""])[0] or "") + "\n" +
                  "2: " + (competitive[0].get("root_cause_questions", ["",""])[1] or ""),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    293: {"text": "1: " + (competitive[1].get("root_cause_questions", ["",""])[0] or "") + "\n" +
                  "2: " + (competitive[1].get("root_cause_questions", ["",""])[1] or ""),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    302: {"text": "1: " + (competitive[2].get("root_cause_questions", ["",""])[0] or "") + "\n" +
                  "2: " + (competitive[2].get("root_cause_questions", ["",""])[1] or ""),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
  }

  # items: id -> (shape_name, slide_idx_hint).
  items: Dict[int, Tuple[str, int]] = {
    146: ('Google Shape;518;p44', 5), 148: ('Google Shape;520;p44', 5), 149: ('Google Shape;521;p44', 5),
    150: ('Google Shape;522;p44', 5), 151: ('Google Shape;523;p44', 5), 152: ('Google Shape;524;p44', 5),
    153: ('Google Shape;530;p44', 5), 154: ('Google Shape;531;p44', 5), 155: ('Google Shape;532;p44', 5),
    156: ('Google Shape;533;p44', 5), 157: ('Google Shape;534;p44', 5), 159: ('Google Shape;536;p44', 5),
    160: ('Google Shape;537;p44', 5), 161: ('Google Shape;538;p44', 5), 162: ('Google Shape;539;p44', 5),
    163: ('Google Shape;540;p44', 5), 165: ('Google Shape;542;p44', 5), 166: ('Google Shape;543;p44', 5),
    167: ('Google Shape;544;p44', 5), 168: ('Google Shape;545;p44', 5), 169: ('Google Shape;546;p44', 5),
    170: ('Google Shape;547;p44', 5), 171: ('Google Shape;548;p44', 5), 172: ('Google Shape;549;p44', 5),
    185: ('Google Shape;581;p45', 6), 186: ('Google Shape;582;p45', 6), 187: ('Google Shape;584;p45', 6),
    189: ('Google Shape;587;p45', 6), 191: ('Google Shape;590;p45', 6), 193: ('Google Shape;593;p45', 6),
    195: ('Google Shape;596;p45', 6), 197: ('Google Shape;599;p45', 6), 198: ('Google Shape;600;p45', 6),
    199: ('Google Shape;602;p45', 6), 200: ('Google Shape;603;p45', 6), 201: ('Google Shape;605;p45', 6),
    202: ('Google Shape;606;p45', 6), 203: ('Google Shape;608;p45', 6), 204: ('Google Shape;609;p45', 6),
    205: ('Google Shape;611;p45', 6), 206: ('Google Shape;612;p45', 6), 207: ('Google Shape;614;p45', 6),
    208: ('Google Shape;615;p45', 6), 209: ('Google Shape;617;p45', 6), 210: ('Google Shape;618;p45', 6),
    211: ('Google Shape;620;p45', 6), 212: ('Google Shape;621;p45', 6), 213: ('Google Shape;623;p45', 6),
    220: ('Google Shape;635;p46', 7), 222: ('Google Shape;637;p46', 7), 224: ('Google Shape;639;p46', 7),
    226: ('Google Shape;641;p46', 7), 229: ('Google Shape;644;p46', 7), 231: ('Google Shape;646;p46', 7),
    233: ('Google Shape;648;p46', 7), 235: ('Google Shape;650;p46', 7), 238: ('Google Shape;653;p46', 7),
    240: ('Google Shape;655;p46', 7), 242: ('Google Shape;657;p46', 7), 244: ('Google Shape;659;p46', 7),
    249: ('Google Shape;669;p47', 8), 250: ('Google Shape;670;p47', 8), 251: ('Google Shape;671;p47', 8),
    253: ('Google Shape;673;p47', 8), 255: ('Google Shape;675;p47', 8), 258: ('Google Shape;678;p47', 8),
    260: ('Google Shape;680;p47', 8), 262: ('Google Shape;682;p47', 8), 264: ('Google Shape;684;p47', 8),
    267: ('Google Shape;687;p47', 8), 269: ('Google Shape;689;p47', 8), 271: ('Google Shape;691;p47', 8),
    273: ('Google Shape;693;p47', 8), 278: ('Google Shape;703;p48', 9), 280: ('Google Shape;705;p48', 9),
    282: ('Google Shape;707;p48', 9), 284: ('Google Shape;709;p48', 9), 287: ('Google Shape;712;p48', 9),
    289: ('Google Shape;714;p48', 9), 291: ('Google Shape;716;p48', 9), 293: ('Google Shape;718;p48', 9),
    296: ('Google Shape;721;p48', 9), 298: ('Google Shape;723;p48', 9), 300: ('Google Shape;725;p48', 9),
    302: ('Google Shape;727;p48', 9)
  }

  # Add graphs (confirm indices if your slide order changed)
  # If slide order is volatile, consider naming a placeholder shape and deriving its slide via _index_shapes_by_name.
  insert_image_fit_units(
    prs,
    slide_idx=6,
    image_bytes=stats['graph1'],
    box_w=4.3,
    box_h=3,
    pos_x=3.65,
    pos_y=1.9,
    units="in"
  )

  insert_image_fit_units(
    prs,
    slide_idx=6,
    image_bytes=stats['graph2'],
    box_w=4.3,
    box_h=3,
    pos_x=8.45,
    pos_y=1.9,
    units="in"
  )

  # Populate text by name, regardless of slide moves
  editPPTX(prs, ref, items, debug=debug)

  buf = BytesIO()
  prs.save(buf)
  buf.seek(0)
  return buf.getvalue()
