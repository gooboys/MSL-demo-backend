from pptx import Presentation
from pptx.util import Pt, Inches, Emu
from pptx.dml.color import RGBColor
from pptx.shapes.group import GroupShape
from io import BytesIO
import os


# EMU conversions
EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_PT = 12700
EMU_PER_PX = 9525  # assumes 96 DPI

def _set_text_simple(shp, text: str, font="Century Gothic", size=14, color=(40, 36, 111)):
  tf = getattr(shp, "text_frame", None)
  if tf is None:
    raise ValueError(f"no text_frame on {getattr(shp,'name','?')}")
  tf.text = text or ""   # robust for titles/placeholders
  run = tf.paragraphs[0].runs[0]
  if font:  run.font.name = font
  if size:  run.font.size = Pt(size)
  if color: run.font.color.rgb = RGBColor(*color)

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

def _iter_shapes_recursive(container):
  # Yields shapes, including those inside GroupShapes
  for shp in container.shapes:
    yield shp
    if isinstance(shp, GroupShape):
      for inner in _iter_shapes_recursive(shp):
        yield inner

def _find_shape_by_name(slide, name: str):
  # Exact name match anywhere on the slide (recurses into groups)
  for shp in _iter_shapes_recursive(slide):
    if getattr(shp, "name", None) == name:
      return shp
  return None

def hex_to_rgb(hex_code: str) -> tuple[int, int, int]:
  s = hex_code.lstrip('#')
  if len(s) != 6:
    raise ValueError(f"hex must be 6 chars, got '{hex_code}'")
  return tuple(int(s[i:i+2], 16) for i in (0, 2, 4))

# FOR GRAPHS
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

def insert_image_fit_units(
  prs,
  slide_idx: int,
  image_bytes: bytes,
  box_w, box_h,           # size of the bounding box
  pos_x, pos_y,           # top-left position of the box
  units: str = "in"       # 'in', 'cm', 'pt', or 'px'
):
  """
  Place an image (bytes) on slide `slide_idx`, scaled to FIT inside a box of (box_w x box_h)
  whose top-left corner is at (pos_x, pos_y), all in the chosen `units`.
  Returns the picture shape.
  """
  # Convert all to EMU
  max_w_emu = _to_emu_units(box_w, units)
  max_h_emu = _to_emu_units(box_h, units)
  left_emu  = _to_emu_units(pos_x, units)
  top_emu   = _to_emu_units(pos_y, units)

  slide = prs.slides[slide_idx]
  stream = BytesIO(image_bytes)

  # Add picture at the intended anchor, then size it
  pic = slide.shapes.add_picture(stream, left_emu, top_emu)

  # Native size (EMU)
  native_w = pic.width
  native_h = pic.height

  # Compute fit size
  fit_w, fit_h = _fit_size(native_w, native_h, max_w_emu, max_h_emu)

  # Apply size and keep anchored to the same top-left
  pic.width = fit_w
  pic.height = fit_h
  pic.left = left_emu
  pic.top = top_emu

  return pic


def editPPTX(pres, ref: dict[int, dict], items: dict[int, tuple[str, int]]):
  """
  Edits shapes by NAME, not by id.

  Parameters:
    pres  : Presentation
    ref   : dict[int, dict]  -> your existing payloads keyed by numeric id
    items : dict[int, (name:str, slide_idx:int)] -> maps numeric id to (shape_name, slide_index)

  Implementation detail:
    We derive ref_by_name = {shape_name: cfg} from (ref, items) and then edit by name.
  """
  # IDs that should use the simpler text setter (titles / theme headers)
  special_ids = {220, 229, 238, 249, 258, 267, 278, 287, 296}
  # compute their names once so we can branch by name during the slide loop
  special_names = { items[i][0] for i in special_ids if i in items }

  # Build slide_index -> [shape_names]
  slide_to_names: dict[int, list[str]] = {}
  for shape_id, (shape_name, slide_idx) in items.items():
    slide_to_names.setdefault(slide_idx, []).append(shape_name)

  # Derive name-keyed config so you don't have to rewrite your 'ref'
  ref_by_name: dict[str, dict] = {}
  for shape_id, cfg in ref.items():
    info = items.get(shape_id)
    if not info:
      continue
    shape_name, _ = info
    ref_by_name[shape_name] = cfg

  for slide_idx, names in slide_to_names.items():
    if slide_idx < 0 or slide_idx >= len(pres.slides):
      continue
    slide = pres.slides[slide_idx]
    for name in names:
      cfg = ref_by_name.get(name)
      if not cfg:
        continue
      shp = _find_shape_by_name(slide, name)
      if shp is None:
        continue
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
      except ValueError:
        continue

def true_replacement(stats, patient, education, competitive, single):
  base_dir = os.path.dirname(os.path.abspath(__file__))
  template_path = os.path.join(base_dir, "New Acquis Template.pptx")
  prs = Presentation(template_path)

  def safe_get(lst, idx, key, default=""):
    return (lst[idx].get(key) if 0 <= idx < len(lst) else default)

  # Category order (unused below, but kept for context)
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
  catcount = [stats['category_count'].get(key, 0) for key in order]

  # Your existing ref stays keyed by numeric ids; items carries the shape names
  ref = {
    146: {"text": safe_get(single, 0, "Raw CRM Input (from MSL)"),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f"), "bold": True},
    148: {"text": safe_get(single, 0, "idea"),
          "font": "Century Gothic", "font_size": 6, "font_color": hex_to_rgb("28246f"), "bold": True},
    149: {"text": safe_get(single, 1, "idea"),
          "font": "Century Gothic", "font_size": 6, "font_color": hex_to_rgb("28246f"), "bold": True},
    150: {"text": safe_get(single, 3, "idea"),
          "font": "Century Gothic", "font_size": 6, "font_color": hex_to_rgb("28246f"), "bold": True},
    151: {"text": safe_get(single, 4, "idea"),
          "font": "Century Gothic", "font_size": 6, "font_color": hex_to_rgb("28246f"), "bold": True},
    152: {"text": safe_get(single, 2, "idea"),
          "font": "Century Gothic", "font_size": 6, "font_color": hex_to_rgb("28246f"), "bold": True},
    159: {"text": safe_get(single, 0, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 6, "font_color": hex_to_rgb("28246f")},
    160: {"text": safe_get(single, 1, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 6, "font_color": hex_to_rgb("28246f")},
    161: {"text": safe_get(single, 3, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 6, "font_color": hex_to_rgb("28246f")},
    162: {"text": safe_get(single, 4, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 6, "font_color": hex_to_rgb("28246f")},
    163: {"text": safe_get(single, 2, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 6, "font_color": hex_to_rgb("28246f")},
    165: {"text": single[0]["categories"].replace("[", "").replace("]", ""),
          "font": "Century Gothic", "font_size": 10, "font_color": (255, 255, 255)},
    166: {"text": single[1]["categories"].replace("[", "").replace("]", ""),
          "font": "Century Gothic", "font_size": 10, "font_color": (255, 255, 255)},
    169: {"text": safe_get(single, 0, "categorization_rationale"),
          "font": "Century Gothic", "font_size": 6, "font_color": hex_to_rgb("28246f")},
    170: {"text": safe_get(single, 1, "categorization_rationale"),
          "font": "Century Gothic", "font_size": 6, "font_color": hex_to_rgb("28246f")},
    172: {"text": safe_get(single, 2, "categorization_rationale"),
          "font": "Century Gothic", "font_size": 6, "font_color": hex_to_rgb("28246f")},
    # END OF FIRST SLIDE
    187: {"text": stats["Reporting_Dates"],
          "font": "Century Gothic", "font_size": 16, "font_color": hex_to_rgb("28246f")},
    191: {"text": str(stats["deployedMSLS"]),
          "font": "Century Gothic", "font_size": 16, "font_color": hex_to_rgb("28246f")},
    193: {"text": "Total: " + str(stats["totalInteractions"]),
          "font": "Century Gothic", "font_size": 16, "font_color": hex_to_rgb("28246f")},
    189: {"text": "\n".join(stats['Congresses']),
          "font": "Century Gothic", "font_size": 16, "font_color": hex_to_rgb("28246f")},
    195: {"text": str(stats['InsightCount']),
          "font": "Century Gothic", "font_size": 16, "font_color": hex_to_rgb("28246f")},
    197: {"text": str(catcount[0]),
          "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    199: {"text": str(catcount[1]),
          "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    201: {"text": str(catcount[2]),
      "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    203: {"text": str(catcount[3]),
          "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    205: {"text": str(catcount[4]),
          "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    207: {"text": str(catcount[5]),
          "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    209: {"text": str(catcount[6]),
          "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    211: {"text": str(catcount[7]),
          "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    213: {"text": str(catcount[8]),
          "font": "Century Gothic", "font_size": 11, "font_color": hex_to_rgb("325fa7")},
    # END SLIDE 2, INPUT GRAPHS LATER
    220: {"text": "Theme 1 (n="+str(len(patient[0]["other_sources"])+3)+")",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    229: {"text": "Theme 2 (n="+str(len(patient[1]["other_sources"])+3)+")",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    238: {"text": "Theme 3 (n="+str(len(patient[2]["other_sources"])+3)+")",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    222: {"text": patient[0]["gap_definition"],
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    231: {"text": patient[1]["gap_definition"],
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    240: {"text": patient[2]["gap_definition"],
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    224: {"text": "\n".join([f"id {q['id']}: '{q['quote']}'" for q in patient[0]["representative_quotes"]]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    233: {"text": "\n".join([f"id {q['id']}: '{q['quote']}'" for q in patient[1]["representative_quotes"]]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    242: {"text": "\n".join([f"id {q['id']}: '{q['quote']}'" for q in patient[2]["representative_quotes"]]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    226: {"text": "1: "+patient[0]["root_cause_questions"][0]+"\n"+"2: "+patient[0]["root_cause_questions"][1],
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    235: {"text": "1: "+patient[1]["root_cause_questions"][0]+"\n"+"2: "+patient[1]["root_cause_questions"][1],
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    244: {"text": "1: "+patient[2]["root_cause_questions"][0]+"\n"+"2: "+patient[2]["root_cause_questions"][1],
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    # END SLIDE 3
    249: {"text": "Theme 1 (n="+str(len(education[0]["other_sources"])+3)+")",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    258: {"text": "Theme 2 (n="+str(len(education[1]["other_sources"])+3)+")",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    267: {"text": "Theme 3 (n="+str(len(education[2]["other_sources"])+3)+")",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    251: {"text": education[0]["gap_definition"],
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    260: {"text": education[1]["gap_definition"],
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    269: {"text": education[2]["gap_definition"],
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    253: {"text": "\n".join([f"id {q['id']}: '{q['quote']}'" for q in education[0]["representative_quotes"]]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    262: {"text": "\n".join([f"id {q['id']}: '{q['quote']}'" for q in education[1]["representative_quotes"]]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    271: {"text": "\n".join([f"id {q['id']}: '{q['quote']}'" for q in education[2]["representative_quotes"]]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    255: {"text": "1: "+education[0]["root_cause_questions"][0]+"\n"+"2: "+education[0]["root_cause_questions"][1],
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    264: {"text": "1: "+education[1]["root_cause_questions"][0]+"\n"+"2: "+education[1]["root_cause_questions"][1],
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    273: {"text": "1: "+education[2]["root_cause_questions"][0]+"\n"+"2: "+education[2]["root_cause_questions"][1],
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    # END SLIDE 4
    278: {"text": "Theme 1 (n="+str(len(competitive[0]["other_sources"])+3)+")",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    287: {"text": "Theme 2 (n="+str(len(competitive[1]["other_sources"])+3)+")",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    296: {"text": "Theme 3 (n="+str(len(competitive[2]["other_sources"])+3)+")",
          "font": "Century Gothic", "font_size": 14, "font_color": hex_to_rgb("28246f")},
    280: {"text": competitive[0]["gap_definition"],
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    289: {"text": competitive[1]["gap_definition"],
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    298: {"text": competitive[2]["gap_definition"],
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    282: {"text": "\n".join([f"id {q['id']}: '{q['quote']}'" for q in competitive[0]["representative_quotes"]]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    291: {"text": "\n".join([f"id {q['id']}: '{q['quote']}'" for q in competitive[1]["representative_quotes"]]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    300: {"text": "\n".join([f"id {q['id']}: '{q['quote']}'" for q in competitive[2]["representative_quotes"]]),
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    284: {"text": "1: "+competitive[0]["root_cause_questions"][0]+"\n"+"2: "+competitive[0]["root_cause_questions"][1],
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    293: {"text": "1: "+competitive[1]["root_cause_questions"][0]+"\n"+"2: "+competitive[1]["root_cause_questions"][1],
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")},
    302: {"text": "1: "+competitive[2]["root_cause_questions"][0]+"\n"+"2: "+competitive[2]["root_cause_questions"][1],
          "font": "Century Gothic", "font_size": 8, "font_color": hex_to_rgb("28246f")}
    # END SLIDES (5)
  }

  # items: id -> (shape_name, slide_idx). We will edit by shape_name.
  items = {146: ('Google Shape;518;p44', 5), 148: ('Google Shape;520;p44', 5), 149: ('Google Shape;521;p44', 5), 150: ('Google Shape;522;p44', 5), 151: ('Google Shape;523;p44', 5), 152: ('Google Shape;524;p44', 5), 153: ('Google Shape;530;p44', 5), 154: ('Google Shape;531;p44', 5), 155: ('Google Shape;532;p44', 5), 156: ('Google Shape;533;p44', 5), 157: ('Google Shape;534;p44', 5), 159: ('Google Shape;536;p44', 5), 160: ('Google Shape;537;p44', 5), 161: ('Google Shape;538;p44', 5), 162: ('Google Shape;539;p44', 5), 163: ('Google Shape;540;p44', 5), 165: ('Google Shape;542;p44', 5), 166: ('Google Shape;543;p44', 5), 167: ('Google Shape;544;p44', 5), 168: ('Google Shape;545;p44', 5), 169: ('Google Shape;546;p44', 5), 170: ('Google Shape;547;p44', 5), 171: ('Google Shape;548;p44', 5), 172: ('Google Shape;549;p44', 5), 185: ('Google Shape;581;p45', 6), 186: ('Google Shape;582;p45', 6), 187: ('Google Shape;584;p45', 6), 189: ('Google Shape;587;p45', 6), 191: ('Google Shape;590;p45', 6), 193: ('Google Shape;593;p45', 6), 195: ('Google Shape;596;p45', 6), 197: ('Google Shape;599;p45', 6), 198: ('Google Shape;600;p45', 6), 199: ('Google Shape;602;p45', 6), 200: ('Google Shape;603;p45', 6), 201: ('Google Shape;605;p45', 6), 202: ('Google Shape;606;p45', 6), 203: ('Google Shape;608;p45', 6), 204: ('Google Shape;609;p45', 6), 205: ('Google Shape;611;p45', 6), 206: ('Google Shape;612;p45', 6), 207: ('Google Shape;614;p45', 6), 208: ('Google Shape;615;p45', 6), 209: ('Google Shape;617;p45', 6), 210: ('Google Shape;618;p45', 6), 211: ('Google Shape;620;p45', 6), 212: ('Google Shape;621;p45', 6), 213: ('Google Shape;623;p45', 6), 220: ('Google Shape;635;p46', 7), 222: ('Google Shape;637;p46', 7), 224: ('Google Shape;639;p46', 7), 226: ('Google Shape;641;p46', 7), 229: ('Google Shape;644;p46', 7), 231: ('Google Shape;646;p46', 7), 233: ('Google Shape;648;p46', 7), 235: ('Google Shape;650;p46', 7), 238: ('Google Shape;653;p46', 7), 240: ('Google Shape;655;p46', 7), 242: ('Google Shape;657;p46', 7), 244: ('Google Shape;659;p46', 7), 249: ('Google Shape;669;p47', 8), 250: ('Google Shape;670;p47', 8), 251: ('Google Shape;671;p47', 8), 253: ('Google Shape;673;p47', 8), 255: ('Google Shape;675;p47', 8), 258: ('Google Shape;678;p47', 8), 260: ('Google Shape;680;p47', 8), 262: ('Google Shape;682;p47', 8), 264: ('Google Shape;684;p47', 8), 267: ('Google Shape;687;p47', 8), 269: ('Google Shape;689;p47', 8), 271: ('Google Shape;691;p47', 8), 273: ('Google Shape;693;p47', 8), 278: ('Google Shape;703;p48', 9), 280: ('Google Shape;705;p48', 9), 282: ('Google Shape;707;p48', 9), 284: ('Google Shape;709;p48', 9), 287: ('Google Shape;712;p48', 9), 289: ('Google Shape;714;p48', 9), 291: ('Google Shape;716;p48', 9), 293: ('Google Shape;718;p48', 9), 296: ('Google Shape;721;p48', 9), 298: ('Google Shape;723;p48', 9), 300: ('Google Shape;725;p48', 9), 302: ('Google Shape;727;p48', 9)}

  # Adding Graphs
  insert_image_fit_units(
    prs,
    slide_idx=3,
    image_bytes=stats['graph1'],
    box_w=4.5,
    box_h=3,           # size of the bounding box
    pos_x=3.65, 
    pos_y=1.5,           # top-left position of the box
    units="in"       # 'in', 'cm', 'pt', or 'px'
  )

  insert_image_fit_units(
    prs,
    slide_idx=3,
    image_bytes=stats['graph2'],
    box_w=4.5,
    box_h=3,           # size of the bounding box
    pos_x=8.5,
    pos_y=1.5,           # top-left position of the box
    units="in"       # 'in', 'cm', 'pt', or 'px'
  )

  editPPTX(prs, ref, items)
  buf = BytesIO()
  prs.save(buf)
  buf.seek(0)
  return buf.getvalue()
