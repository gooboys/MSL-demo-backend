from pptx import Presentation
from pptx.util import Pt, Inches, Emu
from pptx.dml.color import RGBColor
from io import BytesIO
import os

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
    raise ValueError(f"Shape id={getattr(shp, 'shape_id', '?')} has no text frame")
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

def _find_shape_by_id(slide, shape_id: int):
  for s in slide.shapes:
    if getattr(s, "shape_id", None) == shape_id:
      return s
  return None

def hex_to_rgb(hex_code: str) -> tuple[int, int, int]:
  s = hex_code.lstrip('#')
  if len(s) != 6:
    raise ValueError(f"hex must be 6 chars, got '{hex_code}'")
  return tuple(int(s[i:i+2], 16) for i in (0, 2, 4))

def editPPTX(pres, ref: dict[int, dict], items: dict[int, tuple[str, int]]):
  # Build slide_index -> [shape_ids]
  slidetoid: dict[int, list[int]] = {}
  for shape_id, (_, slide_idx) in items.items():
    slidetoid.setdefault(slide_idx, []).append(shape_id)

  for slide_idx, shape_ids in slidetoid.items():
    if slide_idx < 0 or slide_idx >= len(pres.slides):
      continue
    slide = pres.slides[slide_idx]
    for shape_id in shape_ids:
      cfg = ref.get(shape_id)
      if not cfg:
        continue
      shp = _find_shape_by_id(slide, shape_id)
      if shp is None:
        continue
      try:
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
  template_path = os.path.join(base_dir, "Acquis Template.pptx")
  prs = Presentation(template_path)

  def safe_get(lst, idx, key, default=""):
    return (lst[idx].get(key) if 0 <= idx < len(lst) else default)

  ref = {
    146: {"text": safe_get(single, 0, "Raw CRM Input (from MSL)"),
          "font": "Century Gothic", "font_size": 12, "font_color": hex_to_rgb("28246f")},
    148: {"text": safe_get(single, 0, "idea"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    149: {"text": safe_get(single, 1, "idea"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    150: {"text": safe_get(single, 4, "idea"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    151: {"text": safe_get(single, 5, "idea"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    152: {"text": safe_get(single, 2, "idea"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    159: {"text": safe_get(single, 0, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    160: {"text": safe_get(single, 1, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    161: {"text": safe_get(single, 4, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    162: {"text": safe_get(single, 5, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    163: {"text": safe_get(single, 2, "value_classification_rationale"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    165: {"text": (safe_get(single, 0, "categories", []) or [""])[0],
          "font": "Century Gothic", "font_size": 10, "font_color": (255, 255, 255)},
    166: {"text": (safe_get(single, 1, "categories", []) or [""])[0],
          "font": "Century Gothic", "font_size": 10, "font_color": (255, 255, 255)},
    169: {"text": safe_get(single, 0, "categorization_rationale"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    170: {"text": safe_get(single, 1, "categorization_rationale"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    172: {"text": safe_get(single, 2, "categorization_rationale"),
          "font": "Century Gothic", "font_size": 10, "font_color": hex_to_rgb("28246f")},
    # END OF FIRST SLIDE
    187: {"text": stats[],
          "font": "Century Gothic", "font_size": 16, "font_color": hex_to_rgb("28246f")},
  }
  items = {146: ('Google Shape;518;p44', 5), 148: ('Google Shape;520;p44', 5), 149: ('Google Shape;521;p44', 5), 150: ('Google Shape;522;p44', 5), 151: ('Google Shape;523;p44', 5), 152: ('Google Shape;524;p44', 5), 153: ('Google Shape;530;p44', 5), 154: ('Google Shape;531;p44', 5), 155: ('Google Shape;532;p44', 5), 156: ('Google Shape;533;p44', 5), 157: ('Google Shape;534;p44', 5), 159: ('Google Shape;536;p44', 5), 160: ('Google Shape;537;p44', 5), 161: ('Google Shape;538;p44', 5), 162: ('Google Shape;539;p44', 5), 163: ('Google Shape;540;p44', 5), 165: ('Google Shape;542;p44', 5), 166: ('Google Shape;543;p44', 5), 167: ('Google Shape;544;p44', 5), 168: ('Google Shape;545;p44', 5), 169: ('Google Shape;546;p44', 5), 170: ('Google Shape;547;p44', 5), 171: ('Google Shape;548;p44', 5), 172: ('Google Shape;549;p44', 5), 185: ('Google Shape;581;p45', 6), 186: ('Google Shape;582;p45', 6), 187: ('Google Shape;584;p45', 6), 189: ('Google Shape;587;p45', 6), 191: ('Google Shape;590;p45', 6), 193: ('Google Shape;593;p45', 6), 195: ('Google Shape;596;p45', 6), 198: ('Google Shape;600;p45', 6), 199: ('Google Shape;602;p45', 6), 200: ('Google Shape;603;p45', 6), 201: ('Google Shape;605;p45', 6), 202: ('Google Shape;606;p45', 6), 203: ('Google Shape;608;p45', 6), 204: ('Google Shape;609;p45', 6), 205: ('Google Shape;611;p45', 6), 206: ('Google Shape;612;p45', 6), 207: ('Google Shape;614;p45', 6), 208: ('Google Shape;615;p45', 6), 209: ('Google Shape;617;p45', 6), 210: ('Google Shape;618;p45', 6), 211: ('Google Shape;620;p45', 6), 212: ('Google Shape;621;p45', 6), 213: ('Google Shape;623;p45', 6), 218: ('Google Shape;633;p46', 7), 222: ('Google Shape;637;p46', 7), 224: ('Google Shape;639;p46', 7), 226: ('Google Shape;641;p46', 7), 227: ('Google Shape;642;p46', 7), 231: ('Google Shape;646;p46', 7), 233: ('Google Shape;648;p46', 7), 235: ('Google Shape;650;p46', 7), 236: ('Google Shape;651;p46', 7), 242: ('Google Shape;657;p46', 7), 244: ('Google Shape;659;p46', 7), 247: ('Google Shape;667;p47', 8), 250: ('Google Shape;670;p47', 8), 251: ('Google Shape;671;p47', 8), 253: ('Google Shape;673;p47', 8), 255: ('Google Shape;675;p47', 8), 256: ('Google Shape;676;p47', 8), 260: ('Google Shape;680;p47', 8), 262: ('Google Shape;682;p47', 8), 264: ('Google Shape;684;p47', 8), 265: ('Google Shape;685;p47', 8), 269: ('Google Shape;689;p47', 8), 271: ('Google Shape;691;p47', 8), 273: ('Google Shape;693;p47', 8), 276: ('Google Shape;701;p48', 9), 280: ('Google Shape;705;p48', 9), 282: ('Google Shape;707;p48', 9), 284: ('Google Shape;709;p48', 9), 285: ('Google Shape;710;p48', 9), 289: ('Google Shape;714;p48', 9), 291: ('Google Shape;716;p48', 9), 293: ('Google Shape;718;p48', 9), 294: ('Google Shape;719;p48', 9), 298: ('Google Shape;723;p48', 9), 300: ('Google Shape;725;p48', 9), 302: ('Google Shape;727;p48', 9)}

  editPPTX(prs, ref, items)
  buf = BytesIO()
  prs.save(buf)
  return