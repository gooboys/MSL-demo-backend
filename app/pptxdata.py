from pptx import Presentation
from pptx.util import Pt, Inches, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from io import BytesIO
import os

# EMU conversions
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
    raise ValueError(f"Shape id={getattr(shp, 'shape_id', '?')} has no text frame (type={shp.shape_type.name})")

  tf.clear()  # clears paragraphs; leaves one empty paragraph
  p = tf.paragraphs[0]
  run = p.add_run()
  run.text = text

  # apply formatting
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

def editPPTX(pres, ref, items):
  slidetoid = {}
  for id in items.keys():
    curslide = items[id][1]
    slidetoid[curslide] = slidetoid.get(curslide, []).append(id)
  for si, slide in enumerate(pres.slides):
    for element in slidetoid[si]:
      cDict = ref[element]
      _overwrite_shape_text(
        element,
        text = cDict.get("text", ""),
        font_name = cDict.get("font","Calibri"),
        font_size = cDict.get("font_size",12),
        font_color = cDict.get("font_color",(0, 0, 0)),
        bold = cDict.get("bold",None),
        italic = cDict.get("italic",None)
      )
  return

def true_replacement(stats, patient, education, competitive):
  base_dir = os.path.dirname(os.path.abspath(__file__))
  template_path = os.path.join(base_dir, "Acquis Template.pptx")
  print('base_dir: '+ base_dir)
  print('template_path: '+template_path)
  prs = Presentation(template_path)

  ref = {
    146: {
      "text": "",
      "font": "Calibri",
      "font_size": 16,
      "font_color": (0,0,0)
    }
  }
  items = {146: ('Google Shape;518;p44', 5), 148: ('Google Shape;520;p44', 5), 149: ('Google Shape;521;p44', 5), 150: ('Google Shape;522;p44', 5), 151: ('Google Shape;523;p44', 5), 152: ('Google Shape;524;p44', 5), 153: ('Google Shape;530;p44', 5), 154: ('Google Shape;531;p44', 5), 155: ('Google Shape;532;p44', 5), 156: ('Google Shape;533;p44', 5), 157: ('Google Shape;534;p44', 5), 159: ('Google Shape;536;p44', 5), 160: ('Google Shape;537;p44', 5), 161: ('Google Shape;538;p44', 5), 162: ('Google Shape;539;p44', 5), 163: ('Google Shape;540;p44', 5), 165: ('Google Shape;542;p44', 5), 166: ('Google Shape;543;p44', 5), 167: ('Google Shape;544;p44', 5), 168: ('Google Shape;545;p44', 5), 169: ('Google Shape;546;p44', 5), 170: ('Google Shape;547;p44', 5), 171: ('Google Shape;548;p44', 5), 172: ('Google Shape;549;p44', 5), 185: ('Google Shape;581;p45', 6), 186: ('Google Shape;582;p45', 6), 187: ('Google Shape;584;p45', 6), 189: ('Google Shape;587;p45', 6), 191: ('Google Shape;590;p45', 6), 193: ('Google Shape;593;p45', 6), 195: ('Google Shape;596;p45', 6), 198: ('Google Shape;600;p45', 6), 199: ('Google Shape;602;p45', 6), 200: ('Google Shape;603;p45', 6), 201: ('Google Shape;605;p45', 6), 202: ('Google Shape;606;p45', 6), 203: ('Google Shape;608;p45', 6), 204: ('Google Shape;609;p45', 6), 205: ('Google Shape;611;p45', 6), 206: ('Google Shape;612;p45', 6), 207: ('Google Shape;614;p45', 6), 208: ('Google Shape;615;p45', 6), 209: ('Google Shape;617;p45', 6), 210: ('Google Shape;618;p45', 6), 211: ('Google Shape;620;p45', 6), 212: ('Google Shape;621;p45', 6), 213: ('Google Shape;623;p45', 6), 218: ('Google Shape;633;p46', 7), 222: ('Google Shape;637;p46', 7), 224: ('Google Shape;639;p46', 7), 226: ('Google Shape;641;p46', 7), 227: ('Google Shape;642;p46', 7), 231: ('Google Shape;646;p46', 7), 233: ('Google Shape;648;p46', 7), 235: ('Google Shape;650;p46', 7), 236: ('Google Shape;651;p46', 7), 242: ('Google Shape;657;p46', 7), 244: ('Google Shape;659;p46', 7), 247: ('Google Shape;667;p47', 8), 250: ('Google Shape;670;p47', 8), 251: ('Google Shape;671;p47', 8), 253: ('Google Shape;673;p47', 8), 255: ('Google Shape;675;p47', 8), 256: ('Google Shape;676;p47', 8), 260: ('Google Shape;680;p47', 8), 262: ('Google Shape;682;p47', 8), 264: ('Google Shape;684;p47', 8), 265: ('Google Shape;685;p47', 8), 269: ('Google Shape;689;p47', 8), 271: ('Google Shape;691;p47', 8), 273: ('Google Shape;693;p47', 8), 276: ('Google Shape;701;p48', 9), 280: ('Google Shape;705;p48', 9), 282: ('Google Shape;707;p48', 9), 284: ('Google Shape;709;p48', 9), 285: ('Google Shape;710;p48', 9), 289: ('Google Shape;714;p48', 9), 291: ('Google Shape;716;p48', 9), 293: ('Google Shape;718;p48', 9), 294: ('Google Shape;719;p48', 9), 298: ('Google Shape;723;p48', 9), 300: ('Google Shape;725;p48', 9), 302: ('Google Shape;727;p48', 9)}

  editPPTX(prs, ref, items)
  return