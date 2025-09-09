import json
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
  content = data["content"]
  # print(content)

  rows = _extract_rows(content)
  print(f"rows type: {type(rows)}  len: {len(rows)}")
  print("sample row:", rows[0] if rows else None)

  _normalize_fields_inplace(rows)
  print("sample normalized row:", rows[0] if rows else None)

  return

def get_pdf_doc():
  return