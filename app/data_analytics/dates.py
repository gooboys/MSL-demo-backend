from typing import List, Dict
from datetime import datetime

def get_date_range(rows: List[Dict]) -> str:
  dates = []
  for r in rows:
    d = r.get("Report Date")
    if d:
      try:
        # parse format m/d/YYYY
        dt = datetime.strptime(d.strip(), "%m/%d/%Y")
        dates.append(dt)
      except ValueError:
        continue  # skip bad formats
  
  if not dates:
    return "No valid dates"
  
  earliest = min(dates)
  latest = max(dates)
  
  # Format as "Month YYYY"
  start_str = earliest.strftime("%B %Y")
  end_str = latest.strftime("%B %Y")
  
  if start_str == end_str:
    return start_str
  return f"{start_str} - {end_str}"