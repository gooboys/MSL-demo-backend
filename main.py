from fastapi import FastAPI, Request, Response, HTTPException, Header, BackgroundTasks
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from app.prompting import attach_education_prompts, attach_initial_prompts, attach_clinical_prompts, attach_competitive_prompts
from app.pptxgenerator import pptx_maker
from app.demosite import data_preprocess, second_process
from app.data_analytics.pptx_generation import full_replacement
from app.pptxdata import true_replacement
from typing import List
from pydantic import BaseModel
from typing import Optional, Dict, Any
import httpx, uuid, time

import io

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allows all origins
    allow_credentials=True,
    allow_methods=["*"],  # Allows all methods
    allow_headers=["*"],  # Allows all headers
)

"""STUFF FOR SINGLE USE TEXT EXTRACTION !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"""
# In-memory store for demo; replace with Postgres table
JOBS: Dict[str, Any] = {}
TTL_SECONDS = 600  # 10 minutes

def _now() -> int:
  return int(time.time())

def _sweep_expired() -> None:
  now = _now()
  for jid, rec in list(JOBS.items()):
    if rec.get("expires_at", 0) <= now:
      del JOBS[jid]

def create_job(id, initial: Optional[Dict[str, Any]] = None) -> str:
  jid = id
  now = _now()
  JOBS[jid] = {
    "status": "queued",
    "created_at": now,
    "updated_at": now,
    "expires_at": now + TTL_SECONDS,
    "result": None,
    "error": None
  }
  return jid

# In-place for webhook calls
webhook = "https://yichao.app.n8n.cloud/webhook-test/b4fcda5e-d82e-4b6b-b3c5-b721375d794a"

@app.post("/single-slide-pptx")
async def start_single_slide(request: Request):
  _sweep_expired()
  data = await request.json()
  job_id = data["id"]
  content = data["content"]
  create_job(job_id)
  with httpx.Client(timeout=30) as client:
    resp = client.post(webhook, json=content, headers=None)
    resp.raise_for_status()
    if resp.headers.get("content-type", "").startswith("application/json"):
      return resp.json()
    return {"status": "accepted", "raw": resp.text}

"""End STUFF FOR SINGLE USE TEXT EXTRACTION !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"""

@app.get("/")
async def root():
    return {"status": "Chart API is alive"}

@app.get("/MSL-preprocessing", response_model=List[str])
async def process_data(request: Request):
  data = await request.json()
  dat = attach_initial_prompts(data)
  # buf = query2
  # return StreamingResponse(buf, media_type="image/png")

  if dat is None: return JSONResponse(status_code=500,
                                          content={"error": "Prompting process failed"})

  return JSONResponse(content=dat)

@app.post("/MSL-prompting")
async def process_data(request: Request):
  data = await request.json()
  content = data["content"]
  run = data["counter"]
  records = data["records"]
  # print(data)
  dat = attach_education_prompts(content, run, records)
  print(dat)
  # buf = query2
  # return StreamingResponse(buf, media_type="image/png")

  if dat is None: return JSONResponse(status_code=500,
                                          content={"error": "Prompting process failed"})

  return {"data":dat,"run":run}

@app.post("/MSL-prompting-clin")
async def process_data(request: Request):
  data = await request.json()
  content = data["content"]
  run = data["counter"]
  records = data["records"]
  # print(data)
  dat = attach_clinical_prompts(content, run, records)
  print(dat)
  # buf = query2
  # return StreamingResponse(buf, media_type="image/png")

  if dat is None: return JSONResponse(status_code=500,
                                          content={"error": "Prompting process failed"})

  return {"data":dat,"run":run}

@app.post("/MSL-prompting-comp")
async def process_data(request: Request):
  data = await request.json()
  content = data["content"]
  run = data["counter"]
  records = data["records"]
  # print(data)
  dat = attach_competitive_prompts(content, run, records)
  print(dat)
  # buf = query2
  # return StreamingResponse(buf, media_type="image/png")

  if dat is None: return JSONResponse(status_code=500,
                                          content={"error": "Prompting process failed"})

  return {"data":dat,"run":run}


@app.post("/PPTX-generation")
async def pptx_generation(request: Request):
  data = await request.json()
  rec = data["content"]
  # print("this is rec:\n", rec)
  pptx = pptx_maker(rec)
  
  if pptx is None: return JSONResponse(status_code=500, content={"error":"Failed to generate pptx"})

@app.get("/presentation")
async def send_pptx(request: Request):
  data = await request.json()
  statdata = data_preprocess(data)
  stat = second_process(statdata)
  patient = data["patient_management"]
  education = data["education"]
  competitive = data["competitive"]
  # print("patient:\n")
  # print(patient)
  # print("\n\n")
  # print("education:\n")
  # print(education)
  # print("\n\n")
  # print("competitive:\n")
  # print(competitive)
  presentation = full_replacement(stat,patient,education,competitive)
  if presentation is None: return JSONResponse(status_code=500, content={"error":"Failed to generate pptx"})
  return Response(
    content=presentation,
    media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    headers={"Content-Disposition": "attachment; filename=out.pptx"}
  )

@app.get("/pdf")
async def pdf_generator(request: Request):
  content = await request.json()
  data = content["data"]
  print(data)
  
  return

# Path for actual pptx generation
@app.get("/real-pptx")
async def real_pptx(request: Request):
  data = await request.json()
  statdata = data_preprocess(data)
  stat = second_process(statdata)
  patient = data["patient_management"]
  education = data["education"]
  competitive = data["competitive"]
  print("single")
  single = data["single"]
  print(single)
  
  pptx_bytes = true_replacement(stat, patient, education, competitive, single)
  if not pptx_bytes:
    raise HTTPException(status_code=500, detail="Failed to generate pptx")

  return Response(
    content=pptx_bytes,
    media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    headers={"Content-Disposition": 'attachment; filename="out.pptx"'},
  )


# Path for single use case pptx processing and storing
@app.post("/single-slide-ppt")
async def one_slide_generation(request: Request):
  content = await request.json()
  job_id = content["id"]
  data = content["content"]
  print("id: ", job_id)
  print("Data: ", data)

  return

# Path for single use case pptx fetching