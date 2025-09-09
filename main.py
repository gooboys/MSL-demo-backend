from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from app.prompting import attach_education_prompts, attach_initial_prompts, attach_clinical_prompts, attach_competitive_prompts
from app.pptxgenerator import pptx_maker
from app.demosite import get_powerpoint
from typing import List

import io

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allows all origins
    allow_credentials=True,
    allow_methods=["*"],  # Allows all methods
    allow_headers=["*"],  # Allows all headers
)

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
  presentation = get_powerpoint(data)
  return presentation

@app.get("/crm_refresh")
async def crm_refresh(request: Request):
  return

@app.get("/reasoning")
async def send_doc(request: Request):
  return

@app.get("checking")
async def check(request: Request):
  return