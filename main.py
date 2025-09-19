from fastapi import FastAPI, Request, Response
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from app.prompting import attach_education_prompts, attach_initial_prompts, attach_clinical_prompts, attach_competitive_prompts
from app.pptxgenerator import pptx_maker
from app.demosite import data_preprocess, second_process
from app.data_analytics.pptx_generation import full_replacement
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