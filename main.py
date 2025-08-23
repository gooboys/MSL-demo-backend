from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse, JSONResponse
from app.prompting import attach_education_prompts, attach_initial_prompts
from typing import List

import io

app = FastAPI()

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
