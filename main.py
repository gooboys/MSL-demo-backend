from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse, JSONResponse
from app.prompting import attach_prompts, attach_initial_prompts

import io

app = FastAPI()

@app.get("/")
async def root():
    return {"status": "Chart API is alive"}

@app.post("/MSL-preprocessing")
async def process_data(request: Request):
  data = await request.json()
  string = attach_initial_prompts(data)
  # buf = query2
  # return StreamingResponse(buf, media_type="image/png")

  if string is None: return JSONResponse(status_code=500,
                                          content={"error": "Prompting process failed"})

  return JSONResponse(content=string)