import json
from app.allprompts import allprompts
from app.initialprompts import initial_prompts


def attach_initial_prompts(data):
  content = data["content"]
  print(content)
  return initial_prompts(json.loads(content))

def attach_second_prompt(data, prev_prompt):
  return allprompts(data, 1, prev_prompt)

def attach_third_prompt(data, prev_prompt):
  return allprompts(data, 2, prev_prompt)

def attach_final_prompt(data, prev_prompt):
  return allprompts(data, 3, prev_prompt)

def attach_prompts(data, prompt_number, prev_prompt):
  content = data["content"]
  if prompt_number == 0:
    response = attach_initial_prompts(content)
  elif prompt_number == 1:
    response = attach_initial_prompts(content, prev_prompt)
  elif prompt_number == 2:
    response = attach_third_prompt(content, prev_prompt)
  elif prompt_number == 3:
    response = attach_final_prompt(content, prev_prompt)
  return