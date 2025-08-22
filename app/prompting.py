import json
from app.allprompts import allprompts
from app.initialprompts import initial_prompts


def attach_initial_prompts(data):
  content = data["content"]
  # print(content)
  return initial_prompts(content)

def attach_second_prompt(data, prev_prompt):
  return allprompts(data, 1, prev_prompt)

def attach_third_prompt(data, prev_prompt):
  return allprompts(data, 2, prev_prompt)

def attach_final_prompt(data, prev_prompt):
  return allprompts(data, 3, prev_prompt)

def attach_prompts(data):
  content = data["content"]
  prompt_number = 0
  if "PROMPT 3" in content:
    prompt_number = 3
  if "PROMPT 2" in content:
    prompt_number = 2
  if "PROMPT 1" in content:
    prompt_number = 1
  if prompt_number == 1:
    response = attach_second_prompt(content)
  elif prompt_number == 2:
    response = attach_third_prompt(content)
  elif prompt_number == 3:
    response = attach_final_prompt(content)
  return