import json
from app.allprompts import allprompts
from app.initialprompts import initial_prompts


def attach_initial_prompts(data):
  content = data["content"]
  # print(content)
  return initial_prompts(content)

def attach_second_prompt(data, cat):
  return allprompts(data, 1, cat)

def attach_third_prompt(data, cat):
  return allprompts(data, 2, cat)

def attach_final_prompt(data, cat):
  return allprompts(data, 3, cat)


#{ 0: education_communication_prompts, 1: clinical_practice_prompts, 2: competitive_intelligence_prompts}
def attach_education_prompts(content, prompt_number, records):
  if prompt_number == 1:
    response = attach_second_prompt(content, 0)
  elif prompt_number == 2:
    response = attach_third_prompt(content, 0)
  elif prompt_number == 3:
    response = attach_final_prompt(content, 0)
  rec = json.dumps(records, indent=2)
  return response+rec