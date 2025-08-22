import json

def initial_prompts(data):
  # 0 is education, 1 is clinical, 2 is competitive intelligence
  prompts = {
    0: """Analyze the insight to identify specific knowledge and communication gaps:

**Focus Areas:**
- **Clinical Knowledge:** What evidence, mechanisms, or protocols are misunderstood or missing?
- **Practical Application:** What barriers exist in translating knowledge to practice?
- **Communication:** Which channels, timing, or trusted sources aren't being leveraged?
- **Context:** What system, competitive, or patient population factors are overlooked?

**For Each Gap:**
- Provide supporting quotes from the insight
- Distinguish between explicit (stated) and implicit (inferred) gaps
- Assess impact on clinical decision-making
- Consider how stakeholder tier/therapeutic area amplifies the gap

**Key Questions:** What assumptions reveal knowledge deficits? What language suggests incomplete understanding? How does timing affect gap relevance?

Format your response as:
[PROMPT 1]
[Copy this entire prompt here]

[PROMPT 1 OUTPUT]
[Your analysis here]

The data is below:
""",
  1: """Analyze the insight to identify clinical knowledge or practice gaps revealed by the healthcare professional:

**Clinical Gap Categories:**
- **Clinical Knowledge:** What uncertainties exist in dosing, mechanisms, contraindications, or clinical evidence?
- **Patient Selection:** What challenges exist in identifying appropriate candidates or risk stratification?
- **Practice Implementation:** What barriers exist in medication administration, monitoring, or workflow integration?
- **Safety Management:** What gaps exist in side-effect recognition, management, or risk mitigation?

**For Each Gap:**
- Provide supporting quotes from the insight
- Distinguish between knowledge deficits vs. practical implementation barriers
- Assess impact on patient care quality and clinical outcomes
- Consider how stakeholder tier/therapeutic area context influences the gap severity

**Evidence Analysis:**
- What clinical uncertainties are explicitly stated vs. implied?
- What practice variations or inconsistencies are suggested?
- How does the timing/context of this insight affect clinical relevance?

Format your response as:
[PROMPT 1]
[Copy this entire prompt here]

[PROMPT 1 OUTPUT]
[Your analysis here]

The data is below:
""",
  2: """Analyze the insight to identify knowledge gaps about [Product]'s positioning versus competitors:

**Competitive Gap Categories:**
- **Efficacy Comparisons:** What misunderstandings exist about relative clinical outcomes, response rates, or durability?
- **Safety Profiles:** What gaps exist in understanding comparative risk-benefit profiles or tolerability?
- **Value Positioning:** What misconceptions about cost-effectiveness, health economics, or overall value exist?
- **Market Perception:** What inaccurate beliefs about [Product]'s place in therapy or competitive advantages persist?

**Competitor Analysis Focus:**
- Which specific competitors are mentioned or implied in comparisons?
- What competitive data or claims are being referenced (accurately or inaccurately)?
- What competitive strengths are being overvalued or [Product] strengths undervalued?
- How do stakeholder perceptions align with or diverge from clinical evidence?

**For Each Gap:**
- Provide supporting quotes from the insight
- Identify whether gaps favor competitors or create neutral confusion
- Assess impact on [Product]'s competitive standing and market access
- Consider how stakeholder tier/therapeutic area affects competitive sensitivity

**Evidence Analysis:**
- What competitive assumptions are explicitly stated vs. implied?
- What language suggests outdated or incomplete competitive intelligence?
- How does timing affect competitive relevance (new data, approvals, market changes)?

Format your response as:
[PROMPT 1]
[Copy this entire prompt here]

[PROMPT 1 OUTPUT]
[Your analysis here]

The data is below:
"""
  }
  remove = [
            "KOL Full Name", 
            "Therapeutic Area", 
            "Product Discussed", 
            "MSL / Submitter Name", 
            "Company Sponsored Research Details",
            "US: Unsolicited Request for Information"
            ]
  ed = []
  clin = []
  comp = []
  for i in data:
    for key in remove:
      i.pop(key, None)
    print(i)
    if i["Insight Category"] == "Educational and Communication":
      ed.append(i)
    elif i["Insight Category"] == "Clinical Practice":
      clin.append(i)
    elif i["Insight Category"] == "Competitive Intelligence":
      comp.append(i)
  eptext = json.dumps(ed, indent=2)
  kymtext = json.dumps(clin, indent=2)
  rittext = json.dumps(comp, indent=2)
  education_prompt = prompts[0]+eptext
  clinical_prompt = prompts[1]+kymtext
  comp_prompt = prompts[2]+rittext
  return [education_prompt,clinical_prompt,comp_prompt]
