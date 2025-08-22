def allprompts(data, id, cat):
  education_communication_prompts = {
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
[Your analysis here]""",
    1: """Based on the gaps identified in the previous step, analyze how they drive stakeholder behaviors and attitudes:

**Behavioral Impact Areas:**
- **Decision-Making:** How do gaps affect treatment choices and clinical algorithms?
- **Risk Perception:** Are they over/under-estimating risks due to incomplete information?
- **Information-Seeking:** What sources are they turning to or avoiding?
- **Communication:** How do gaps affect patient counseling and peer discussions?

**Psychological Factors:**
- What cognitive biases (confirmation, anchoring) are reinforced by gaps?
- What emotional responses (anxiety, overconfidence) stem from uncertainty?
- How do gaps threaten or support their professional identity?
- How does their stakeholder tier influence behavioral intensity?

**Evidence Analysis:**
For each gap, identify:
- Observable behaviors mentioned in the insight
- Predictable behaviors based on the gap
- Whether responses are strong patterns or situational

Format your response as:
[PREVIOUS CONTENT]
[Include all previous prompts and outputs]

[PROMPT 2]
[Copy this entire prompt here]

[PROMPT 2 OUTPUT]
[Your analysis here]""",
    2:"""Based on the gaps and behaviors identified in previous steps, identify the underlying unmet needs driving stakeholder gaps and behaviors:

**Need Categories:**
- **Functional:** What clinical capabilities or decision-support tools are missing?
- **Informational:** What evidence depth, comparative data, or contextual knowledge is lacking?
- **Emotional:** What confidence, reassurance, or anxiety reduction is needed?
- **Social:** What peer validation, expert consultation, or institutional support is required?

**Need Prioritization:**
For each need, assess:
- **Urgency:** How quickly must this be addressed?
- **Impact:** How much would meeting this need change behavior?
- **Feasibility:** How realistic is it for Medical Affairs to address?

**Root Cause Analysis:**
- Why have current educational efforts failed to meet these needs?
- What barriers prevent need fulfillment?
- What incorrect assumptions about stakeholder needs exist?

**Journey Stage Consideration:**
What do they need at each stage: Awareness → Consideration → Trial → Adoption → Advocacy?

Format your response as:
[PREVIOUS CONTENT]
[Include all previous prompts and outputs]

[PROMPT 3]
[Copy this entire prompt here]

[PROMPT 3 OUTPUT]
[Your analysis here]""",
    3:"""Based on the comprehensive analysis from previous steps, develop specific Medical Affairs actions to close gaps and meet identified needs:

**Action Framework:**
- **Immediate (0-30 days):** Quick wins using existing resources
- **Short-term (1-3 months):** New content, tools, or targeted outreach
- **Long-term (3-12 months):** Strategic changes to approach or capabilities

**For Each Action, Specify:**
- **What:** Precise description of the action
- **Why:** Which gaps/needs it addresses and expected behavioral impact
- **How:** Implementation approach and required resources
- **When:** Timeline and key milestones
- **Success:** How effectiveness will be measured

**Design Considerations:**
- How will actions account for identified cognitive biases and emotional needs?
- What combination of channels will maximize stakeholder engagement?
- How will actions integrate with existing workflows and decision processes?
- What behavioral change model underlies the approach?

**Implementation Requirements:**
- Personnel and budget needs
- Cross-functional collaboration required
- Risk mitigation strategies
- Scalability for similar stakeholder challenges

Format your response as:
[PREVIOUS CONTENT]
[Include all previous prompts and outputs]

[PROMPT 4]
[Copy this entire prompt here]

[PROMPT 4 OUTPUT]
[Your analysis here]"""
  }

  clinical_practice_prompts = {
    0: """Analyze the insight to identify clinical knowledge or practice gaps revealed by the healthcare professional:

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
[Your analysis here]""",
    1: """Based on the clinical gaps identified, analyze how they influence clinical behavior and decision-making:

**Clinical Decision Impact:**
- **Treatment Initiation:** How do gaps affect willingness to prescribe or recommend?
- **Dosing Decisions:** What impact on dose selection, titration, or duration choices?
- **Patient Management:** How do gaps influence monitoring frequency, follow-up protocols?
- **Safety Responses:** What behaviors emerge around adverse event management?

**Practice Pattern Analysis:**
- **Delays:** What hesitations or postponements in clinical action are evident?
- **Avoidance:** Are there signs of preferring alternatives due to uncertainty?
- **Over/Under-utilization:** How do gaps lead to inappropriate usage patterns?
- **Referral Patterns:** What impact on specialist consultation or care coordination?

**Clinical Context Factors:**
- How does practice setting (academic vs. community) modify behavioral responses?
- What institutional or formulary constraints amplify gap-driven behaviors?
- How do patient characteristics influence gap-related decision patterns?

Format your response as:
[PREVIOUS CONTENT]
[Include all previous prompts and outputs]

[PROMPT 2]
[Copy this entire prompt here]

[PROMPT 2 OUTPUT]
[Your analysis here]""",
    2: """Based on identified gaps and behaviors, infer the underlying unmet needs of clinicians and institutions:

**Clinician Need Categories:**
- **Knowledge:** What clearer evidence, guidelines, or educational content is needed?
- **Skills:** What training, competency development, or hands-on support is required?
- **Tools:** What decision aids, protocols, or assessment instruments are missing?
- **Support:** What peer consultation, expert guidance, or institutional backing is needed?

**Institutional Need Categories:**
- **Systems:** What workflow optimization, EHR integration, or process improvements are needed?
- **Resources:** What staffing, equipment, or logistical support gaps exist?
- **Policies:** What formulary decisions, pathway development, or guideline updates are required?

**Need Prioritization:**
For each need, assess:
- **Clinical Impact:** How significantly would addressing this improve patient outcomes?
- **Feasibility:** How realistic is it for Medical Affairs to influence this need?
- **Urgency:** How quickly must this be addressed to prevent suboptimal care?
- **Scope:** How many clinicians/institutions share this need?

**Root Cause Analysis:**
- Why haven't existing clinical resources addressed these needs?
- What systemic barriers prevent optimal clinical practice?
- What assumptions about clinician capabilities have proven incorrect?

Format your response as:
[PREVIOUS CONTENT]
[Include all previous prompts and outputs]

[PROMPT 3]
[Copy this entire prompt here]

[PROMPT 3 OUTPUT]
[Your analysis here]""",
    3: """Based on the comprehensive clinical analysis, develop specific Medical Affairs actions to address practice challenges and support stakeholders:

**Clinical Action Framework:**
- **Immediate (0-30 days):** Rapid clinical support using existing resources
- **Short-term (1-3 months):** New clinical tools, training, or targeted education
- **Long-term (3-12 months):** Strategic clinical program development or guideline influence

**For Each Clinical Action, Specify:**
- **What:** Precise clinical intervention or support mechanism
- **Why:** Which clinical gaps/needs it addresses and expected practice improvement
- **How:** Clinical implementation approach and required expertise/resources
- **When:** Clinical timeline and key practice milestones
- **Success:** How clinical effectiveness and practice change will be measured

**Clinical Design Considerations:**
- How will actions integrate with existing clinical workflows and decision points?
- What clinical evidence or validation will support credibility and adoption?
- How will actions account for different practice settings and patient populations?
- What clinical champions or KOL engagement strategy will drive uptake?

**Implementation Requirements:**
- Clinical expertise and MSL involvement needed
- Educational content development and delivery mechanisms
- Cross-functional collaboration (Clinical, Regulatory, Market Access)
- Measurement of clinical outcomes and practice behavior change

Format your response as:
[PREVIOUS CONTENT]
[Include all previous prompts and outputs]

[PROMPT 4]
[Copy this entire prompt here]

[PROMPT 4 OUTPUT]
[Your analysis here]"""
  }

  competitive_intelligence_prompts = {
    0: """Analyze the insight to identify knowledge gaps about [Product]'s positioning versus competitors:

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
[Your analysis here]""",

    1: """Based on competitive gaps identified, analyze how perceptions and comparisons drive stakeholder behavior:

**Decision-Making Impact:**
- **Prescribing Patterns:** How do competitive perceptions influence treatment selection and sequencing?
- **Patient Counseling:** What competitive messaging are clinicians sharing with patients?
- **Referral Decisions:** How do competitive views affect specialist consultation or care pathways?
- **Formulary Influence:** What impact on payer coverage decisions and access restrictions?

**Competitive Behavior Analysis:**
- **Default Preferences:** Are competitors becoming the "go-to" choice due to perception gaps?
- **Risk Aversion:** Is competitive uncertainty driving conservative prescribing toward familiar options?
- **Value Questioning:** Are stakeholders questioning [Product]'s cost-benefit compared to alternatives?
- **Advocacy Patterns:** Are stakeholders actively promoting competitors over [Product]?

**Stakeholder-Specific Impacts:**
- **Clinicians:** How do competitive perceptions affect clinical confidence and recommendation strength?
- **Patients:** What competitive concerns influence patient acceptance and adherence?
- **Payers:** How do competitive assessments impact coverage and reimbursement decisions?
- **Institutions:** What competitive factors drive formulary and pathway decisions?

**Market Context Factors:**
- How does practice setting influence competitive decision-making?
- What external factors (guidelines, KOL opinions, peer influence) amplify competitive gaps?

Format your response as:
[PREVIOUS CONTENT]
[Include all previous prompts and outputs]

[PROMPT 2]
[Copy this entire prompt here]

[PROMPT 2 OUTPUT]
[Your analysis here]""",

    2: """Based on competitive perceptions and behaviors, infer the underlying needs of decision-makers that influence choice between [Product] and alternatives:

**Decision-Maker Need Categories:**

**Clinician Needs:**
- **Comparative Evidence:** What head-to-head data, real-world comparisons, or evidence synthesis is needed?
- **Decision Support:** What tools or frameworks would help navigate competitive choices?
- **Confidence Building:** What reassurance about [Product]'s competitive advantages is required?

**Patient Needs:**
- **Informed Choice:** What comparative information do patients need for shared decision-making?
- **Value Understanding:** What cost, convenience, or outcome comparisons matter most to patients?
- **Risk Clarity:** What safety or efficacy trade-offs need better explanation?

**Payer Needs:**
- **Economic Evidence:** What pharmacoeconomic data or budget impact analyses are missing?
- **Value Demonstration:** What real-world outcomes or cost-offset evidence is needed?
- **Risk Management:** What utilization management or safety monitoring tools are required?

**Institutional Needs:**
- **Workflow Integration:** How do competitive options differ in implementation complexity?
- **Resource Optimization:** What staffing, training, or system requirements favor certain choices?
- **Quality Metrics:** What outcome measures help differentiate competitive value?

**Need Prioritization:**
For each need, assess:
- **Competitive Impact:** How significantly would addressing this shift competitive preference?
- **Stakeholder Influence:** How much decision-making power does this stakeholder have?
- **Addressability:** How realistic is it for Medical Affairs to meet this need?
- **Timeline Sensitivity:** How quickly must this be addressed to prevent competitive loss?

Format your response as:
[PREVIOUS CONTENT]
[Include all previous prompts and outputs]

[PROMPT 3]
[Copy this entire prompt here]

[PROMPT 3 OUTPUT]
[Your analysis here]""",

    3: """Based on comprehensive competitive analysis, develop specific Medical Affairs actions to address needs and strengthen competitive positioning:

**Competitive Action Framework:**
- **Immediate (0-30 days):** Rapid competitive response using existing evidence and materials
- **Short-term (1-3 months):** Targeted competitive education and positioning initiatives
- **Long-term (3-12 months):** Strategic competitive intelligence and evidence generation

**For Each Competitive Action, Specify:**
- **What:** Precise competitive intervention or positioning strategy
- **Why:** Which competitive gaps/needs it addresses and expected market impact
- **How:** Implementation approach and required competitive intelligence/resources
- **When:** Timeline and key competitive milestones
- **Success:** How competitive effectiveness and market position improvement will be measured

**Competitive Design Considerations:**
- How will actions leverage [Product]'s true competitive advantages while addressing misconceptions?
- What evidence or data will most effectively counter competitive disadvantages?
- How will messaging be tailored to different stakeholder competitive concerns?
- What proactive vs. reactive competitive strategies are most appropriate?

**Stakeholder-Specific Approaches:**
- **Clinician Focus:** Comparative efficacy/safety education, decision algorithms, expert perspectives
- **Patient Focus:** Benefit-risk communication, value messaging, shared decision-making tools
- **Payer Focus:** Economic evidence, budget impact models, outcomes research
- **Institutional Focus:** Implementation support, workflow optimization, quality metrics

**Implementation Requirements:**
- Competitive intelligence gathering and analysis capabilities
- Cross-functional alignment (Medical, Commercial, Market Access) on competitive messaging
- KOL engagement strategy for competitive positioning
- Measurement of competitive market share, perception shifts, and stakeholder preference changes

Format your response as:
[PREVIOUS CONTENT]
[Include all previous prompts and outputs]

[PROMPT 4]
[Copy this entire prompt here]

[PROMPT 4 OUTPUT]
[Your analysis here]"""
  }
  prompt_dict = { 0: education_communication_prompts, 1: clinical_practice_prompts, 2: competitive_intelligence_prompts}
  return data + prompt_dict[cat][id]