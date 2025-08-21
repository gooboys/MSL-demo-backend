def attach_prompts(data):
  remove = ["KOL Full Name", 
            "Therapeutic Area", 
            "Product Discussed", 
            "MSL / Submitter Name", 
            "Company Sponsored Research Details",
            "Insight Category",
            "US: Unsolicited Request for Information"
            ]
  for i in remove:del data[i]
  
  return data