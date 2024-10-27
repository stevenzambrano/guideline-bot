import os
from pinecone import Pinecone
from dotenv import load_dotenv

#make sure you load environmental variables
load_dotenv()

pc = Pinecone(api_key= os.getenv("PINECONE_API_KEY"))
index = pc.Index("guideline-bot")

def rule_ids(vectorized_dict):
    # Retrieve rule IDs
    rules_dict = {}

    for index4, values4 in vectorized_dict.items():
        temp4 = {}
        for index5, values5 in values4.items():
            query_response = index.query(vector=values5, top_k=1, include_values=False, include_metadata=False)
            temp4[index5] = query_response['matches'][0]['id']
        rules_dict[index4] = temp4
        
    return rules_dict