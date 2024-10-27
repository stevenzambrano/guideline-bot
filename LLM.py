# Defining the model
import os 
from openai import OpenAI
from dotenv import load_dotenv

#make sure you load environmental variables
load_dotenv()

openai_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key = openai_key)
print(client)

def classify_sentence(sentence):
    system_prompt = """You are a financial expert in investment guidelines. You will classify each sentence in triple backticks given to you as either 'investment restriction' or 'general text'. I have provided
    some examples of sentences with there classification. These have been generalised where the X% or £X can represent any numerical value."""

    response = client.chat.completions.create(
        model = "gpt-3.5-turbo",
        messages=[{"role":"system", "content": system_prompt},
                  {"role": "user", "content": "sentence: ```The cumulative PV01 of the benchmark at each tenor should not deviate by more than X% than the total benchmark ``` " },
                  {"role": "assistant", "content": "investment restriction"},
                  {"role": "user", "content": "sentence: ```The duration of liabilities should match within X% of the benchmark liabilities duration ``` " },
                  {"role": "assistant", "content": "investment restriction"},
                  {"role": "user", "content": "sentence: ```The Bond PV01 of the portfolio shall not deviate by more than X% of the benchmark ``` " },
                  {"role": "assistant", "content": "investment restriction"},
                  {"role": "user", "content": "sentence: ```This document outlines the investment policy for the fiscal year. ``` " },
                  {"role": "assistant", "content": "general text"},
                  {"role": "user", "content": "sentence: ```Asset allocation should not deviate by more than X% in any asset class ``` " },
                  {"role": "assistant", "content": "investment restriction"},
                  {"role": "user", "content": "sentence: ```The Swap PV01 of the portfolio shall not deviate by more than X% of the benchmark``` " },
                  {"role": "assistant", "content": "investment restriction"},
                  {"role": "user", "content": "sentence: ```Cash holdings should not exceed X% of the total portfolio ``` " },
                  {"role": "assistant", "content": "investment restriction"},
                  {"role": "user", "content": "sentence: ```The portfolio manager is responsible for executing trades based on the client's needs. ``` " },
                  {"role": "assistant", "content": "general text"},
                  {"role": "user", "content": "sentence: ```Permitted assets: UK Gilts, Repo, Reverse repo ``` " },
                  {"role": "assistant", "content": "investment restriction"},
                  {"role": "user", "content": "sentence: ```The PV01 of the portfolio shall not deviate by more than X% of the benchmark ``` " },
                  {"role": "assistant", "content": "investment restriction"},
                  {"role": "user", "content": "sentence: ```Portfolio duration should be within X% of the benchmark duration ``` " },
                  {"role": "assistant", "content": "investment restriction"},
                  {"role": "user", "content": "sentence: ```Derivatives exposure should not exceed X% of the portfolio value ``` " },
                  {"role": "assistant", "content": "investment restriction"},
                  {"role": "user", "content": f"sentence: ```{sentence}```"}],
                  temperature=0,
                  max_tokens=10
    )
    return response.choices[0].message.content.strip().lower()


#call open ai to embedd your data
def get_embedding(text_to_embed):
    response = client.embeddings.create(
        model = "text-embedding-3-small",
        input = [text_to_embed]
    )
    embedding = response.data[0].embedding
    return embedding