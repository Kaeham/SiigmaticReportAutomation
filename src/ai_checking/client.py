from openai import OpenAI
import os
from dotenv import load_dotenv
load_dotenv()

key = os.environ.get("OPEN_AI_API_KEY")
client = OpenAI(api_key=key)

def ask_gpt(prompt):
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content