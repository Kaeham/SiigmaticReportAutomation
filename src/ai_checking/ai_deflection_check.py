from src.ai_checking.client import *
from src.ai_checking.prompts import DEFLECTION_PROMPT

def deflection_checks(deflection_data):
    results = []
    for name, pos, neg in deflection_data:
        prompt = f"{DEFLECTION_PROMPT}, {name}, {pos}, {neg}"
        results.append(ask_gpt(prompt))
    return results