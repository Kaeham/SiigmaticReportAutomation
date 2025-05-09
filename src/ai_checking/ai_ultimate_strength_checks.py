from src.ai_checking.client import *
from src.ai_checking.prompts import ULTIMATE_PROMPT

def ultimate_strength_checks(ultimate_data):
    flattened = " ".join(item if isinstance(item, str) else " ".join(item) for item in ultimate_data)
    prompt = ULTIMATE_PROMPT + flattened
    return ask_gpt(prompt)
