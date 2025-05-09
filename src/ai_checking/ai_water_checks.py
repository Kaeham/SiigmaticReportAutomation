from src.ai_checking.client import *
from src.ai_checking.prompts import WATER_PROMPT

def water_penetration_checks(water_data):
    val, comment = water_data[:2]
    prompt = f"{WATER_PROMPT} Final Pressure: {val}. Comments: {comment}"
    return ask_gpt(prompt)
