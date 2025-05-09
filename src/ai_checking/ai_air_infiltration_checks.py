from src.ai_checking.client import *
from src.ai_checking.prompts import AIR_PROMPT

def air_infiltration_checks(ai_data):
    return ask_gpt(AIR_PROMPT + str(ai_data))
