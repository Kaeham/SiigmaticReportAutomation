from src.ai_checking.client import *
from src.ai_checking.prompts import FORCE_PROMPT

def operating_force_checks(of_data):
    open_force, close_force = of_data
    results = []
    for init, maint in [open_force, close_force]:
        prompt = f"{FORCE_PROMPT}\n{init}, {maint}"
        results.append(ask_gpt(prompt))
    return results
