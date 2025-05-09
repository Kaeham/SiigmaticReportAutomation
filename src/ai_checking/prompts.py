DEFLECTION_PROMPT =  "The following values are pressure value, positive and negative pressure differential. Are there any inconsistencies or causes for concern in the data? Start with True/False. Limit response to <50 words."


FORCE_PROMPT = "These values are average initiating and maintaining force. Are there inconsistencies? " + "Start with True/False. Limit to <25 words."

AIR_PROMPT = "Result for an air infiltration test at +-75Pa. State True/False for any concern. " + "Ignore 0 rating for negative pressure. Limit <20 words. Data: "

WATER_PROMPT = "I am providing the results of a water penetration test. " + "Numbers in the comments correspond to test pressure. " + "State True/False for concern. Limit to <50 words. "

ULTIMATE_PROMPT = "The following are positive and negative pressure readings for an ultimate strength test. " + "The Y/N values are observations. If any are Y, the specimen fails. " + "State True/False if there's concern. Limit <50 words. "
