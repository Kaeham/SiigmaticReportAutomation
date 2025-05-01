from openai import OpenAI


# check data

# deflection
# use the span and deflection results to check whether there is any inconsistency
def deflection_checks(deflection_data):
    """
    check deflection value with GPT
    :Inputs:
        - deflection_data: array contanining the deflection data
    :Outputs:
        - output: the output from GPT analysis
    """
    data, value = deflection_data
    outputs = []

    prompt = "The following values are pressure value, positive and negative pressure pressure differential, are there any inconsistencies or causes for concern in the data. " 
    prompt += "if there are, state True at the start of the sentence. Limit response to <50 words."

    for values in data:
        print(values)
        name, pos, neg = values
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "user", "content": f"{prompt}: {value}, {pos}, {neg}"}
            ]
        )
        outputs.append(response.choices[0].message.content)
        
    return outputs

# operating force
def operating_force_checks(operating_force_data):
    # check the init and maintaining force for open and close for inconsistencies
    prompt = "The following are in the form: average initiating force, average maintaining force. Are there any inconcsistencies in the data provided?"
    prompt += " Limit your response to <25 words, true or false for inconsistency existence and what the inconsistency is."
    # print(oFData)
    open, close = operating_force_data
    for data in [open, close]:
        init, maint = data[:2]
        # print(init, maint)
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "user", "content": f"{prompt}\n{init}, {maint}"}
            ])
        output = response.choices[0].message.content
        return output

# air infiltration
def air_infiltration_checks(air_infiltration_data):
    prompt = "result for an air infiltration test at +-75Pa. state true or false for any cause of concern and describe any potential causes for concern for this data? "
    prompt += "ignore 0 rating if it is for negative pressure test limit response to <20 words"
    prompt += str(air_infiltration_data)
    response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}])
    output = response.choices[0].message.content
    return output

# water penetration
def water_penetration_checks(water_data):
    waterVal, waterCom, _ = water_data
    prompt = "I am providing the results of a water penetration test, numbers in the comments correspond to test pressure"
    prompt += f"Final Test Pressure: {waterVal}. "
    prompt += f"Test Comments {waterCom}. "
    prompt += "state true or false if there are any causes of concern or inconsistencies. State those concerns in <50 words"
    response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}])
    output = response.choices[0].message.content
    
    return output

# ultimate strength
def ultimate_strength_checks(ultimate_data):
    prompt = "Limit response to <50 words. The following are positive and negative pressure readings for an ultimate strength test. the Y/N values correspond to checks performed during the test. If any of them are Y, the specimen fails. state true or false if there is any inconsistency or cause for concern."
    for data in ultimate_data:
        if isinstance(data, list):
            for item in data:
                prompt += item
        else:
            prompt += data
    response = client.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}])
    output = response.choices[0].message.content
    return output

