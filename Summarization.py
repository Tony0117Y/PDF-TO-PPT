import openai
import os

openai.api_key = "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"

def gpt_summarise(text):
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo-0301",
        messages=[
            {
                "role": "system",
                "content": "Summarize the following paragraph into a 5 bullet points"
            },
            {
                "role": "user",
                "content": text
            }
        ],
        temperature=0.5,
        max_tokens=400,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0,
    )
    return {"text": response['choices'][0]['message']['content']}