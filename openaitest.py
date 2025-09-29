from openai import OpenAI
from config import apikey

# Initialize client with your key
client = OpenAI(api_key=apikey)

response = client.chat.completions.create(
    model="gpt-4o",   # use gpt-4o unless you really have gpt-5 access
    messages=[
        {"role": "system", "content": "You are a helpful assistant."},
        {"role": "user", "content": "Write a short resignation email addressed to Ms. Smith."}
    ]
)

print(response.choices[0].message.content)
