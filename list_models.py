from google import genai

client = genai.Client(api_key="API_KEY")
models = client.models.list()

for m in models:
    print("-", m.name)
