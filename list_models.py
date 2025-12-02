from google import genai

client = genai.Client(api_key="AIzaSyA4cRry1326Ax4ru2fhyqEKns8tix2aP7w")
models = client.models.list()

for m in models:
    print("-", m.name)
