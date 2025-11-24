from google import genai

client = genai.Client(api_key="AIzaSyDlC9SOyIPQU3ridllTZF3AB98Q9yxiW9U")
models = client.models.list()

for m in models:
    print("-", m.name)
