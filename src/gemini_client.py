from google import genai
from google.genai import types
import os
import json
from dotenv import load_dotenv

load_dotenv()


client = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))
MODEL_NAME = "gemini-2.5-flash"

# Define a grounding tool to enable Google Search
grounding_tool = types.Tool(
    google_search=types.GoogleSearch()
)

# Configure the generation to use the grounding tool
generation_config = types.GenerateContentConfig(
    tools=[grounding_tool]
)

def generate_slides(prompt: str):
    try:
        response = client.models.generate_content(
            model=MODEL_NAME,
            contents=prompt,
            config=generation_config
        )
        # Clean up the response to get raw Python code
        code = response.text.strip()
        if code.startswith("```python"):
            code = code[9:]
        if code.endswith("```"):
            code = code[:-3]
        return code.strip()
    except Exception as e:
        raise RuntimeError(f"Gemini API call error: {e}")
