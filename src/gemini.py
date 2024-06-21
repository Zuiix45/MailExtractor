import google.generativeai as genai
from google.generativeai.types import HarmCategory, HarmBlockThreshold

__ai_model = {}

def connectToGemini(google_api_key: str, model_name: str = "gemini-1.5-flash"):
    genai.configure(api_key=google_api_key)
    
    safety_settings = {
        HarmCategory.HARM_CATEGORY_HARASSMENT: HarmBlockThreshold.BLOCK_NONE,
        HarmCategory.HARM_CATEGORY_HATE_SPEECH: HarmBlockThreshold.BLOCK_NONE,
        HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT: HarmBlockThreshold.BLOCK_NONE,
        HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT: HarmBlockThreshold.BLOCK_NONE,
    }
    
    __ai_model['model'] = genai.GenerativeModel(model_name=model_name, safety_settings=safety_settings)
    __ai_model['system_instructions'] = ""
    __ai_model['generation_config'] = genai.GenerationConfig(
        max_output_tokens=2048,
        temperature=0.5,
        top_p=0.9,
        top_k=40
    )
    
    print(f"{model_name} connected successfully...")

def setSystemInstructions(system_instructions: str):
    __ai_model['system_instructions'] = system_instructions

def setGenerationConfig(generation_config: dict):
    __ai_model['generation_config'] = generation_config

def generateContent(prompt: str, images = []):
    #print(f"Generating content for prompt: {prompt}...")
    
    content = [
        "System Instructions: " + __ai_model['system_instructions'],
        "Prompt: " + prompt + " "
    ]
        
    for image in images:
        content.append(image)
    
    return __ai_model['model'].generate_content(content, generation_config=__ai_model['generation_config']).text
