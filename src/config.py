from openai import AzureOpenAI

client = AzureOpenAI(
        azure_endpoint = "", 
        api_key="",  
        api_version=""
        #api_version="2024-12-01-preview"
    )
MODEL_NAME = ""





PROMPT_MAX_CHARS = 370_000

USE_GPT_HEADER_REGION = False

USE_GPT_SHEET_SELECTION = False