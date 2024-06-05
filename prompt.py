import os
from extract import get_data
from openai import AzureOpenAI
from dotenv import load_dotenv

load_dotenv("keys.env")

def generate_video_script():
    
    client = AzureOpenAI(
    azure_endpoint = os.getenv("azure_endpoint"), 
    api_key= os.getenv("api_key"),  
    api_version=os.getenv("api_version")
    )

    response = client.chat.completions.create(
        model=os.getenv("model"), 
        messages=[
            {"role": "system", "content": "You are a financial analyst."},
            {"role": "user", "content": f"""
            Create a report for a financial report presentation. The data includes:
            {get_data('PISCIS Network Pvt Ltd Financials March 24 final (1).xlsx')}

            The script should be in such a way that it just has the content of what the reporter needs to talk, not more than that. Strictly give the report within 1000 characters.
            """}
        ]
    )

    video_script = response.choices[0].message.content.strip()
    return video_script