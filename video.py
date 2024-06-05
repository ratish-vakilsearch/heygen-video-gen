import requests
from prompt import generate_video_script
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import os
from dotenv import load_dotenv

load_dotenv("keys.env")

def gen_video():

    heygen_api_key = os.getenv("heygen_api_key")

    # Ensure your API key is correct
    if not heygen_api_key:
        raise ValueError("Invalid Heygen API key. Please provide a valid key.")

    # Script content for the video
    script = generate_video_script()

    # Define the data for video creation
    video_data = {
        "video_inputs": [
            {
                "character": {
                    "type": "avatar",
                    "avatar_id": os.getenv("avatar_id"),
                    "avatar_style": "normal"
                },
                "voice": {
                    "type": "text",
                    "input_text": script,
                    "voice_id": os.getenv("voice_id"),
                    "speed": 1
                }
            }
        ],
        "dimension": {
            "width": 1280,
            "height": 720
        },
        "aspect_ratio": "16:9",
        "test": True
    }

    # Headers with the API key
    headers = {
        'X-Api-Key': heygen_api_key,
        'Content-Type': 'application/json'
    }

    # Send request to Heygen API
    response = requests.post(
        'https://api.heygen.com/v2/video/generate',
        headers=headers,
        json=video_data
    )

    # Check if the request was successful
    if response.status_code == 200:
        try:
            # Try to parse JSON response
            response_json = response.json()
            print(f"Response JSON: {response_json}")  # Print the full JSON response for debugging
            video_url = response_json.get('video_url')
            print(f"Video URL: {video_url}")

            if video_url:
                # Download and save the video file locally
                video_response = requests.get(video_url)
                local_video_path = 'financial_video.mp4'
                with open(local_video_path, 'wb') as f:
                    f.write(video_response.content)
                print(f"Video saved to {local_video_path}")

                # Authenticate and create the Google Drive service
                credentials = service_account.Credentials.from_service_account_file(
                    'path/to/your/credentials.json',
                    scopes=['https://www.googleapis.com/auth/drive']
                )
                service = build('drive', 'v3', credentials=credentials)

                # Upload the video to Google Drive
                file_metadata = {'name': 'financial_video.mp4'}
                media = MediaFileUpload(local_video_path, mimetype='video/mp4')
                file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                print(f"File ID: {file.get('id')}")

            else:
                print("Video URL not found in the response.")
        except ValueError:
            # If JSON parsing fails, print the response content for debugging
            print("Failed to parse JSON response:")
            print(response.content)
    else:
        # Print the status code and response content for debugging
        print(f"Request failed with status code {response.status_code}")
        print("Response content:")
        print(response.content)
        if response.status_code == 401:
            print("Unauthorized error. Please check your API key and permissions.")
        elif response.status_code == 429:
            print("Rate limit exceeded. Please try again later.")

gen_video()