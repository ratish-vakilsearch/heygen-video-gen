from fastapi import FastAPI, File, UploadFile
import json
import uvicorn
import os
from extract import get_data

app = FastAPI()
UPLOAD_DIRECTORY = "uploaded_files"

@app.post("/upload_excel/")
async def upload_excel(file: UploadFile = File(...)):
    file_path = os.path.join(UPLOAD_DIRECTORY, file.filename)
    
    # Create the directory if it doesn't exist
    if not os.path.exists(UPLOAD_DIRECTORY):
        os.makedirs(UPLOAD_DIRECTORY)
    
    with open(file_path, "wb") as buffer:
        buffer.write(await file.read())
    
    result = get_data(file_path)

    os.remove(file_path)

    return json.loads(result)

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
