from fastapi import FastAPI, UploadFile, File
import tempfile
import json
import os

from schedule_to_json import parse_schedule

app = FastAPI()

@app.post("/parse-schedule")
async def parse_schedule_endpoint(file: UploadFile = File(...)):
    contents = await file.read()

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(contents)
        input_path = tmp.name

    output_path = input_path + ".json"

    try:
        records = parse_schedule(input_path, output_path)

        with open(output_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        return {
            "success": True,
            "record_count": len(data),
            "data": data
        }

    finally:
        if os.path.exists(input_path):
            os.remove(input_path)
        if os.path.exists(output_path):
            os.remove(output_path)
