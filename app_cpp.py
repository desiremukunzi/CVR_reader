import os
import whisper
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for
from werkzeug.utils import secure_filename
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import Font

# Flask setup
app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
EXCEL_FILE = "workbench/WORKBENCH_BUR.xlsm"
#SHEET_NAME = "Checklist"
CHECKED_COLUMN = "B"
TRANSCRIPT_FOLDER = "transcripts"

# Ensure trascript folder exists
os.makedirs(TRANSCRIPT_FOLDER, exist_ok=True)
# Ensure upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Load checklist from Excel
def load_checklist(sheet_name):
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, engine="openpyxl")
    return df, df.iloc[:, 0].dropna().tolist()

# Check compliance with fuzzy matching
from statistics import mean

def check_compliance(transcript, checklist, threshold=50):
    transcript_lower = transcript.lower()
    sentences = [s.strip() for s in transcript_lower.split('.') if s.strip()]
    results = []

    for step in checklist:
        step_lower = step.lower()
        best_score = 0
        best_chunk = ""

        for i in range(len(sentences)):
            for window in range(1, 4):  # Window size: 1 to 3 sentences
                chunk = ' '.join(sentences[i:i + window])
                if not chunk:
                    continue

                # Calculate blended fuzzy score
                pr = fuzz.partial_ratio(step_lower, chunk)
                tsr = fuzz.token_set_ratio(step_lower, chunk)
                ratio = fuzz.ratio(step_lower, chunk)
                score = mean([pr, tsr, ratio])

                if score > best_score:
                    best_score = score
                    best_chunk = chunk

        # Optional: avoid false 100% if not exact match
        if best_score == 100.0 and step_lower not in transcript_lower:
            best_score = 99.0

        # Logging match info to console
        print(f"\n✅ Checklist Item: {step}")
        print(f"   🔍 Matched Transcript Chunk: \"{best_chunk}\"")
        print(f"   🎯 Accuracy Score: {best_score:.1f}%")

        results.append(("PASS" if best_score >= threshold else "FAIL", step, best_score))

    return results


# Update Excel with results
def update_excel(results, sheet_name):
    wb = load_workbook(EXCEL_FILE, keep_vba=True)
    ws = wb[sheet_name]
    
    # Your logic for updating the Excel sheet goes here


    row = 2  # Start from row 10
    for result in results:
        status_icon = "✔" if result[0] == "PASS" else "✘"
        cell = ws[f"{CHECKED_COLUMN}{row}"]
        cell.value = status_icon
        cell.font = Font(color="008000" if result[0] == "PASS" else "FF0000")
        row += 1

    wb.save(EXCEL_FILE)
import subprocess
import os
import shutil

TRANSCRIPT_FOLDER = "transcripts"
WHISPER_CPP_DIR = os.path.abspath("../whisper.cpp")
WHISPER_MODEL_PATH = os.path.join(WHISPER_CPP_DIR, "models", "ggml-medium.en.bin")
WHISPER_EXECUTABLE = os.path.join(WHISPER_CPP_DIR, "build", "bin", "main.exe")

import subprocess
import os

def transcribe_audio(audio_path):
    WHISPER_CLI_PATH = r"A:\whisper.cpp\build\bin\whisper-cli.exe"
    MODEL_PATH = r"A:\whisper.cpp\models\ggml-medium.en.bin"
    THREADS = "4"
    TRANSCRIPTS_DIR = "transcripts"

    # Ensure the transcripts folder exists
    os.makedirs(TRANSCRIPTS_DIR, exist_ok=True)

    # Generate output filename and full output path
    base_filename = os.path.splitext(os.path.basename(audio_path))[0]
    output_file_path = os.path.join(TRANSCRIPTS_DIR, base_filename)

    # Whisper CLI command with full output path
    command = [
        WHISPER_CLI_PATH,
        "--model", MODEL_PATH,
        "--file", audio_path,
        "--output-txt",
        "--output-file", output_file_path,
        "--threads", THREADS
    ]

    try:
        subprocess.run(command, check=True)
    except subprocess.CalledProcessError as e:
        print(f"❌ Whisper CLI failed: {e}")
        return ""

    transcript_path = output_file_path + ".txt"

    # Read the generated transcript
    try:
        with open(transcript_path, "r", encoding="utf-8") as f:
            transcript = f.read()
    except FileNotFoundError:
        print("❌ Transcription file not found at:", transcript_path)
        return ""

    return transcript



# Routes
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "file" not in request.files:
            return redirect(request.url)

        file = request.files["file"]
        threshold = int(request.form.get("threshold", 50))  # default is 50
        sheet_name = request.form.get("sheet_name")

        if file.filename == "":
         return redirect(request.url)

        if file:
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(file_path)

             # Process audio
            df, checklist = load_checklist(sheet_name)
            transcript = transcribe_audio(file_path)

            #print("\n--- Transcript ---")
            #print(transcript)

            results = check_compliance(transcript, checklist, threshold)

            # Update Excel
            update_excel(results, sheet_name)

            return render_template("report.html", results=results, transcript=transcript)

    return render_template("index.html")

# Run Flask app
if __name__ == "__main__":
    app.run(debug=True)
