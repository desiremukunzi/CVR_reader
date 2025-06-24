import os
#import whisper
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
EXCEL_FILE = "workbench/WORKBENCH_CARE.xlsm"
#SHEET_NAME = "Checklist"
CHECKED_COLUMN = "B"
TRANSCRIPT_FOLDER = "transcripts"

# Ensure trascript folder exists
os.makedirs(TRANSCRIPT_FOLDER, exist_ok=True)
# Ensure upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

import subprocess

def preprocess_audio(input_path, output_folder="uploads", filename_prefix="cleaned_"):
    """
    Converts, trims silence, normalizes volume, and prepares audio for transcription.
    Returns path to the cleaned WAV file.
    """
    os.makedirs(output_folder, exist_ok=True)

    base = os.path.basename(input_path)
    name, _ = os.path.splitext(base)
    output_path = os.path.join(output_folder, f"{filename_prefix}{name}.wav")

    command = [
        "ffmpeg", "-y",
        "-i", input_path,
        "-ac", "1",
        "-ar", "16000",
        "-af", "silenceremove=1:0:-50dB,loudnorm",
        "-c:a", "pcm_s16le",
        output_path
    ]

    try:
        subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        return output_path
    except subprocess.CalledProcessError as e:
        print(f"FFmpeg failed: {e}")
        return None


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

    

from faster_whisper import WhisperModel

# Initialize once globally (to avoid reloading for every request)
model = WhisperModel("medium", device="cuda", compute_type="float16")  # OR "medium" if GPU RAM is low

def transcribe_audio(audio_path):
    segments, info = model.transcribe(audio_path, language="en")

    transcript_text = " ".join([segment.text for segment in segments])

    # Save transcript
    base_filename = os.path.splitext(os.path.basename(audio_path))[0]
    transcript_filename = f"{base_filename}.txt"
    transcript_path = os.path.join(TRANSCRIPT_FOLDER, transcript_filename)

    with open(transcript_path, "w", encoding="utf-8") as f:
        f.write(transcript_text)

    return transcript_text


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

            # 🔧 Preprocess the audio file
            cleaned_path = preprocess_audio(file_path)

            if not cleaned_path:
                return "Audio preprocessing failed", 500

            # 🎧 Transcribe cleaned audio
            df, checklist = load_checklist(sheet_name)
            transcript = transcribe_audio(cleaned_path)


            #print("\n--- Transcript ---")
            #print(transcript)

            results = check_compliance(transcript, checklist, threshold)

            # Update Excel
            update_excel(results, sheet_name)
            
            passed_count = sum(1 for r in results if r[0] == "PASS")
            compliance_percent = round((passed_count / len(results)) * 100, 1)
            return render_template("report.html", results=results, transcript=transcript, compliance=compliance_percent)

           # return render_template("report.html", results=results, transcript=transcript)

    return render_template("index.html")

# Run Flask app
if __name__ == "__main__":
    app.run(debug=True)