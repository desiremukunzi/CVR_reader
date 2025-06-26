import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for
from werkzeug.utils import secure_filename
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import Font
import subprocess
from statistics import mean
import re
from faster_whisper import WhisperModel

# Flask setup
app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
EXCEL_FILE = "workbench/WORKBENCH_CARE.xlsm" 
CHECKED_COLUMN = "B"
TRANSCRIPT_FOLDER = "transcripts"

# Ensure transcript folder exists
os.makedirs(TRANSCRIPT_FOLDER, exist_ok=True)
# Ensure upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
# Ensure workbench folder exists for the excel file
os.makedirs("workbench", exist_ok=True)


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
    """
    Loads the checklist items from the specified Excel sheet.
    Assumes checklist items are in the first column.
    """
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, engine="openpyxl")
    return df, df.iloc[:, 0].dropna().tolist()

def clean_text(text):
    """
    Cleans text by converting to lowercase, removing non-alphanumeric characters,
    and removing common filler words.
    """
    text = text.lower()
    text = re.sub(r"[^a-zA-Z0-9\s]", "", text)
    text = re.sub(r"\b(?:roger|copy|standby|okay|affirmative|negative|check)\b", "", text)
    return text.strip()


def check_compliance(transcript, checklist, threshold=50):
    """
    Checks compliance of the transcript against the checklist using fuzzy matching.
    This version uses a sliding window across the entire transcript (word-by-word)
    to find the best possible match for each checklist item.
    """
    transcript_lower = transcript.lower()
    # Split the entire transcript into words for flexible windowing,
    # rather than relying on sentence splitting.
    transcript_words_raw = transcript_lower.split()

    results = []
    # Define a reasonable maximum number of words for a chunk to match against.
    # Checklist items are typically phrases, so 20 words should be sufficient.
    MAX_CHUNK_WORDS = 20 

    for step in checklist:
        step_clean = clean_text(step) # Clean the checklist item
        best_score = 0
        best_chunk_raw = "" # This will store the original (lowercase) transcript chunk that yielded the best score

        # Iterate through all possible starting points in the transcript words
        for i in range(len(transcript_words_raw)):
            # Iterate through possible end points to create chunks of varying lengths.
            # The chunk length will range from 1 word up to MAX_CHUNK_WORDS.
            for j in range(i + 1, min(i + MAX_CHUNK_WORDS + 1, len(transcript_words_raw) + 1)):
                current_chunk_words_raw = transcript_words_raw[i:j]
                current_chunk_raw = ' '.join(current_chunk_words_raw) # Reconstruct the raw chunk string

                current_chunk_clean = clean_text(current_chunk_raw) # Clean this current chunk for comparison

                if not current_chunk_clean: # Skip if the cleaned chunk is empty
                    continue

                # Calculate different fuzzy matching scores using rapidfuzz
                pr = fuzz.partial_ratio(step_clean, current_chunk_clean)
                tsr = fuzz.token_set_ratio(step_clean, current_chunk_clean)
                ratio = fuzz.ratio(step_clean, current_chunk_clean)
                
                # Calculate a combined score using a weighted average.
                # Max score is given more weight, balanced by the mean of all scores.
                score = max(pr, tsr, ratio) * 0.6 + mean([pr, tsr, ratio]) * 0.4

                # If this chunk's score is better than the current best for this checklist item
                if score > best_score:
                    best_score = score
                    best_chunk_raw = current_chunk_raw # Store the raw chunk that produced the best score

        # Edge case: if a perfect match (100%) is found, but the cleaned checklist item
        # is not actually present as a direct substring in the entire cleaned transcript.
        # This helps prevent false positives for very short or common checklist items.
        if best_score == 100.0 and step_clean not in clean_text(transcript):
            best_score = 99.0

        # Print the results for the current checklist item to the console
        print(f"\n✅ Checklist Item: {step}")
        print(f"   🔍 Matched: \"{best_chunk_raw}\"") # Display the raw best match
        print(f"   🎯 Score: {best_score:.1f}%")

        # Determine PASS/FAIL status based on the threshold and add to results list
        results.append(("PASS" if best_score >= threshold else "FAIL", step, best_score))

    return results


def update_excel(results, sheet_name):
    """
    Updates the Excel file with compliance results.
    Marks cells in CHECKED_COLUMN with a checkmark (✔) for PASS and cross (✘) for FAIL,
    with corresponding green/red colors.
    """
    wb = load_workbook(EXCEL_FILE, keep_vba=True)
    ws = wb[sheet_name]
    
    # Starting row for results in Excel (assuming header is in row 1, checklist starts from row 2)
    row = 2
    for result in results:
        status_icon = "✔" if result[0] == "PASS" else "✘"
        cell = ws[f"{CHECKED_COLUMN}{row}"]
        cell.value = status_icon
        cell.font = Font(color="008000" if result[0] == "PASS" else "FF0000")
        row += 1

    wb.save(EXCEL_FILE)


# Initialize Whisper model globally (to avoid reloading for every request)
# Changed device to "cpu" for broader compatibility. If you have a CUDA GPU, you can change it back to "cuda".
model = WhisperModel("medium", device="cuda", compute_type="float16") 

def transcribe_audio(audio_path):
    """
    Transcribes audio using the FasterWhisper model and saves the transcript to a file.
    """
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
    """
    Handles file upload, audio preprocessing, transcription, compliance checking,
    and displays the report.
    """
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


            # Check compliance
            results = check_compliance(transcript, checklist, threshold)

            # Update Excel
            try:
                update_excel(results, sheet_name)
            except Exception as e:
                return f"Error updating Excel file: {e}. Please ensure the Excel file is not open and accessible.", 500
            
            passed_count = sum(1 for r in results if r[0] == "PASS")
            compliance_percent = round((passed_count / len(results)) * 100, 1) if results else 0

            return render_template("report.html", results=results, transcript=transcript, compliance=compliance_percent)

    return render_template("index.html")

# Run Flask app
if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)
