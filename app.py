import os
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_from_directory, url_for
from werkzeug.utils import secure_filename
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
import subprocess
from statistics import mean
import re
from faster_whisper import WhisperModel
import tempfile
import shutil
from datetime import datetime

# Flask setup
app = Flask(__name__)

# Define these variables FIRST
UPLOAD_FOLDER = "uploads" # For concatenated audio
COMPLIANCE_EXCEL_OUTPUT = "compliance_excel_output" # New folder for modified Excel files
CHECKED_COLUMN = "B"
TRANSCRIPT_FOLDER = "transcripts"
COMPLIANCE_TEXT_REPORTS_FOLDER = "compliance_text_reports" # Renamed for clarity from COMPLIANCE_FOLDER

# THEN assign them to app.config
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER # Good practice to also add this one
app.config['COMPLIANCE_EXCEL_OUTPUT'] = COMPLIANCE_EXCEL_OUTPUT
app.config['TRANSCRIPT_FOLDER'] = TRANSCRIPT_FOLDER # And this one
app.config['COMPLIANCE_TEXT_REPORTS_FOLDER'] = COMPLIANCE_TEXT_REPORTS_FOLDER # And this one too

# Ensure necessary folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['COMPLIANCE_EXCEL_OUTPUT'], exist_ok=True)
os.makedirs("workbench", exist_ok=True) # Still optional if not used for other assets
os.makedirs(app.config['TRANSCRIPT_FOLDER'], exist_ok=True)
os.makedirs(app.config['COMPLIANCE_TEXT_REPORTS_FOLDER'], exist_ok=True)

# ... (rest of your code)

# Initialize Whisper model globally (to avoid reloading for every request)
model = WhisperModel("medium", device="cuda", compute_type="float16")

#normal audio preprocessing function
def preprocess_audio(input_path):
    # All files are .wav already, assume they’re good
   print(f"Skipping preprocessing. Using original WAV: {input_path}")
   return input_path

# #The below to be used if you want to preprocess audio files when the transcript generated is ------- ,okay okay,etc (broken)
# def preprocess_audio(input_path, output_folder="uploads", filename_prefix="cleaned_"):
#     """
#     Converts, trims silence, normalizes volume, and prepares audio for transcription.
#     Returns path to the cleaned WAV file.
#     """
#     os.makedirs(output_folder, exist_ok=True)

#     base = os.path.basename(input_path)
#     name, _ = os.path.splitext(base)
#     output_path = os.path.join(output_folder, f"{filename_prefix}{name}.wav")

#     command = [
#         "ffmpeg", "-y",
#         "-i", input_path,
#         "-ac", "1",
#         "-ar", "16000",
#         "-af", "silenceremove=1:0:-50dB,loudnorm",
#         "-c:a", "pcm_s16le",
#         output_path
#     ]

#     try:
#         # Capture stderr to see FFmpeg's error messages
#         result = subprocess.run(command, capture_output=True, text=True, check=True)
#         print(f"FFmpeg stdout (preprocessing): {result.stdout}")
#         return output_path
#     except subprocess.CalledProcessError as e:
#         print(f"FFmpeg failed during preprocessing: {e}")
#         print(f"FFmpeg stderr (preprocessing): {e.stderr}")
#         return None

def concatenate_audio_files(input_paths, output_filename, upload_folder):
    """
    Concatenates multiple audio files into a single WAV file using ffmpeg.
    Returns the path to the concatenated file.
    """
    concat_list_path = os.path.join(tempfile.gettempdir(), "files.txt")

    with open(concat_list_path, "w") as f:
        for path in input_paths:
            f.write(f"file '{path}'\n")

    # Ensure output_filename has a .wav extension
    if not output_filename.lower().endswith(".wav"):
        output_filename += ".wav"

    output_path = os.path.join(upload_folder, secure_filename(output_filename))

    command = [
        "ffmpeg", "-y",
        "-f", "concat",
        "-safe", "0", # Required for external file paths in concat list
        "-i", concat_list_path,
        "-c", "copy", # Use stream copy for maximum speed if codecs are compatible
        output_path
    ]

    try:
        result = subprocess.run(command, capture_output=True, text=True, check=True)
        print(f"FFmpeg stdout (concatenation): {result.stdout}")
        return output_path
    except subprocess.CalledProcessError as e:
        print(f"FFmpeg failed during concatenation: {e}")
        print(f"FFmpeg stderr (concatenation): {e.stderr}")
        return None
    finally:
        # Clean up the temporary concat list file
        if os.path.exists(concat_list_path):
            os.remove(concat_list_path)

def load_checklist(excel_file_path, sheet_name):
    """
    Loads the checklist items from the specified Excel sheet in the given file path.
    Assumes checklist items are in the first column.
    """
    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, engine="openpyxl")
        return df, df.iloc[:, 0].dropna().tolist()
    except Exception as e:
        raise Exception(f"Failed to load checklist from Excel: {e}. Check sheet name or file integrity.")

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
    transcript_words_raw = transcript_lower.split()

    results = []
    MAX_CHUNK_WORDS = 20

    for step in checklist:
        step_clean = clean_text(step)
        best_score = 0
        best_chunk_raw = ""

        for i in range(len(transcript_words_raw)):
            for j in range(i + 1, min(i + MAX_CHUNK_WORDS + 1, len(transcript_words_raw) + 1)):
                current_chunk_words_raw = transcript_words_raw[i:j]
                current_chunk_raw = ' '.join(current_chunk_words_raw)

                current_chunk_clean = clean_text(current_chunk_raw)

                if not current_chunk_clean:
                    continue

                pr = fuzz.partial_ratio(step_clean, current_chunk_clean)
                tsr = fuzz.token_set_ratio(step_clean, current_chunk_clean)
                ratio = fuzz.ratio(step_clean, current_chunk_clean)

                #tested on 90 threshold less rigorous
                score = max(pr, tsr, ratio) * 0.6 + mean([pr, tsr, ratio]) * 0.4

                if score > best_score:
                    best_score = score
                    best_chunk_raw = current_chunk_raw

        if best_score == 100.0 and step_clean not in clean_text(transcript):
            best_score = 99.0

        # Print to console (for backend debugging/logging)
        print(f"\n✅ Checklist Item: {step}")
        print(f"   🔍 Matched: \"{best_chunk_raw}\"")
        print(f"   🎯 Score: {best_score:.1f}%")

        results.append(("PASS" if best_score >= threshold else "FAIL", step, best_score, best_chunk_raw)) # Added best_chunk_raw

    return results

def update_excel(excel_input_path, results, sheet_name, not_complied_count, compliance_percent):
    """
    Updates the Excel file with compliance results and saves it to a new path
    in COMPLIANCE_EXCEL_OUTPUT.
    """
    try:
        # Load the workbook from the specified input path, keeping VBA macros
        wb = load_workbook(excel_input_path, keep_vba=True)
        
        # Get the target sheet
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found in the uploaded Excel file.")
        ws = wb[sheet_name]

        # Update checklist results
        # Assuming the checklist items start from the second row (row index 2 in Excel)
        row = 2
        for result in results:
            status_icon = "✔" if result[0] == "PASS" else "✘"
            cell = ws[f"{CHECKED_COLUMN}{row}"]
            cell.value = status_icon
            cell.font = Font(color="008000" if result[0] == "PASS" else "FF0000")
            row += 1

        # --- New and Modified Formatting for Checklist Compliance Summary ---

        # Get or create the "Summary" sheet
        if "Summary" not in wb.sheetnames:
            summary_ws = wb.create_sheet("Summary")
        else:
            summary_ws = wb["Summary"]

        # Define common styles
        bold_font_white = Font(bold=True, color="FFFFFF") # White color for title
        blue_background = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid") # RGB(68, 114, 196)
        thin_border_side = Side(style='thick') # This defines the side style for the thick border

        # 1. Add title "Checklist Compliance"
        # The title will now start at E8
        summary_ws['E8'].value = "Checklist Compliance"
        summary_ws['E8'].font = bold_font_white
        summary_ws['E8'].fill = blue_background
        summary_ws['E8'].alignment = Alignment(horizontal='center', vertical='center')
        summary_ws.merge_cells('E8:F8')

        # Set row 8 height
        summary_ws.row_dimensions[8].height = 24

        # Set column E width (and F for balance)
        summary_ws.column_dimensions['E'].width = 20
        summary_ws.column_dimensions['F'].width = 10 # Give some width to F as well for values

        # 2. Apply professional formatting for "Checks Not Complied"
        # These will now start at row 9
        summary_ws['E9'].value = "Checks Not Complied:"
        summary_ws['E9'].font = Font(bold=True)
        summary_ws['F9'].value = not_complied_count
        summary_ws['F9'].font = Font(bold=True, color="FF0000") # Red color for not complied count
        summary_ws['F9'].fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # Light orange background
        summary_ws['F9'].alignment = Alignment(horizontal='center', vertical='center')

        # 3. Apply professional formatting for "Complied Percentage"
        # These will now start at row 10
        summary_ws['E10'].value = "Complied Percentage:"
        summary_ws['E10'].font = Font(bold=True)
        summary_ws['F10'].value = f"{compliance_percent:.1f}%"
        summary_ws['F10'].font = Font(bold=True, color="008000") # Green color for complied percentage
        summary_ws['F10'].fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid") # Light green background
        summary_ws['F10'].alignment = Alignment(horizontal='center', vertical='center')

        # 4. Apply thick border around the compliance summary table (now E8:F10)
        # The merged cell for the title is E8:F8.
        summary_ws['E8'].border = Border(left=thin_border_side, right=thin_border_side)
        summary_ws['E9'].border = Border(left=thin_border_side)
        summary_ws['F9'].border = Border(right=thin_border_side)
        summary_ws['E10'].border = Border(bottom=thin_border_side, left=thin_border_side)
        summary_ws['F10'].border = Border(bottom=thin_border_side, right=thin_border_side)

        # Generate a unique filename for the output Excel file
        base_name = os.path.splitext(os.path.basename(excel_input_path))[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_excel_filename = f"{base_name}.xlsm"
        output_excel_path = os.path.join(COMPLIANCE_EXCEL_OUTPUT, output_excel_filename)

        # Save the workbook to the new, dedicated output folder
        wb.save(output_excel_path)
        print(f"Updated Excel file saved to: {output_excel_path}")
        return output_excel_path # Return the path where the updated Excel is saved

    except Exception as e:
        print(f"Error updating Excel file: {e}")
        raise Exception(f"Failed to update Excel file: {e}. Please ensure the Excel file is not open and accessible and has the correct sheet name.")


def transcribe_audio(audio_path, custom_name=None):
    segments, info = model.transcribe(audio_path, language="en")
    transcript_text = " ".join([segment.text for segment in segments])

    if custom_name:
        base_filename = os.path.splitext(secure_filename(custom_name))[0]
    else:
        base_filename = os.path.splitext(os.path.basename(audio_path))[0]

    transcript_filename = f"{base_filename}.txt"
    transcript_path = os.path.join(TRANSCRIPT_FOLDER, transcript_filename)

    with open(transcript_path, "w", encoding="utf-8") as f:
        f.write(transcript_text)

    print(f"Transcript saved to: {transcript_path}")
    return transcript_text


def save_compliance_report(results, output_file_name):
    """
    Saves the compliance results to a text file in the compliance folder.
    """
    # Sanitize output_file_name for use in filename, remove extension if present
    base_name = os.path.splitext(secure_filename(output_file_name))[0]

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_filename = f"{base_name}_compliance_report_{timestamp}.txt"
    report_path = os.path.join(COMPLIANCE_TEXT_REPORTS_FOLDER, report_filename) # Use new folder name

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(f"Compliance Report for: {output_file_name}\n")
        f.write(f"Generated On: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("-" * 50 + "\n\n")

        for status, checklist_item, score, matched_text in results:
            f.write(f"Status: {status}\n")
            f.write(f"Checklist Item: {checklist_item}\n")
            f.write(f"Matched Text: \"{matched_text}\"\n") # Include matched text
            f.write(f"Score: {score:.1f}%\n")
            f.write("-" * 20 + "\n")
    print(f"Compliance report saved to: {report_path}")


# New route to serve the updated Excel file for download
@app.route('/download_updated_excel/<filename>', methods=['GET'])
def download_updated_excel(filename):
    """
    Allows users to download the modified Excel file.
    """
    # Ensure the filename is secure to prevent directory traversal attacks
    secure_filename_download = secure_filename(filename)
    return send_from_directory(
        directory=app.config['COMPLIANCE_EXCEL_OUTPUT'], # Serve from the dedicated output folder
        path=secure_filename_download,
        as_attachment=True # Forces download instead of opening in browser
    )


# Main route
@app.route("/", methods=["GET", "POST"])
def index():
    """
    Handles file upload, audio preprocessing, transcription, compliance checking,
    and returns the report as JSON.
    """
    if request.method == "POST":
        # Create a temporary directory for all uploads for this request
        temp_request_dir = tempfile.mkdtemp()
        excel_file_path = None # Path to the uploaded Excel file (temporary)
        final_excel_output_path = None # Path to the *modified* Excel file (in COMPLIANCE_EXCEL_OUTPUT)
        concatenated_audio_path = None
        cleaned_audio_path = None

        try:
            # 1. Handle Excel file upload
            if 'excel_file' not in request.files:
                raise ValueError("No Excel file part in the request.")
            
            excel_file_upload = request.files['excel_file']
            if excel_file_upload.filename == '':
                raise ValueError("No selected Excel file.")
            
            # Secure filename and save the Excel file to the temporary directory
            excel_filename = secure_filename(excel_file_upload.filename)
            excel_file_path = os.path.join(temp_request_dir, excel_filename)
            excel_file_upload.save(excel_file_path)
            print(f"Excel file saved temporarily for processing at: {excel_file_path}")

            # 2. Handle Audio file(s) upload
            if 'audio_files[]' not in request.files:
                raise ValueError("No audio files part in the request.")
            
            uploaded_audio_files = request.files.getlist('audio_files[]')
            if not uploaded_audio_files or uploaded_audio_files[0].filename == '':
                raise ValueError("No audio files selected.")

            saved_audio_paths = []
            for file in uploaded_audio_files:
                if file:
                    audio_filename = secure_filename(file.filename)
                    # Save audio files also to the temporary directory
                    audio_file_path = os.path.join(temp_request_dir, audio_filename)
                    file.save(audio_file_path)
                    saved_audio_paths.append(audio_file_path)

            if not saved_audio_paths:
                raise ValueError("No valid audio files uploaded.")

            # Get other form data
            output_file_name = request.form.get("output_file_name", "concatenated_audio.wav")
            threshold = int(request.form.get("threshold", 50))
            sheet_name = request.form.get("sheet_name")

            if not sheet_name:
                raise ValueError("Sheet name is required.")
            if not output_file_name:
                raise ValueError("Output file name is required.")

            # 3. Audio Processing (Concatenation, Preprocessing, Transcription)
            if len(saved_audio_paths) == 1:
                concatenated_audio_path = saved_audio_paths[0]
                print(f"Skipping concatenation: Directly processing single audio file: {os.path.basename(concatenated_audio_path)}")
            else:
                concatenated_audio_path = concatenate_audio_files(saved_audio_paths, output_file_name, app.config["UPLOAD_FOLDER"])
                if not concatenated_audio_path:
                    raise Exception("Audio concatenation failed. Check FFmpeg installation and file formats.")
                print(f"Concatenated audio saved to: {os.path.basename(concatenated_audio_path)}")

            cleaned_audio_path = preprocess_audio(concatenated_audio_path)
            if not cleaned_audio_path:
                raise Exception("Audio preprocessing failed.")

            transcript = transcribe_audio(cleaned_audio_path, output_file_name)

            # 4. Load Checklist and Check Compliance
            df, checklist = load_checklist(excel_file_path, sheet_name) # Pass the uploaded excel_file_path
            results = check_compliance(transcript, checklist, threshold)

            # Calculate compliance statistics
            passed_count = sum(1 for r in results if r[0] == "PASS")
            total_checks = len(results)
            compliance_percent = round((passed_count / total_checks) * 100, 1) if total_checks else 0
            not_complied_count = total_checks - passed_count

            # 5. Update the *uploaded* Excel file and get the path to the saved modified version
            final_excel_output_path = update_excel(excel_file_path, results, sheet_name, not_complied_count, compliance_percent)

            # 6. Save compliance report to a separate text file
            save_compliance_report(results, output_file_name)

            # 7. Return results as JSON for frontend, including the download URL for the updated Excel
            updated_excel_filename = os.path.basename(final_excel_output_path)
            download_url = url_for('download_updated_excel', filename=updated_excel_filename, _external=True)

            return jsonify({
                "results": results,
                "compliance_percent": compliance_percent,
                "not_complied_count": not_complied_count,
                "excel_updated": True,
                "updated_excel_filename": updated_excel_filename,
                "download_excel_url": download_url # NEW: URL to download the updated Excel file
            })

        except Exception as e:
            print(f"An error occurred: {e}")
            return jsonify({"error": f"Error processing files: {e}"}), 500
        finally:
            # Clean up: remove the temporary directory and all its contents
            if os.path.exists(temp_request_dir):
                shutil.rmtree(temp_request_dir)
            # The concatenated_audio_path might be in UPLOAD_FOLDER (which is app.config["UPLOAD_FOLDER"]),
            # and it's up to you if you want to clear that folder after some time or keep the files.
            # For now, it's not deleted here as it's outside temp_request_dir.
            # cleaned_audio_path is usually the same as concatenated_audio_path or a temp copy that should be handled by tempfile.
            # If `preprocess_audio` creates a new file, ensure it's in temp_request_dir or explicitly deleted.

    return render_template("index.html")

# Run Flask app
if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)