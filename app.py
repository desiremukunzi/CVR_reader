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
from flight_analyzer import FlightAnalyzer

# Flask setup
app = Flask(__name__)

# Define folders
UPLOAD_FOLDER = "uploads"
COMPLIANCE_EXCEL_OUTPUT = "compliance_excel_output"
CHECKED_COLUMN = "B"
TRANSCRIPT_FOLDER = "transcripts"
COMPLIANCE_TEXT_REPORTS_FOLDER = "compliance_text_reports"
FLIGHT_DATA_FOLDER = "flight_data"

# Assign to app.config
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['COMPLIANCE_EXCEL_OUTPUT'] = COMPLIANCE_EXCEL_OUTPUT
app.config['TRANSCRIPT_FOLDER'] = TRANSCRIPT_FOLDER
app.config['COMPLIANCE_TEXT_REPORTS_FOLDER'] = COMPLIANCE_TEXT_REPORTS_FOLDER
app.config['FLIGHT_DATA_FOLDER'] = FLIGHT_DATA_FOLDER

# Ensure necessary folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['COMPLIANCE_EXCEL_OUTPUT'], exist_ok=True)
os.makedirs("workbench", exist_ok=True)
os.makedirs(app.config['TRANSCRIPT_FOLDER'], exist_ok=True)
os.makedirs(app.config['COMPLIANCE_TEXT_REPORTS_FOLDER'], exist_ok=True)
os.makedirs(app.config['FLIGHT_DATA_FOLDER'], exist_ok=True)

# Initialize Whisper model
model = WhisperModel("medium", device="cuda", compute_type="float16")

# Initialize Flight Analyzer
flight_analyzer = FlightAnalyzer(data_folder=app.config['FLIGHT_DATA_FOLDER'])

def preprocess_audio(input_path):
    """Skip preprocessing - use original WAV files."""
    print(f"Using original WAV: {input_path}")
    return input_path

def concatenate_audio_files(input_paths, output_filename, upload_folder):
    """Concatenate multiple audio files into a single WAV file."""
    concat_list_path = os.path.join(tempfile.gettempdir(), "files.txt")

    with open(concat_list_path, "w") as f:
        for path in input_paths:
            f.write(f"file '{path}'\n")

    if not output_filename.lower().endswith(".wav"):
        output_filename += ".wav"

    output_path = os.path.join(upload_folder, secure_filename(output_filename))

    command = [
        "ffmpeg", "-y",
        "-f", "concat",
        "-safe", "0",
        "-i", concat_list_path,
        "-c", "copy",
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
        if os.path.exists(concat_list_path):
            os.remove(concat_list_path)

def load_checklist(excel_file_path, sheet_name):
    """Load checklist items from Excel sheet."""
    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, engine="openpyxl")
        return df, df.iloc[:, 0].dropna().tolist()
    except Exception as e:
        raise Exception(f"Failed to load checklist from Excel: {e}")

def clean_text(text):
    """Clean text for fuzzy matching."""
    text = text.lower()
    text = re.sub(r"[^a-zA-Z0-9\s]", "", text)
    text = re.sub(r"\b(?:roger|copy|standby|okay|affirmative|negative|check)\b", "", text)
    return text.strip()

def check_compliance(transcript, checklist, threshold=50):
    """Check compliance using fuzzy matching with sliding window."""
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

                score = max(pr, tsr, ratio) * 0.6 + mean([pr, tsr, ratio]) * 0.4

                if score > best_score:
                    best_score = score
                    best_chunk_raw = current_chunk_raw

        if best_score == 100.0 and step_clean not in clean_text(transcript):
            best_score = 99.0

        print(f"\n✅ Checklist Item: {step}")
        print(f"   🔍 Matched: \"{best_chunk_raw}\"")
        print(f"   🎯 Score: {best_score:.1f}%")

        results.append(("PASS" if best_score >= threshold else "FAIL", step, best_score, best_chunk_raw))

    return results

def update_excel(excel_input_path, results, sheet_name, not_complied_count, compliance_percent):
    """Update Excel file with compliance results."""
    try:
        wb = load_workbook(excel_input_path, keep_vba=True)
        
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found in the uploaded Excel file.")
        ws = wb[sheet_name]

        # Update checklist results
        row = 2
        for result in results:
            status_icon = "✔" if result[0] == "PASS" else "✘"
            cell = ws[f"{CHECKED_COLUMN}{row}"]
            cell.value = status_icon
            cell.font = Font(color="008000" if result[0] == "PASS" else "FF0000")
            row += 1

        # Update Summary sheet
        if "Summary" not in wb.sheetnames:
            summary_ws = wb.create_sheet("Summary")
        else:
            summary_ws = wb["Summary"]

        bold_font_white = Font(bold=True, color="FFFFFF")
        blue_background = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        thin_border_side = Side(style='thick')

        summary_ws['E8'].value = "Checklist Compliance"
        summary_ws['E8'].font = bold_font_white
        summary_ws['E8'].fill = blue_background
        summary_ws['E8'].alignment = Alignment(horizontal='center', vertical='center')
        summary_ws.merge_cells('E8:F8')

        summary_ws.row_dimensions[8].height = 24
        summary_ws.column_dimensions['E'].width = 20
        summary_ws.column_dimensions['F'].width = 10

        summary_ws['E9'].value = "Checks Not Complied:"
        summary_ws['E9'].font = Font(bold=True)
        summary_ws['F9'].value = not_complied_count
        summary_ws['F9'].font = Font(bold=True, color="FF0000")
        summary_ws['F9'].fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        summary_ws['F9'].alignment = Alignment(horizontal='center', vertical='center')

        summary_ws['E10'].value = "Complied Percentage:"
        summary_ws['E10'].font = Font(bold=True)
        summary_ws['F10'].value = f"{compliance_percent:.1f}%"
        summary_ws['F10'].font = Font(bold=True, color="008000")
        summary_ws['F10'].fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
        summary_ws['F10'].alignment = Alignment(horizontal='center', vertical='center')

        summary_ws['E8'].border = Border(left=thin_border_side, right=thin_border_side)
        summary_ws['E9'].border = Border(left=thin_border_side)
        summary_ws['F9'].border = Border(right=thin_border_side)
        summary_ws['E10'].border = Border(bottom=thin_border_side, left=thin_border_side)
        summary_ws['F10'].border = Border(bottom=thin_border_side, right=thin_border_side)

        base_name = os.path.splitext(os.path.basename(excel_input_path))[0]
        output_excel_filename = f"{base_name}.xlsm"
        output_excel_path = os.path.join(COMPLIANCE_EXCEL_OUTPUT, output_excel_filename)

        wb.save(output_excel_path)
        print(f"Updated Excel file saved to: {output_excel_path}")
        return output_excel_path

    except Exception as e:
        print(f"Error updating Excel file: {e}")
        raise Exception(f"Failed to update Excel file: {e}")

def transcribe_audio(audio_path, custom_name=None):
    """Transcribe audio using Whisper."""
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
    """Save compliance results to text file."""
    base_name = os.path.splitext(secure_filename(output_file_name))[0]
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_filename = f"{base_name}_compliance_report_{timestamp}.txt"
    report_path = os.path.join(COMPLIANCE_TEXT_REPORTS_FOLDER, report_filename)

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(f"Compliance Report for: {output_file_name}\n")
        f.write(f"Generated On: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write("-" * 50 + "\n\n")

        for status, checklist_item, score, matched_text in results:
            f.write(f"Status: {status}\n")
            f.write(f"Checklist Item: {checklist_item}\n")
            f.write(f"Matched Text: \"{matched_text}\"\n")
            f.write(f"Score: {score:.1f}%\n")
            f.write("-" * 20 + "\n")
    print(f"Compliance report saved to: {report_path}")

@app.route("/analyze_flight_anomalies", methods=["POST"])
def analyze_flight_anomalies():
    """
    Endpoint to analyze flight data for anomalies using the "Clean Data" sheet
    from the compliance Excel file.
    After analysis, stores the report and returns success.
    """
    global latest_anomaly_report
    
    try:
        data = request.get_json()
        excel_filename = data.get('excel_filename')
        add_to_training = data.get('add_to_training', False)
        sheet_name = data.get('sheet_name', 'Clean Data')
        
        if not excel_filename:
            return jsonify({'error': 'No Excel filename provided'}), 400
        
        # The Excel file is in COMPLIANCE_EXCEL_OUTPUT from the compliance check
        excel_path = os.path.join(COMPLIANCE_EXCEL_OUTPUT, excel_filename)
        
        if not os.path.exists(excel_path):
            return jsonify({'error': f'Excel file not found: {excel_filename}'}), 404
        
        print(f"Analyzing flight data from: {excel_path}")
        print(f"Sheet name: {sheet_name}")
        print(f"Add to training: {add_to_training}")
        
        # Analyze the flight using "Clean Data" sheet
        results = flight_analyzer.analyze_flight(
            excel_path=excel_path,
            sheet_name=sheet_name,
            add_to_training=add_to_training
        )
        
        if 'error' in results:
            return jsonify({'error': results['error']}), 500
        
        # Store the report with timestamp
        results['analysis_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        latest_anomaly_report = results
        print(f"Stored anomaly report for Flight {results.get('flight_id', 'unknown')}")
        
        # Return success - frontend will redirect to /anomaly_report
        return jsonify({'success': True})
        
    except Exception as e:
        print(f"Error in flight anomaly analysis: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


# Store the latest anomaly report in memory (simple approach)
latest_anomaly_report = None

@app.route("/anomaly_report", methods=["GET"])
def view_anomaly_report():
    """
    Display the anomaly report in a dedicated web page with tabs.
    Data should already be stored by the analyze_flight_anomalies endpoint.
    """
    global latest_anomaly_report
    
    if not latest_anomaly_report:
        return "No anomaly report available. Please run an analysis first.", 404
    
    return render_template("anomaly_report.html", report=latest_anomaly_report)


@app.route("/", methods=["GET", "POST"])
def index():
    """Handle file upload and compliance checking."""
    if request.method == "POST":
        temp_request_dir = tempfile.mkdtemp()
        excel_file_path = None
        final_excel_output_path = None
        concatenated_audio_path = None
        cleaned_audio_path = None

        try:
            # Handle Excel file upload
            if 'excel_file' not in request.files:
                raise ValueError("No Excel file part in the request.")
            
            excel_file_upload = request.files['excel_file']
            if excel_file_upload.filename == '':
                raise ValueError("No selected Excel file.")
            
            excel_filename = secure_filename(excel_file_upload.filename)
            excel_file_path = os.path.join(temp_request_dir, excel_filename)
            excel_file_upload.save(excel_file_path)
            print(f"Excel file saved temporarily at: {excel_file_path}")

            # Handle Audio file(s) upload
            if 'audio_files[]' not in request.files:
                raise ValueError("No audio files part in the request.")
            
            uploaded_audio_files = request.files.getlist('audio_files[]')
            if not uploaded_audio_files or uploaded_audio_files[0].filename == '':
                raise ValueError("No audio files selected.")

            saved_audio_paths = []
            for file in uploaded_audio_files:
                if file:
                    audio_filename = secure_filename(file.filename)
                    audio_file_path = os.path.join(temp_request_dir, audio_filename)
                    file.save(audio_file_path)
                    saved_audio_paths.append(audio_file_path)

            if not saved_audio_paths:
                raise ValueError("No valid audio files uploaded.")

            # Get form data
            output_file_name = request.form.get("output_file_name", "concatenated_audio.wav")
            threshold = int(request.form.get("threshold", 50))
            sheet_name = request.form.get("sheet_name")

            if not sheet_name:
                raise ValueError("Sheet name is required.")
            if not output_file_name:
                raise ValueError("Output file name is required.")

            # Audio Processing
            if len(saved_audio_paths) == 1:
                concatenated_audio_path = saved_audio_paths[0]
            else:
                concatenated_audio_path = concatenate_audio_files(
                    saved_audio_paths, output_file_name, app.config["UPLOAD_FOLDER"]
                )
                if not concatenated_audio_path:
                    raise Exception("Audio concatenation failed.")

            cleaned_audio_path = preprocess_audio(concatenated_audio_path)
            if not cleaned_audio_path:
                raise Exception("Audio preprocessing failed.")

            transcript = transcribe_audio(cleaned_audio_path, output_file_name)

            # Load Checklist and Check Compliance
            df, checklist = load_checklist(excel_file_path, sheet_name)
            results = check_compliance(transcript, checklist, threshold)

            # Calculate compliance statistics
            passed_count = sum(1 for r in results if r[0] == "PASS")
            total_checks = len(results)
            compliance_percent = round((passed_count / total_checks) * 100, 1) if total_checks else 0
            not_complied_count = total_checks - passed_count

            # Update Excel
            final_excel_output_path = update_excel(
                excel_file_path, results, sheet_name, not_complied_count, compliance_percent
            )

            # Save compliance report
            save_compliance_report(results, output_file_name)

            # Return results with Excel filename for later anomaly analysis
            updated_excel_filename = os.path.basename(final_excel_output_path)

            return jsonify({
                "results": results,
                "compliance_percent": compliance_percent,
                "not_complied_count": not_complied_count,
                "excel_updated": True,
                "updated_excel_filename": updated_excel_filename,
                "sheet_name": sheet_name
            })

        except Exception as e:
            print(f"An error occurred: {e}")
            return jsonify({"error": f"Error processing files: {e}"}), 500
        finally:
            if os.path.exists(temp_request_dir):
                shutil.rmtree(temp_request_dir)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)