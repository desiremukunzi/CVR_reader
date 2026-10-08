import os
import json
import logging
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_from_directory, url_for, session, send_file
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
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta  # For dashboard date ranges
import mysql.connector  # For dashboard database queries

# ============================================================================
# LOGGING SETUP
# Configured here, before flight_analyzer_with_db is imported, so its
# module-level logger('fdaps.database') picks up this file handler and every
# database error (even ones caught and swallowed as a bool False) gets
# written to app_errors.log with a full traceback.
# ============================================================================
logging.basicConfig(
    filename='app_errors.log',
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(name)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Try to import enhanced FlightAnalyzer with database support
# Falls back to regular FlightAnalyzer if database version not available
try:
    from flight_analyzer_with_db import FlightAnalyzer
    DATABASE_ENABLED = True
    print("Using FlightAnalyzer with database support")
except ImportError:
    from populate_db import FlightAnalyzer
    DATABASE_ENABLED = False
    print("Warning: Using FlightAnalyzer without database support")


# ============================================================================
# CONFIGURATION
# ============================================================================

# Exceedance parameters mapping (Summary sheet cell location -> parameter name)
# This matches the VBA FDR analysis system output format
EXCEEDANCE_PARAMS = {
    'B9': 'IAS',
    'B10': 'Alt',
    'B11': 'Roll',
    'B12': 'PITCH',
    'B13': 'Fcp',
    'B14': 'N1/N2 Split',
    'B15': 'N1',
    'B16': 'N2',
    'B17': 'Nmr',
    'H3': 'iAPr/p',
    'H4': 'iChips',
    'H5': 'iEMG1',
    'H6': 'iEMG2',
    'H7': 'iF_gen1',
    'H8': 'iF_gen2',
    'H9': 'iF_pump1',
    'H10': 'iF_pump2',
    'H11': 'iF_pumpS',
    'H12': 'iFire_KO-50',
    'H13': 'iFire_mgb',
    'H14': 'iFire_v1',
    'H15': 'iFire_v2',
    'H16': 'iFire1',
    'H17': 'iFire2',
    'H18': 'inFT1',
    'H19': 'inFT2',
    'H20': 'iOP_mgb',
    'H21': 'iOP1',
    'H22': 'iOP2',
    'H23': 'iQTmin',
}

# Flask setup
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production-12345')  # NEW: Required for sessions

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
print("Initializing Whisper model...")
model = WhisperModel("medium", device="cuda", compute_type="float16")

# Initialize Flight Analyzer (UPDATED)
print("Initializing FlightAnalyzer...")
if DATABASE_ENABLED:
    # Initialize with database support
    flight_analyzer = FlightAnalyzer(
        data_folder=app.config['FLIGHT_DATA_FOLDER'],
        enable_database=True
    )
else:
    # Initialize without database
    flight_analyzer = FlightAnalyzer(data_folder=app.config['FLIGHT_DATA_FOLDER'])

# Load historical data if available
historical_folder = os.path.join(app.config['FLIGHT_DATA_FOLDER'], 'historical')
if os.path.exists(historical_folder):
    try:
        print(f"Loading historical data from: {historical_folder}")
        flight_analyzer.load_historical_from_folder(historical_folder)
        historical_count = (
            flight_analyzer.historical_data['flight_id'].nunique()
            if hasattr(flight_analyzer, 'historical_data') and not flight_analyzer.historical_data.empty
            else 0
        )
        print(f"✓ Loaded {historical_count} historical flights")
    except Exception as e:
        print(f"Note: Could not load historical data: {e}")
else:
    print(f"Note: Historical data folder not found at: {historical_folder}")

# Store the latest anomaly report (for backward compatibility)
latest_anomaly_report = None


# ============================================================================
# DASHBOARD DATABASE CONNECTION AND UTILITIES
# ============================================================================

def get_db_connection():
    """
    Create database connection for dashboard queries.
    UPDATE these credentials to match your database!
    """
    return mysql.connector.connect(
        host='localhost',        # UPDATE: Your database host
        user='root',             # UPDATE: Your database username
        password='',             # UPDATE: Your database password
        database='flight_data'      # UPDATE: Your database name
    )


def get_aircraft_id_by_call_sign(call_sign):
    """
    NEW: Resolve the real aircraft_id from the aircrafts table by call_sign,
    instead of the hardcoded aircraft_id=3 default used throughout this file.
    Returns None if call_sign is missing/UNK or not found in the table, so
    callers can fall back to a literal default only as a last resort.
    """
    if not call_sign or call_sign == 'UNK':
        return None
    connection = None
    cursor = None
    try:
        connection = get_db_connection()
        cursor = connection.cursor()
        cursor.execute("SELECT id FROM aircrafts WHERE call_sign = %s", (call_sign,))
        result = cursor.fetchone()
        if result:
            return result[0]
        print(f"  ⚠️ No aircraft found in aircrafts table with call_sign='{call_sign}'")
        return None
    except Exception as e:
        print(f"  ⚠️ Error looking up aircraft_id for call_sign '{call_sign}': {e}")
        return None
    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()


def get_date_range(filter_type, custom_start=None, custom_end=None):
    """
    Calculate start and end dates based on filter type.
    Returns tuple: (start_date, end_date, prev_start_date, prev_end_date)
    """
    today = datetime.now().date()
    
    if filter_type == 'today':
        start = end = today
        prev_start = prev_end = today - timedelta(days=1)
    elif filter_type == 'yesterday':
        start = end = today - timedelta(days=1)
        prev_start = prev_end = today - timedelta(days=2)
    elif filter_type == 'last_7_days':
        end = today
        start = today - timedelta(days=6)
        prev_end = start - timedelta(days=1)
        prev_start = prev_end - timedelta(days=6)
    elif filter_type == 'last_30_days':
        end = today
        start = today - timedelta(days=29)
        prev_end = start - timedelta(days=1)
        prev_start = prev_end - timedelta(days=29)
    elif filter_type == 'this_month':
        start = today.replace(day=1)
        end = today
        prev_end = start - timedelta(days=1)
        prev_start = prev_end.replace(day=1)
    elif filter_type == 'last_month':
        first_of_this_month = today.replace(day=1)
        end = first_of_this_month - timedelta(days=1)
        start = end.replace(day=1)
        prev_end = start - timedelta(days=1)
        prev_start = prev_end.replace(day=1)
    elif filter_type == 'this_quarter':
        quarter = (today.month - 1) // 3
        start = datetime(today.year, quarter * 3 + 1, 1).date()
        end = today
        prev_quarter_start = start - relativedelta(months=3)
        prev_end = start - timedelta(days=1)
        prev_start = prev_quarter_start
    elif filter_type == 'last_quarter':
        quarter = (today.month - 1) // 3
        current_quarter_start = datetime(today.year, quarter * 3 + 1, 1).date()
        start = current_quarter_start - relativedelta(months=3)
        end = current_quarter_start - timedelta(days=1)
        prev_start = start - relativedelta(months=3)
        prev_end = start - timedelta(days=1)
    elif filter_type == 'this_year':
        start = datetime(today.year, 1, 1).date()
        end = today
        prev_start = datetime(today.year - 1, 1, 1).date()
        prev_end = prev_start + (end - start)
    elif filter_type == 'last_year':
        start = datetime(today.year - 1, 1, 1).date()
        end = datetime(today.year - 1, 12, 31).date()
        prev_start = datetime(today.year - 2, 1, 1).date()
        prev_end = datetime(today.year - 2, 12, 31).date()
    elif filter_type == 'custom' and custom_start and custom_end:
        start = datetime.strptime(custom_start, '%Y-%m-%d').date()
        end = datetime.strptime(custom_end, '%Y-%m-%d').date()
        days_diff = (end - start).days
        prev_end = start - timedelta(days=1)
        prev_start = prev_end - timedelta(days=days_diff)
    else:
        # Default to this quarter
        return get_date_range('this_quarter')
    
    return (start, end, prev_start, prev_end)


def calculate_percentage_change(current, previous):
    """Calculate percentage change between two values"""
    if previous == 0:
        return 100 if current > 0 else 0
    return round(((current - previous) / previous) * 100, 1)



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
    """
    Load checklist items from Excel sheet.
    Returns: (df, checklist_items, row_positions)
    
    row_positions maps item index to Excel row number (starting from row 2)
    """
    try:
        df = pd.read_excel(excel_file_path, sheet_name=sheet_name, engine="openpyxl")
        checklist_items = df.iloc[:, 0].dropna().tolist()
        
        # Create mapping: checklist item index -> Excel row number
        # Excel rows start at 1, header is row 1, data starts at row 2
        row_positions = {i: i + 2 for i in range(len(checklist_items))}
        
        return df, checklist_items, row_positions
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


def extract_flight_metadata_from_excel(excel_path, excel_filename):
    """
    Extract flight metadata from Excel filename and Summary sheet.
    
    Filename format: CallSign_DD-MM-YY_Sortie
    Example: UNO-561P_17-10-25_1
    
    Summary sheet cells:
    - B2: PIC
    - B3: SIC
    - B4: FE
    
    Args:
        excel_path: Full path to Excel file
        excel_filename: Just the filename (for parsing)
    
    Returns:
        dict: Flight metadata
    """
    metadata = {
        'flight_date': date.today().strftime('%Y-%m-%d'),
        'pic': 'UNK',
        'sic': 'UNK',
        'fe': 'UNK',
        'sortie': 1,
        'aircraft_id': 3,
        'call_sign': 'UNK'
    }
    
    # Extract from filename: CallSign_DD-MM-YY_Sortie
    try:
        # Remove extension
        base_name = os.path.splitext(excel_filename)[0]
        parts = base_name.split('_')
        
        if len(parts) >= 3:
            # Extract call sign
            metadata['call_sign'] = parts[0]

            # NEW: resolve the real aircraft_id from call sign instead of
            # leaving the hardcoded default of 3 above.
            resolved_aircraft_id = get_aircraft_id_by_call_sign(metadata['call_sign'])
            if resolved_aircraft_id is not None:
                metadata['aircraft_id'] = resolved_aircraft_id
                print(f"  Extracted call sign: {metadata['call_sign']} -> aircraft_id {resolved_aircraft_id}")
            else:
                print(f"  ⚠️ Warning: call sign '{metadata['call_sign']}' not found in aircrafts table, "
                      f"keeping default aircraft_id={metadata['aircraft_id']}")
            
            # Extract date: DD-MM-YY
            date_str = parts[1]
            date_parts = date_str.split('-')
            if len(date_parts) == 3:
                day, month, year = date_parts
                # Convert YY to YYYY (assume 2000s)
                full_year = f"20{year}" if len(year) == 2 else year
                # Create date in YYYY-MM-DD format
                metadata['flight_date'] = f"{full_year}-{month.zfill(2)}-{day.zfill(2)}"
                print(f"  Extracted date: {metadata['flight_date']} from {date_str}")
            
            # Extract sortie number
            try:
                metadata['sortie'] = int(parts[2])
                print(f"  Extracted sortie: {metadata['sortie']}")
            except (ValueError, IndexError):
                pass
    except Exception as e:
        print(f"  Warning: Could not parse filename '{excel_filename}': {e}")
    
    # Extract crew info from Summary sheet
    try:
        wb = load_workbook(excel_path, read_only=True, keep_vba=False, data_only=True)
        
        if 'Summary' in wb.sheetnames:
            ws = wb['Summary']
            
            # Read crew from specific cells
            pic_value = ws['B2'].value
            sic_value = ws['B3'].value
            fe_value = ws['B4'].value
            
            # Clean and validate
            if pic_value and str(pic_value).strip():
                metadata['pic'] = str(pic_value).strip().upper()
                print(f"  Extracted PIC: {metadata['pic']}")
            
            if sic_value and str(sic_value).strip():
                metadata['sic'] = str(sic_value).strip().upper()
                print(f"  Extracted SIC: {metadata['sic']}")
            
            if fe_value and str(fe_value).strip():
                metadata['fe'] = str(fe_value).strip().upper()
                print(f"  Extracted FE: {metadata['fe']}")
        else:
            print("  Warning: No 'Summary' sheet found in Excel file")
        
        wb.close()
    except Exception as e:
        print(f"  Warning: Could not read Summary sheet: {e}")
    
    return metadata



def extract_exceedances_from_excel(excel_path):
    """
    Extract exceedance counts from the Summary sheet of the analyzed Excel file.
    This reads the VBA FDR analysis results that are already in the Summary sheet.
    
    Args:
        excel_path: Path to the Excel file with Summary sheet
        
    Returns:
        List of dicts: [{'parameter': 'IAS', 'count': 5}, ...]
    """
    exceedances = []
    
    try:
        wb = load_workbook(excel_path, read_only=True, keep_vba=False, data_only=True)
        
        if 'Summary' not in wb.sheetnames:
            print("  ⚠ Warning: No Summary sheet found for exceedances extraction")
            return exceedances
        
        ws = wb['Summary']
        
        # Extract exceedance counts from specific cells
        for cell_ref, param_name in EXCEEDANCE_PARAMS.items():
            try:
                cell_value = ws[cell_ref].value
                
                # Convert to integer, handle None, empty, or '-' values
                if cell_value is None or cell_value == '' or cell_value == '-':
                    count = 0
                else:
                    count = int(float(cell_value))
                
                # Only add if count > 0
                if count > 0:
                    exceedances.append({
                        'parameter': param_name,
                        'count': count
                    })
                    
            except (ValueError, TypeError) as e:
                print(f"  ⚠ Warning: Could not read exceedance from {cell_ref}: {e}")
                continue
        
        wb.close()
        print(f"  ✓ Extracted {len(exceedances)} exceedances from Summary sheet")
        
    except Exception as e:
        print(f"  ✗ Error extracting exceedances: {e}")
    
    return exceedances


def extract_compliance_from_excel(excel_path):
    """
    Extract CVR compliance data from the Summary sheet.
    
    Reads:
    - F9: Checks Not Complied (integer)
    - F10: Compliance Percentage (formatted as "XX.X%")
    
    Args:
        excel_path: Path to the Excel file with Summary sheet
        
    Returns:
        dict: {
            'checks_not_complied': int,
            'compliance_percentage': float,
            'has_cvr_data': bool
        }
    """
    compliance_data = {
        'checks_not_complied': None,
        'compliance_percentage': None,
        'has_cvr_data': False
    }
    
    try:
        wb = load_workbook(excel_path, read_only=True, keep_vba=False, data_only=True)
        
        if 'Summary' not in wb.sheetnames:
            print("  ℹ️ No Summary sheet found for compliance extraction")
            return compliance_data
        
        ws = wb['Summary']
        
        # Extract checks not complied (F9)
        checks_not_complied_value = ws['F9'].value
        if checks_not_complied_value is not None and str(checks_not_complied_value).strip():
            try:
                compliance_data['checks_not_complied'] = int(float(checks_not_complied_value))
                compliance_data['has_cvr_data'] = True
            except (ValueError, TypeError):
                pass
        
        # Extract compliance percentage (F10)
        compliance_percent_value = ws['F10'].value
        if compliance_percent_value is not None and str(compliance_percent_value).strip():
            try:
                # Remove % symbol if present and convert to float
                percent_str = str(compliance_percent_value).replace('%', '').strip()
                compliance_data['compliance_percentage'] = float(percent_str)
                compliance_data['has_cvr_data'] = True
            except (ValueError, TypeError):
                pass
        
        wb.close()
        
        if compliance_data['has_cvr_data']:
            print(f"  ✓ Extracted compliance from Excel:")
            print(f"     Checks Not Complied: {compliance_data['checks_not_complied']}")
            print(f"     Compliance: {compliance_data['compliance_percentage']}%")
        else:
            print("  ℹ️ No CVR compliance data found in Excel Summary sheet")
            
    except Exception as e:
        print(f"  ⚠️ Warning: Could not extract compliance data from Excel: {e}")
    
    return compliance_data


def extract_missed_checks_from_excel(excel_path):
    """
    Extract individual missed check items from the checklist sheet.
    Reads column B to find rows with ✘ marks (failed checks).
    
    Args:
        excel_path: Path to the Excel file with checklist sheet
        
    Returns:
        dict: {
            'missed_checks': [(item, score, excel_row), ...],
            'checklist_type_id': int,
            'sheet_name': str
        }
    """
    result = {
        'missed_checks': [],
        'checklist_type_id': 1,  # Default to AC-GPU
        'sheet_name': None
    }
    
    # Checklist sheet names and their IDs
    checklist_sheets = {
        'STARTING WITH AC-GPU CHECKLIST': 1,
        'STARTING WITH DC-GPU CHECKLIST': 2,
        'STARTING WITHOUT GPU CHECKLIST': 3
    }
    
    try:
        wb = load_workbook(excel_path, read_only=True, keep_vba=False, data_only=True)
        
        # Find which checklist sheet exists and has data
        for sheet_name, checklist_id in checklist_sheets.items():
            if sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                result['sheet_name'] = sheet_name
                result['checklist_type_id'] = checklist_id
                
                # Read column B starting from row 2 (row 1 is header)
                # The column should have ✔ or ✘ marks
                missed_checks = []
                row_num = 2
                
                while True:
                    cell_value = ws[f'B{row_num}'].value
                    
                    # Stop if we hit an empty cell
                    if cell_value is None:
                        break
                    
                    # Check if this row has a ✘ (failed check)
                    if cell_value == '✘':
                        # Read the checklist item from column A
                        checklist_item = ws[f'A{row_num}'].value
                        if checklist_item:
                            # Format: (item, score, excel_row)
                            # We don't have the score from Excel, so use 0.0
                            missed_checks.append((str(checklist_item).strip(), 0.0, row_num))
                    
                    row_num += 1
                    
                    # Safety limit to avoid infinite loop
                    if row_num > 200:
                        break
                
                result['missed_checks'] = missed_checks
                
                if missed_checks:
                    print(f"  ✓ Extracted {len(missed_checks)} missed checks from '{sheet_name}'")
                else:
                    print(f"  ℹ️ No missed checks found in '{sheet_name}'")
                
                break  # Found the sheet, stop looking
        
        wb.close()
        
    except Exception as e:
        print(f"  ⚠️ Warning: Could not extract missed checks from Excel: {e}")
    
    return result


def extract_compliance_from_excel(excel_path):
    """
    Extract CVR compliance data from the Summary sheet.
    
    Reads:
    - F9: Checks Not Complied (integer)
    - F10: Compliance Percentage (formatted as "XX.X%")
    
    Args:
        excel_path: Path to the Excel file with Summary sheet
        
    Returns:
        dict: {
            'checks_not_complied': int,
            'compliance_percentage': float,
            'has_cvr_data': bool
        }
    """
    compliance_data = {
        'checks_not_complied': None,
        'compliance_percentage': None,
        'has_cvr_data': False
    }
    
    try:
        wb = load_workbook(excel_path, read_only=True, keep_vba=False, data_only=True)
        
        if 'Summary' not in wb.sheetnames:
            print("  ℹ️ No Summary sheet found for compliance extraction")
            return compliance_data
        
        ws = wb['Summary']
        
        # Extract checks not complied (F9)
        checks_not_complied_value = ws['F9'].value
        if checks_not_complied_value is not None and str(checks_not_complied_value).strip():
            try:
                compliance_data['checks_not_complied'] = int(float(checks_not_complied_value))
                compliance_data['has_cvr_data'] = True
            except (ValueError, TypeError):
                pass
        
        # Extract compliance percentage (F10)
        compliance_percent_value = ws['F10'].value
        if compliance_percent_value is not None and str(compliance_percent_value).strip():
            try:
                # Remove % symbol if present and convert to float
                percent_str = str(compliance_percent_value).replace('%', '').strip()
                compliance_data['compliance_percentage'] = float(percent_str)
                compliance_data['has_cvr_data'] = True
            except (ValueError, TypeError):
                pass
        
        wb.close()
        
        if compliance_data['has_cvr_data']:
            print(f"  ✓ Extracted compliance from Excel:")
            print(f"     Checks Not Complied: {compliance_data['checks_not_complied']}")
            print(f"     Compliance: {compliance_data['compliance_percentage']}%")
        else:
            print("  ℹ️ No CVR compliance data found in Excel Summary sheet")
            
    except Exception as e:
        print(f"  ⚠️ Warning: Could not extract compliance data from Excel: {e}")
    
    return compliance_data


def extract_exceedances_from_excel(excel_path):
    """
    Extract exceedance counts from the Summary sheet of the analyzed Excel file.
    This reads the VBA FDR analysis results that are already in the Summary sheet.
    
    Args:
        excel_path: Path to the Excel file with Summary sheet
        
    Returns:
        List of dicts: [{'parameter': 'IAS', 'count': 5}, ...]
    """
    exceedances = []
    
    try:
        wb = load_workbook(excel_path, read_only=True, keep_vba=False, data_only=True)
        
        if 'Summary' not in wb.sheetnames:
            print("  ⚠ Warning: No Summary sheet found for exceedances extraction")
            return exceedances
        
        ws = wb['Summary']
        
        # Extract exceedance counts from specific cells
        for cell_ref, param_name in EXCEEDANCE_PARAMS.items():
            try:
                cell_value = ws[cell_ref].value
                
                # Convert to integer, handle None, empty, or '-' values
                if cell_value is None or cell_value == '' or cell_value == '-':
                    count = 0
                else:
                    count = int(float(cell_value))
                
                # Only add if count > 0
                if count > 0:
                    exceedances.append({
                        'parameter': param_name,
                        'count': count
                    })
                    
            except (ValueError, TypeError) as e:
                print(f"  ⚠ Warning: Could not read exceedance from {cell_ref}: {e}")
                continue
        
        wb.close()
        print(f"  ✓ Extracted {len(exceedances)} exceedances from Summary sheet")
        
    except Exception as e:
        print(f"  ✗ Error extracting exceedances: {e}")
    
    return exceedances


# ============================================================================
# SAVE SUMMARY SHEET TO DATABASE  (BUR + CARE FDR)
#
# Both recorder formats lay the Summary sheet out identically:
#   continuous : names in column A, counts in column B, from row 9, ends at TOTAL
#   binary     : names in column G, counts in column H, from row 3
# Only the NAMES and the number of binary rows differ (BUR uses the
# MI_17V_5_name values, CARE uses the MI_17_1V_name values).
#
# So instead of hardcoding cell -> parameter maps (which breaks whenever the
# row count changes, as it did when iHSaux/iHSmain left the BUR layout), the
# name is read from the sheet and resolved against the `parameters` table.
# Adding a new recorder format then becomes a data change, not a code change.
# ============================================================================

# Rows are scanned until this many consecutive blanks, so tall/wrapped rows
# in the middle of a block don't truncate the scan.
_BLANK_TOLERANCE = 3
_MAX_SCAN_ROW = 60

# ----------------------------------------------------------------------------
# The after-mission review block in the Summary sheet.
#
# This is its OWN table below the exceedance tables, not extra columns on
# them. Layout (from the template):
#     row 20  header: EXCEEDANCES | COUNT | OBSERVED BY | AFTER MISSION REVIEW
#     row 21+ one row per parameter that actually recorded an exceedance
#
# Its column A reuses the same parameter names as the blocks above, so rows
# are matched back by name, not by position. Its COUNT column is read only
# as a cross-check; the authoritative counts stay the ones in the continuous
# and binary blocks.
REVIEW_BLOCK_START_ROW = 21
REVIEW_NAME_COL = 'A'
REVIEW_COUNT_COL = 'B'
REVIEW_OBSERVED_COL = 'C'
REVIEW_NOTE_COL = 'D'

# The template pre-draws empty bordered rows under the filled ones, so the
# scan runs to a fixed bound rather than stopping at the first blank.
REVIEW_BLOCK_MAX_ROW = 80

# Words that mean "a new table starts here", so the scan never runs on into
# whatever sits below the review block.
_REVIEW_STOP_WORDS = ('TOTAL', 'EXCEEDANCES', 'PARAMETER', 'BINARY')

# Guard rails matching the exceedances columns. Values longer than these are
# rejected with a clear message rather than silently truncated: a clipped
# after-mission note is worse than a refused save.
# exceedances.observed_by is VARCHAR(100).
MAX_OBSERVED_BY = 100
# exceedances.after_mission_review is TEXT (65,535 BYTES). utf8mb4 uses up to
# 4 bytes per character, so capping at 16,000 CHARACTERS can never overflow
# the column even in the worst case, while being long enough that it will
# not block a real after-mission write-up.
MAX_REVIEW_NOTE = 16000

_param_map_cache = None
_param_detail_cache = None


def _load_parameter_tables(force_reload=False):
    """
    Load both parameter lookups in one query:
      name map    {lowercased name: canonical MI_17V_5_name}, covering BOTH
                  name columns so a sheet in either convention resolves.
      detail map  {canonical name: {'description': ..., 'discrete': ...}},
                  used to label rows in the review table.
    """
    global _param_map_cache, _param_detail_cache
    if _param_map_cache is not None and not force_reload:
        return _param_map_cache, _param_detail_cache

    name_map = {}
    detail_map = {}
    connection = None
    cursor = None
    try:
        connection = get_db_connection()
        cursor = connection.cursor()
        cursor.execute(
            "SELECT MI_17V_5_name, MI_17_1V_name, description, discrete FROM parameters"
        )
        for mi17v5, mi171v, description, discrete in cursor.fetchall():
            if mi17v5 and str(mi17v5).strip():
                name_map[str(mi17v5).strip().lower()] = mi17v5
                detail_map[mi17v5] = {
                    'description': description or '',
                    'discrete': bool(discrete),
                }
            # Guard against the empty-string entry: '' would otherwise match
            # every blank cell and silently map it to a real parameter.
            if mi171v and str(mi171v).strip():
                name_map[str(mi171v).strip().lower()] = mi17v5
        _param_map_cache = name_map
        _param_detail_cache = detail_map
        print(f"  Loaded {len(name_map)} parameter name variants")
    except Exception as e:
        print(f"  ⚠️ Could not load parameter tables: {e}")
        name_map, detail_map = {}, {}
    finally:
        if cursor:
            cursor.close()
        if connection:
            connection.close()
    return name_map, detail_map


def get_parameter_name_map(force_reload=False):
    """Name -> canonical MI_17V_5_name. See _load_parameter_tables."""
    return _load_parameter_tables(force_reload)[0]


def get_parameter_details(force_reload=False):
    """Canonical name -> {'description', 'discrete'}. See _load_parameter_tables."""
    return _load_parameter_tables(force_reload)[1]


def _to_count(value):
    """Cell value -> non-negative int, treating blanks and '-' as 0."""
    if value is None or value == '' or value == '-':
        return 0
    try:
        return int(float(value))
    except (ValueError, TypeError):
        return 0


def _scan_parameter_block(ws, name_col, count_col, start_row, name_map, block=''):
    """
    Walk one block of the Summary sheet, resolving each parameter name.

    Returns (resolved, unmapped, total) where resolved is a list of
    {'parameter', 'count', 'sheet_name', 'block', 'row'} for counts > 0,
    unmapped is a list of sheet names with no row in `parameters`, and total
    is the sum of every count in the block (including zeros and unmapped).

    sheet_name/block/row are carried so the review screen can show each row
    the way it appears in Excel, which is how the reviewer recognises it.

    The after-mission fields are NOT read here: in these blocks columns C
    and D hold "% of Total" and "Description". They come from the separate
    review block below, merged in by parameter name.
    """
    resolved = []
    unmapped = []
    total = 0
    blanks = 0
    row = start_row

    while row <= _MAX_SCAN_ROW and blanks < _BLANK_TOLERANCE:
        raw_name = ws[f'{name_col}{row}'].value

        if raw_name is None or str(raw_name).strip() == '':
            blanks += 1
            row += 1
            continue
        blanks = 0

        name = str(raw_name).strip()
        if name.upper().startswith('TOTAL'):
            break

        count = _to_count(ws[f'{count_col}{row}'].value)
        total += count

        canonical = name_map.get(name.lower())
        if canonical is None:
            unmapped.append(name)
        elif count > 0:
            resolved.append({
                'parameter': canonical,
                'count': count,
                'sheet_name': name,
                'block': block,
                'row': row,
                'observed_by': '',
                'review_note': '',
            })

        row += 1

    return resolved, unmapped, total


def _scan_review_block(ws, name_map):
    """
    Read the after-mission review block (header row 20, data from row 21).

    Returns {canonical parameter name: {'observed_by', 'review_note',
    'sheet_count', 'row'}} for every row that names a parameter.

    Matching is by NAME, not by position: the block lists only the parameters
    that actually recorded an exceedance, in whatever order the template wrote
    them. Pairing by row number would silently attach one parameter's review
    note to another parameter's exceedance.

    Only CONTINUOUS parameters are written here. Binary parameters never
    appear, so they find no match and reach the review screen with both
    fields blank for the reviewer to fill in. That needs no special case:
    name-matching simply does not find them.

    The block's own COUNT column is returned as sheet_count for cross-checking
    only. The authoritative counts remain the ones in the blocks above.
    """
    found = {}
    unmatched = []

    for row in range(REVIEW_BLOCK_START_ROW, REVIEW_BLOCK_MAX_ROW + 1):
        raw_name = ws[f'{REVIEW_NAME_COL}{row}'].value
        if raw_name is None or str(raw_name).strip() == '':
            # The template pre-draws empty rows, so a blank is not the end.
            continue

        name = str(raw_name).strip()
        if name.upper().startswith(_REVIEW_STOP_WORDS):
            # A new table header: stop rather than read whatever follows.
            if row > REVIEW_BLOCK_START_ROW:
                break
            continue

        canonical = name_map.get(name.lower())
        if canonical is None:
            unmatched.append(name)
            continue

        def cell_text(col):
            raw = ws[f'{col}{row}'].value
            return str(raw).strip() if raw is not None else ''

        found[canonical] = {
            'observed_by': cell_text(REVIEW_OBSERVED_COL),
            'review_note': cell_text(REVIEW_NOTE_COL),
            'sheet_count': _to_count(ws[f'{REVIEW_COUNT_COL}{row}'].value),
            'row': row,
        }

    if unmatched:
        print(f"  \u26a0\ufe0f Review block rows with unrecognised parameter names: "
              f"{', '.join(unmatched)}")

    return found


def extract_summary_for_db(excel_path):
    """
    Read the Summary sheet of a BUR or CARE file.

    Returns a dict with crew, exceedance rows, per-block totals, and any
    parameter names that could not be resolved. Unmapped names are reported
    rather than raised: exceedances.parameter_MI_17V_5_name is a foreign key,
    so saving an unmapped name would fail the whole insert. Skipping it saves
    what is valid and tells the user exactly what is missing.
    """
    result = {
        'pic': None, 'sic': None, 'fe': None,
        'exceedances': [],
        'unmapped': [],
        'continuous_total': 0,
        'discrete_total': 0,
    }

    name_map = get_parameter_name_map()
    if not name_map:
        raise RuntimeError("Parameter name map is empty; cannot resolve any parameter")

    def scan_with(current_map):
        c_rows, c_unmapped, c_total = _scan_parameter_block(
            ws, 'A', 'B', 9, current_map, block='continuous')
        b_rows, b_unmapped, b_total = _scan_parameter_block(
            ws, 'G', 'H', 3, current_map, block='discrete')
        return c_rows, c_unmapped, c_total, b_rows, b_unmapped, b_total

    wb = load_workbook(excel_path, read_only=True, keep_vba=False, data_only=True)
    try:
        if 'Summary' not in wb.sheetnames:
            raise ValueError(f"No 'Summary' sheet in {os.path.basename(excel_path)}")

        ws = wb['Summary']

        def text(ref):
            v = ws[ref].value
            return str(v).strip() if v is not None and str(v).strip() else None

        result['pic'] = text('B2')
        result['sic'] = text('B3')
        result['fe'] = text('B4')

        (cont_rows, cont_unmapped, cont_total,
         bin_rows, bin_unmapped, bin_total) = scan_with(name_map)

        # The name map is cached per process. If anything failed to resolve,
        # the cache may simply predate a row added to `parameters`, so reload
        # once and rescan before reporting it as unmapped. This means editing
        # the parameters table takes effect without restarting Flask.
        if cont_unmapped or bin_unmapped:
            print("  Unresolved names found, reloading parameter map...")
            refreshed = get_parameter_name_map(force_reload=True)
            if refreshed:
                (cont_rows, cont_unmapped, cont_total,
                 bin_rows, bin_unmapped, bin_total) = scan_with(refreshed)

        # B18 holds a pre-computed continuous total. Prefer it when present,
        # fall back to the summed block so a layout shift can't zero this out.
        b18 = _to_count(ws['B18'].value)
        result['continuous_total'] = b18 if b18 else cont_total
        result['discrete_total'] = bin_total

        result['exceedances'] = cont_rows + bin_rows
        result['unmapped'] = cont_unmapped + bin_unmapped

        # Merge in anything already written into the after-mission review
        # block, matched by parameter name.
        review_rows = _scan_review_block(ws, name_map)

        # Every row starts flagged abnormal (a genuine exceedance). The
        # reviewer only ever downgrades, never upgrades, so the default is
        # the conservative one: an unreviewed exceedance counts as real.
        details = get_parameter_details()
        for row_data in result['exceedances']:
            meta = details.get(row_data['parameter'], {})
            row_data['abnormal'] = True
            row_data['description'] = meta.get('description', '')

            review = review_rows.get(row_data['parameter'])
            if review:
                row_data['observed_by'] = review['observed_by']
                row_data['review_note'] = review['review_note']
                # The review block repeats the count. If it disagrees with the
                # block above, say so rather than picking one silently: it
                # means the sheet was edited in one place and not the other.
                if review['sheet_count'] and review['sheet_count'] != row_data['count']:
                    result.setdefault('count_mismatches', []).append(
                        f"{row_data['parameter']}: {row_data['count']} in the "
                        f"exceedance table vs {review['sheet_count']} in the "
                        f"review block (row {review['row']})"
                    )

        # Reviewed rows naming a parameter with no exceedance above are
        # surfaced, not dropped: usually a stale row left in the template.
        orphans = set(review_rows) - {e['parameter'] for e in result['exceedances']}
        if orphans:
            result['review_orphans'] = sorted(orphans)
    finally:
        wb.close()

    print(f"  Summary: {len(result['exceedances'])} exceedance rows, "
          f"continuous={result['continuous_total']}, discrete={result['discrete_total']}")
    if result['unmapped']:
        print(f"  ⚠️ Unmapped parameter names: {', '.join(result['unmapped'])}")

    return result


def parse_flight_filename(filename):
    """
    CALLSIGN_DD-MM-YY_SORTIE -> (call_sign, date, sortie).
    Same convention for BUR and CARE. Returns (None, None, None) on failure.
    """
    try:
        stem = os.path.splitext(os.path.basename(filename))[0]
        parts = stem.split('_')
        if len(parts) < 3:
            return None, None, None

        call_sign = parts[0].strip()
        d, m, y = (int(x) for x in parts[1].split('-'))
        if y < 100:
            y += 2000
        return call_sign, date(y, m, d), int(parts[2])
    except (ValueError, IndexError):
        return None, None, None


class SummaryRequestError(Exception):
    """Raised when a summary request cannot be resolved. Carries an HTTP status."""

    def __init__(self, message, status=400):
        super().__init__(message)
        self.status = status


def _resolve_summary_request():
    """
    Shared front half of the preview and save routes: locate the Excel file,
    read the flight identity out of its filename, resolve the aircraft, and
    extract the Summary sheet.

    Both routes must read the workbook the SAME way. If preview and save
    parsed independently, a reviewer could approve one set of numbers and
    save another, which for a safety record is worse than no review at all.

    Returns (context_dict, temp_path). temp_path is non-None only when this
    request wrote an upload to disk and is responsible for cleaning it up.
    """
    temp_path = None

    if 'excel_file' in request.files and request.files['excel_file'].filename:
        upload = request.files['excel_file']
        source_name = upload.filename
        safe_name = secure_filename(source_name)
        temp_path = os.path.join(app.config['UPLOAD_FOLDER'], safe_name)
        upload.save(temp_path)
        excel_path = temp_path
    else:
        payload = request.get_json(silent=True) or {}
        source_name = payload.get('excel_filename') or request.form.get('excel_filename')
        if not source_name:
            raise SummaryRequestError('No Excel file provided')
        candidate = os.path.join(app.config['COMPLIANCE_EXCEL_OUTPUT'], source_name)
        if not os.path.exists(candidate):
            candidate = os.path.join(app.config['UPLOAD_FOLDER'], source_name)
        if not os.path.exists(candidate):
            raise SummaryRequestError(f'File not found: {source_name}', 404)
        excel_path = candidate

    # secure_filename() is applied above for the on-disk name only; the
    # ORIGINAL name is parsed here, because secure_filename can rewrite
    # the call sign (it strips characters that appear in some call signs).
    call_sign, flight_date, sortie = parse_flight_filename(source_name)
    if not flight_date:
        raise SummaryRequestError(
            f"Could not read date/sortie from '{source_name}'. "
            f"Expected CALLSIGN_DD-MM-YY_SORTIE."
        )

    aircraft_id = get_aircraft_id_by_call_sign(call_sign)
    if aircraft_id is None:
        raise SummaryRequestError(
            f"Call sign '{call_sign}' is not in the aircrafts table. "
            f"Add it before saving, so the flight is not filed "
            f"against the wrong aircraft."
        )

    summary = extract_summary_for_db(excel_path)

    missing_crew = [k.upper() for k in ('pic', 'sic', 'fe') if not summary[k]]
    if missing_crew:
        raise SummaryRequestError(
            f"Summary sheet is missing crew code(s): "
            f"{', '.join(missing_crew)} (cells B2/B3/B4)"
        )

    context = {
        'source_name': source_name,
        'excel_path': excel_path,
        'call_sign': call_sign,
        'flight_date': flight_date,
        'sortie': sortie,
        'aircraft_id': aircraft_id,
        'summary': summary,
    }
    return context, temp_path


def _cleanup_temp(temp_path):
    """Remove only a file this request created; never a compliance output."""
    if temp_path and os.path.exists(temp_path):
        try:
            os.remove(temp_path)
        except OSError:
            pass


@app.route("/preview_summary_exceedances", methods=["POST"])
def preview_summary_exceedances():
    """
    Read the Summary sheet and return what WOULD be saved, writing nothing.

    This is the review step: the reviewer sees every exceedance the sheet
    reports, each pre-flagged abnormal (genuine), and can mark individual
    rows as sensor false positives before committing. Nothing touches the
    database until /save_summary_to_db is called.
    """
    temp_path = None
    try:
        context, temp_path = _resolve_summary_request()
        summary = context['summary']

        print("\n" + "=" * 60)
        print(f"PREVIEW SUMMARY: {context['source_name']}")
        print(f"  {len(summary['exceedances'])} exceedance rows for review")
        print("=" * 60)

        # If this flight was reviewed before, show THAT, not a blank slate.
        # Excel only ever seeds continuous parameters, and binary ones are
        # reviewed solely in this app, so without reading back what was saved
        # a second review would reset earlier work to empty.
        already_saved = 0
        if DATABASE_ENABLED and flight_analyzer.db_manager:
            existing_flight_id = flight_analyzer.db_manager.find_flight_id(
                flight_date=context['flight_date'],
                pic=summary['pic'], sic=summary['sic'], fe=summary['fe'],
                sortie=context['sortie'],
            )
            if existing_flight_id:
                stored = flight_analyzer.db_manager.get_exceedance_reviews(existing_flight_id)
                for row_data in summary['exceedances']:
                    prior = stored.get(row_data['parameter'])
                    if not prior:
                        continue
                    row_data['abnormal'] = prior['abnormal']
                    # A stored value wins over the Excel seed: it is the later
                    # and more deliberate of the two.
                    if prior['observed_by']:
                        row_data['observed_by'] = prior['observed_by']
                    if prior['review_note']:
                        row_data['review_note'] = prior['review_note']
                    if (prior['observed_by'] or prior['review_note']
                            or not prior['abnormal']):
                        already_saved += 1
                if already_saved:
                    print(f"  Loaded previously saved review for {already_saved} row(s) "
                          f"from flight {existing_flight_id}")

        response = {
            'success': True,
            'excel_filename': context['source_name'],
            'call_sign': context['call_sign'],
            'aircraft_id': context['aircraft_id'],
            'flight_date': context['flight_date'].strftime('%Y-%m-%d'),
            'sortie': context['sortie'],
            'crew': {
                'pic': summary['pic'],
                'sic': summary['sic'],
                'fe': summary['fe'],
            },
            'exceedances': summary['exceedances'],
            'continuous_total': summary['continuous_total'],
            'discrete_total': summary['discrete_total'],
            'previously_reviewed': already_saved,
        }
        notes = []
        if summary['unmapped']:
            notes.append(
                f"{len(summary['unmapped'])} parameter(s) will be skipped, not in "
                f"the parameters table: {', '.join(summary['unmapped'])}"
            )
            response['unmapped'] = summary['unmapped']
        if summary.get('count_mismatches'):
            notes.append(
                "Counts disagree between the exceedance table and the review "
                "block: " + "; ".join(summary['count_mismatches'])
            )
        if summary.get('review_orphans'):
            notes.append(
                "Review block rows with no matching exceedance (ignored): "
                + ", ".join(summary['review_orphans'])
            )
        if notes:
            response['warning'] = " | ".join(notes)

        return jsonify(response)

    except SummaryRequestError as e:
        return jsonify({'success': False, 'error': str(e)}), e.status
    except Exception as e:
        import traceback
        print(f"❌ Error in preview_summary_exceedances: {e}")
        traceback.print_exc()
        logger.error(f"preview_summary_exceedances failed: {e}")
        logger.error(traceback.format_exc())
        return jsonify({'success': False, 'error': str(e)}), 500
    finally:
        _cleanup_temp(temp_path)


@app.route("/save_summary_to_db", methods=["POST"])
def save_summary_to_db():
    """
    Save the Summary sheet (flight record + exceedances) to the database,
    independent of the compliance workflow.

    Accepts either a freshly uploaded file (multipart 'excel_file') or the
    name of a file already produced by the compliance run ('excel_filename').
    Compliance figures are optional: when no compliance report has been run
    they are left NULL rather than written as zero, which would read as
    "0% compliant" instead of "not assessed".

    Review: the caller may send 'false_positives', a list of canonical
    parameter names the reviewer judged to be sensor faults. Those rows are
    still saved, with abnormal = 0, so the record shows the sensor fired and
    that a human dismissed it. COUNTS come from the workbook, never from the
    client: the reviewer labels rows, they do not get to restate the numbers.
    """
    temp_path = None
    try:
        context, temp_path = _resolve_summary_request()
        summary = context['summary']
        source_name = context['source_name']
        call_sign = context['call_sign']
        flight_date = context['flight_date']
        sortie = context['sortie']
        aircraft_id = context['aircraft_id']

        print("\n" + "=" * 60)
        print(f"SAVE SUMMARY TO DB: {source_name}")
        print("=" * 60)

        # ---- apply the reviewer's decisions ---------------------------------
        # 'reviews' carries one entry per row: the abnormal flag plus the
        # after-mission fields as edited on screen. 'false_positives' (a bare
        # list of names) is still accepted so an older client keeps working.
        #
        # Only labels and notes travel from the client. COUNTS are re-read
        # from the workbook above and never taken from the request.
        def _field(name):
            body = request.get_json(silent=True) or {}
            if name in body:
                return body.get(name)
            raw = request.form.get(name)
            if raw:
                try:
                    return json.loads(raw)
                except (ValueError, TypeError):
                    return None
            return None

        reviews_raw = _field('reviews')
        fp_raw = _field('false_positives')

        reviews_by_param = {}
        if isinstance(reviews_raw, list):
            for item in reviews_raw:
                if isinstance(item, dict) and item.get('parameter'):
                    reviews_by_param[str(item['parameter']).strip()] = item

        false_positives = set()
        if reviews_by_param:
            false_positives = {
                param for param, item in reviews_by_param.items()
                if item.get('abnormal') is False
            }
        elif fp_raw:
            false_positives = {str(p).strip() for p in fp_raw}

        sheet_params = {e['parameter'] for e in summary['exceedances']}
        unknown = (set(reviews_by_param) | false_positives) - sheet_params
        if unknown:
            # Names that aren't in this workbook mean the review screen and the
            # file have drifted apart (a different file, or an edited sheet).
            # Saving anyway would silently ignore the reviewer's decision.
            return jsonify({
                'success': False,
                'error': (f"These parameters were reviewed but are not in this file's "
                          f"Summary sheet: {', '.join(sorted(unknown))}. "
                          f"Re-run the review against the current file.")
            }), 400

        def _clip_check(value, limit, label, param):
            text = str(value).strip() if value is not None else ''
            if len(text) > limit:
                raise SummaryRequestError(
                    f"{label} for '{param}' is {len(text)} characters; the column "
                    f"holds {limit}. Shorten it rather than letting it be cut off."
                )
            return text

        for row_data in summary['exceedances']:
            item = reviews_by_param.get(row_data['parameter'], {})
            row_data['abnormal'] = row_data['parameter'] not in false_positives
            # Edited value wins; the sheet value seeded it and stands otherwise.
            if 'observed_by' in item:
                row_data['observed_by'] = _clip_check(
                    item.get('observed_by'), MAX_OBSERVED_BY,
                    'Observed by', row_data['parameter'])
            if 'review_note' in item:
                row_data['review_note'] = _clip_check(
                    item.get('review_note'), MAX_REVIEW_NOTE,
                    'After-mission review', row_data['parameter'])

        # Flight-level totals count GENUINE exceedances only. A dismissed
        # sensor fault is not an exceedance the crew flew, so including it
        # would inflate every report built on these columns. The raw sheet
        # totals are still returned below so the difference stays visible.
        dismissed_continuous = sum(
            e['count'] for e in summary['exceedances']
            if not e['abnormal'] and e.get('block') == 'continuous'
        )
        dismissed_discrete = sum(
            e['count'] for e in summary['exceedances']
            if not e['abnormal'] and e.get('block') == 'discrete'
        )
        genuine_continuous = max(0, summary['continuous_total'] - dismissed_continuous)
        genuine_discrete = max(0, summary['discrete_total'] - dismissed_discrete)

        if false_positives:
            print(f"  Reviewer dismissed {len(false_positives)} parameter(s) as "
                  f"sensor false positives: {', '.join(sorted(false_positives))}")
            print(f"  Totals: continuous {summary['continuous_total']} -> {genuine_continuous}, "
                  f"discrete {summary['discrete_total']} -> {genuine_discrete}")

        # ---- optional compliance -------------------------------------------
        # Checked in three places, in order: multipart form fields (the button
        # posts FormData, so request.get_json() is None for those requests),
        # then a JSON body, then the session left by a compliance run.
        # Left as None when absent, so the column stays NULL: "not assessed"
        # and "0% compliant" must not look the same in the database.
        def _num(raw, caster):
            if raw is None or raw == '':
                return None
            try:
                return caster(float(str(raw).replace('%', '').strip()))
            except (ValueError, TypeError):
                return None

        payload = request.get_json(silent=True) or {}

        compliance_percentage = _num(
            request.form.get('compliance_percent', payload.get('compliance_percent')), float
        )
        checks_not_complied = _num(
            request.form.get('not_complied_count', payload.get('not_complied_count')), int
        )

        if compliance_percentage is None:
            cvr = session.get('cvr_results') or {}
            compliance_percentage = _num(cvr.get('compliance_percent'), float)
            checks_not_complied = _num(cvr.get('not_complied_count'), int)

        # ---- write ----------------------------------------------------------
        if not DATABASE_ENABLED or not flight_analyzer.db_manager:
            return jsonify({'success': False,
                            'error': 'Database integration is not enabled'}), 500

        db = flight_analyzer.db_manager

        # Shared with the anomaly-report save path on purpose: one writer for
        # the flights table means a column added there cannot go stale here.
        flight_id = db.get_or_create_flight(
            flight_date=flight_date,
            pic=summary['pic'],
            sic=summary['sic'],
            fe=summary['fe'],
            sortie=sortie,
            aircraft_id=aircraft_id,
            compliance_percentage=compliance_percentage,
            checks_not_complied=checks_not_complied,
            continuous_exceedances=genuine_continuous,
            discrete_exceedances=genuine_discrete,
        )

        if not flight_id:
            detail = getattr(db, 'last_error', None) or 'Unknown error'
            return jsonify({'success': False,
                            'error': f'Could not create/update flight: {detail}'}), 500

        db.delete_flight_exceedances(flight_id)
        saved = db.save_exceedances(flight_id, summary['exceedances'])
        if not saved:
            detail = getattr(db, 'last_error', None) or 'Unknown error'
            return jsonify({'success': False,
                            'error': f'Flight {flight_id} saved, but exceedances failed: {detail}'}), 500

        response = {
            'success': True,
            'flight_id': flight_id,
            'call_sign': call_sign,
            'aircraft_id': aircraft_id,
            'flight_date': flight_date.strftime('%Y-%m-%d'),
            'sortie': sortie,
            'crew': {'pic': summary['pic'], 'sic': summary['sic'], 'fe': summary['fe']},
            'exceedances_saved': len(summary['exceedances']),
            'abnormal_saved': sum(1 for e in summary['exceedances'] if e['abnormal']),
            'false_positives_saved': sorted(false_positives),
            'reviewed_rows': sum(
                1 for e in summary['exceedances']
                if e.get('observed_by') or e.get('review_note')
            ),
            # What went into the flights table (genuine only) alongside what
            # the sheet reported, so the reviewer can see their own effect.
            'continuous_exceedances': genuine_continuous,
            'discrete_exceedances': genuine_discrete,
            'sheet_continuous_total': summary['continuous_total'],
            'sheet_discrete_total': summary['discrete_total'],
            'compliance_saved': compliance_percentage is not None,
        }
        if summary['unmapped']:
            response['warning'] = (
                f"{len(summary['unmapped'])} parameter(s) skipped, not in the "
                f"parameters table: {', '.join(summary['unmapped'])}"
            )
            response['unmapped'] = summary['unmapped']

        print(f"✓ Saved flight {flight_id}: {len(summary['exceedances'])} exceedances "
              f"({response['abnormal_saved']} abnormal, "
              f"{len(false_positives)} dismissed as sensor faults)")
        return jsonify(response)

    except SummaryRequestError as e:
        return jsonify({'success': False, 'error': str(e)}), e.status
    except Exception as e:
        import traceback
        print(f"❌ Error in save_summary_to_db: {e}")
        traceback.print_exc()
        logger.error(f"save_summary_to_db failed: {e}")
        logger.error(traceback.format_exc())
        return jsonify({'success': False, 'error': str(e)}), 500
    finally:
        # Only remove a file this request created; never a compliance output.
        if temp_path and os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except OSError:
                pass


# ============================================================================
# NEW ENDPOINTS FOR DATABASE INTEGRATION
# ============================================================================

@app.route("/add_to_training", methods=["POST"])
def add_to_training():
    """
    NEW ENDPOINT: Add flight to training data and retrain models.
    Called when user clicks "Add to Training Data" button on anomaly report page.
    """
    try:
        data = request.get_json() or {}
        flight_id = data.get('flight_id')
        
        print(f"🔄 Retraining models with accumulated flight data (flight_id: {flight_id})...")
        
        # Retrain models with all accumulated data
        flight_analyzer.train_models()
        
        # Update session metadata to mark as added to training
        if 'analysis_metadata' in session:
            session['analysis_metadata']['added_to_training'] = True
            session.modified = True
        
        # Also update global variable for current view
        global latest_anomaly_report
        if latest_anomaly_report:
            latest_anomaly_report['added_to_training'] = True
        
        print("✓ Model retraining completed successfully")
        return jsonify({'success': True})
        
    except Exception as e:
        print(f"❌ Error in add_to_training: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route("/save_to_database", methods=["POST"])
def save_to_database():
    """
    NEW ENDPOINT: Save flight record, anomalies, and missed checks to MySQL database.
    Called when user clicks "Save to Database" button on anomaly report page.
    """
    try:
        if not DATABASE_ENABLED:
            return jsonify({
                'success': False, 
                'error': 'Database support not available. Install flight_analyzer_with_db.py'
            }), 400
        
        data = request.get_json()
        frontend_metadata = data.get('flight_metadata')
        anomalies_summary = data.get('anomalies_summary')
        
        if not frontend_metadata:
            return jsonify({'success': False, 'error': 'Flight metadata is required'}), 400
        
        print(f"\n{'='*60}")
        print(f"DEBUG: Frontend sent metadata:")
        print(f"  Date: {frontend_metadata.get('flight_date')}")
        print(f"  Call Sign: {frontend_metadata.get('call_sign')}")
        print(f"  PIC: {frontend_metadata.get('pic')}")
        print(f"{'='*60}\n")
        
        # Anomalies summary is optional (may be empty if no anomalies detected)
        if not anomalies_summary:
            print("  ℹ️ No anomalies summary provided - will save flight record only")
            anomalies_summary = {}  # Empty dict instead of error
        
        # Start with frontend metadata as fallback
        flight_metadata = frontend_metadata
        
        # NEW: Retrieve CVR results from session
        cvr_results = session.get('cvr_results', None)
        
        # CRITICAL: Re-extract flight metadata from Excel file to ensure current data
        excel_path = session.get('excel_path', None)
        excel_filename = session.get('excel_filename', None)
        
        if excel_path and os.path.exists(excel_path) and excel_filename:
            print(f"  📋 Re-extracting flight metadata from current Excel file...")
            print(f"     Excel: {excel_filename}")
            # Extract fresh metadata from the Excel file
            fresh_metadata = extract_flight_metadata_from_excel(excel_path, excel_filename)
            
            # Override frontend metadata with fresh data
            if fresh_metadata and fresh_metadata.get('flight_date'):
                print(f"  ✓ Using FRESH metadata from Excel:")
                print(f"     Date: {fresh_metadata.get('flight_date')} (was: {frontend_metadata.get('flight_date')})")
                print(f"     PIC: {fresh_metadata.get('pic')} (was: {frontend_metadata.get('pic')})")
                flight_metadata = fresh_metadata
                # Update session with fresh metadata
                session['flight_metadata'] = fresh_metadata
                session.modified = True
            else:
                print(f"  ⚠️ Warning: Could not extract metadata, using frontend data")
                print(f"     Fresh metadata was: {fresh_metadata}")
        else:
            print(f"  ⚠️ Warning: Excel path not in session, using frontend metadata")
            print(f"     excel_path: {excel_path}")
            print(f"     excel_filename: {excel_filename}")
        
        # NEW: Extract exceedances from Excel file
        exceedances_list = []
        
        if excel_path and os.path.exists(excel_path):
            print(f"  📊 Extracting exceedances from Excel file...")
            exceedances_list = extract_exceedances_from_excel(excel_path)
        else:
            print(f"  ⚠️ Warning: Excel path not found in session, skipping exceedances")
        
        # NEW: Extract compliance data from Excel Summary sheet
        compliance_data_from_excel = None
        missed_checks_from_excel = None
        
        if excel_path and os.path.exists(excel_path):
            print(f"  📊 Extracting compliance data from Excel file...")
            compliance_data_from_excel = extract_compliance_from_excel(excel_path)
            
            # Also extract individual missed checks from the checklist sheet
            print(f"  📊 Extracting missed checks from Excel checklist...")
            missed_checks_from_excel = extract_missed_checks_from_excel(excel_path)
            
            # If we found compliance data in Excel, update or create cvr_results
            if compliance_data_from_excel and compliance_data_from_excel['has_cvr_data']:
                if not cvr_results:
                    # Create cvr_results from Excel data
                    cvr_results = {
                        'compliance_percent': compliance_data_from_excel['compliance_percentage'],
                        'not_complied_count': compliance_data_from_excel['checks_not_complied'],
                        'results': [],  # Will be populated below
                        'checklist_type_id': missed_checks_from_excel.get('checklist_type_id', 1) if missed_checks_from_excel else 1,
                        'sheet_name': missed_checks_from_excel.get('sheet_name', 'Unknown') if missed_checks_from_excel else 'Unknown'
                    }
                    
                    # DEBUG: Show what we're creating
                    print(f"  🔍 DEBUG: Created cvr_results with:")
                    print(f"     compliance_percent: {cvr_results['compliance_percent']}")
                    print(f"     not_complied_count: {cvr_results['not_complied_count']}")
                    print(f"     checklist_type_id: {cvr_results['checklist_type_id']}")
                    
                    # Add missed checks if available
                    if missed_checks_from_excel and missed_checks_from_excel['missed_checks']:
                        # Convert to the format expected by the database
                        # Format: (status, item, score, matched_text, excel_row)
                        cvr_results['results'] = [
                            ('FAIL', item, score, '', excel_row)  # matched_text empty as we don't have it
                            for item, score, excel_row in missed_checks_from_excel['missed_checks']
                        ]
                        print(f"  ✓ Added {len(cvr_results['results'])} missed checks to cvr_results")
                    
                    print(f"  ✓ Created cvr_results from Excel data")
                else:
                    # Update existing cvr_results with Excel data (Excel is source of truth)
                    cvr_results['compliance_percent'] = compliance_data_from_excel['compliance_percentage']
                    cvr_results['not_complied_count'] = compliance_data_from_excel['checks_not_complied']
                    
                    # DEBUG: Show what we're updating
                    print(f"  🔍 DEBUG: Updated cvr_results with:")
                    print(f"     compliance_percent: {cvr_results['compliance_percent']}")
                    print(f"     not_complied_count: {cvr_results['not_complied_count']}")
                    
                    # Update missed checks if available from Excel
                    if missed_checks_from_excel and missed_checks_from_excel['missed_checks']:
                        cvr_results['results'] = [
                            ('FAIL', item, score, '', excel_row)
                            for item, score, excel_row in missed_checks_from_excel['missed_checks']
                        ]
                        cvr_results['checklist_type_id'] = missed_checks_from_excel.get('checklist_type_id', 1)
                        cvr_results['sheet_name'] = missed_checks_from_excel.get('sheet_name', 'Unknown')
                        print(f"  ✓ Updated cvr_results with {len(cvr_results['results'])} missed checks from Excel")
                    
                    print(f"  ✓ Updated cvr_results with Excel compliance data")
        
        print(f"💾 Saving to database: Flight {flight_metadata.get('call_sign', 'N/A')}")
        if cvr_results:
            print(f"  - Including CVR results: {cvr_results['not_complied_count']} missed checks")
        if exceedances_list:
            print(f"  - Including exceedances: {len(exceedances_list)} parameters exceeded")
        
        # Convert anomalies_summary from dict format back to tuple keys
        # Frontend sends: {"Fcp_when airborne": 12, "IAS_before takeoff": 5, ...}
        # Backend needs: {("Fcp", "when airborne"): 12, ("IAS", "before takeoff"): 5, ...}
        anomalies_dict = {}
        for key, value in anomalies_summary.items():
            # Split by last underscore to separate param from phase
            # Handle phases with underscores like "when_airborne"
            parts = key.rsplit('_', 1) if '_' in key else [key, 'unknown']
            if len(parts) == 2:
                param, phase = parts
                # Replace underscores in phase back to spaces
                phase = phase.replace('_', ' ')
                anomalies_dict[(param, phase)] = value
            else:
                print(f"Warning: Could not parse anomaly key: {key}")
        
        # Convert flight_date from string to date object if needed
        if isinstance(flight_metadata.get('flight_date'), str):
            try:
                flight_metadata['flight_date'] = datetime.strptime(
                    flight_metadata['flight_date'], '%Y-%m-%d'
                ).date()
            except ValueError:
                # Try alternative format
                flight_metadata['flight_date'] = datetime.strptime(
                    flight_metadata['flight_date'], '%Y/%m/%d'
                ).date()
        
        # Ensure required fields have defaults
        flight_metadata.setdefault('sortie', 1)
        if not flight_metadata.get('aircraft_id'):
            fallback_id = get_aircraft_id_by_call_sign(flight_metadata.get('call_sign'))
            flight_metadata['aircraft_id'] = fallback_id if fallback_id is not None else 3
        
        # Save to database using FlightAnalyzer method (now includes CVR results and exceedances)
        success = flight_analyzer._save_to_database(
            flight_metadata,
            anomalies_dict,
            cvr_results,      # CVR results with missed checks
            exceedances_list  # NEW: Exceedances from Summary sheet
        )
        
        if success:
            # Update session metadata to mark as saved to database
            if 'analysis_metadata' in session:
                session['analysis_metadata']['saved_to_database'] = True
                session.modified = True
            
            # Also update global variable for current view
            global latest_anomaly_report
            if latest_anomaly_report:
                latest_anomaly_report['saved_to_database'] = True
            
            print(f"✓ Successfully saved to database")
            return jsonify({
                'success': True,
                'flight_id': flight_metadata.get('flight_id')
            })
        else:
            # NEW: pull the real reason from FlightAnalyzer instead of a generic message
            detail = getattr(flight_analyzer, 'last_save_error', None) or 'Unknown error (check app_errors.log)'
            logger.error(f"save_to_database returned False. Detail: {detail}")
            logger.error(f"Flight metadata: {flight_metadata}")
            return jsonify({'success': False, 'error': detail}), 500
            
    except Exception as e:
        print(f"❌ Error in save_to_database: {e}")
        import traceback
        traceback.print_exc()
        logger.error(f"Exception in save_to_database route: {e}")
        logger.error(traceback.format_exc())
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/download_updated_excel/<filename>', methods=['GET'])
def download_updated_excel(filename):
    """
    NEW ENDPOINT: Download the updated Excel file with compliance results.
    """
    try:
        secure_filename_download = secure_filename(filename)
        return send_from_directory(
            directory=app.config['COMPLIANCE_EXCEL_OUTPUT'],
            path=secure_filename_download,
            as_attachment=True
        )
    except Exception as e:
        print(f"Error downloading file: {e}")
        return jsonify({'error': str(e)}), 404


# ============================================================================
# UPDATED ENDPOINTS
# ============================================================================

@app.route("/analyze_flight_anomalies", methods=["POST"])
def analyze_flight_anomalies():
    """
    UPDATED: Analyze flight data for anomalies using the "Clean Data" sheet.
    NO LONGER accepts add_to_training parameter - user decides on report page.
    """
    global latest_anomaly_report
    
    try:
        data = request.get_json()
        excel_filename = data.get('excel_filename')
        sheet_name = data.get('sheet_name', 'Clean Data')
        
        # REMOVED: add_to_training parameter - user decides later on report page
        
        if not excel_filename:
            return jsonify({'error': 'No Excel filename provided'}), 400
        
        # The Excel file is in COMPLIANCE_EXCEL_OUTPUT from the compliance check
        excel_path = os.path.join(COMPLIANCE_EXCEL_OUTPUT, excel_filename)
        
        if not os.path.exists(excel_path):
            return jsonify({'error': f'Excel file not found: {excel_filename}'}), 404
        
        print(f"\n{'='*60}")
        print(f"📊 ANALYZING FLIGHT DATA")
        print(f"{'='*60}")
        print(f"Excel file: {excel_path}")
        print(f"Sheet name: {sheet_name}")
        
        # IMPORTANT: Re-extract flight metadata from the Excel file to ensure current data
        # The Excel file now has updated Summary sheet with compliance data
        print("\\n📋 Re-extracting flight metadata from Excel file...")
        flight_metadata = extract_flight_metadata_from_excel(excel_path, excel_filename)
        
        # Override with session data if present (user may have manually entered data)
        session_metadata = session.get('flight_metadata', None)
        if session_metadata:
            print("  Merging with session metadata (preserving manual entries)...")
            # Preserve manual entries from session, but use Excel data as base
            # NOTE: 'aircraft_id' is deliberately NOT in this list. It is
            # always re-derived from the filename's call sign by the fresh
            # extraction above. A stale session value (e.g. a previous
            # flight's aircraft_id=3) is "valid" by the check below and
            # would silently clobber the correct value.
            for key in ['pic', 'sic', 'fe', 'sortie', 'flight_date']:
                if key in session_metadata and session_metadata[key] not in ['UNK', None, '']:
                    # Only override if session has valid data
                    if key == 'flight_date':
                        # Ensure consistent format
                        if isinstance(session_metadata[key], str):
                            flight_metadata[key] = session_metadata[key]
                    else:
                        flight_metadata[key] = session_metadata[key]

            # Carry over aircraft_id ONLY when the user explicitly set it via the form.
            if session_metadata.get('aircraft_id_manual') and session_metadata.get('aircraft_id'):
                flight_metadata['aircraft_id'] = session_metadata['aircraft_id']
                flight_metadata['aircraft_id_manual'] = True
                print(f"  Using manually overridden aircraft_id={flight_metadata['aircraft_id']}")
        
        # Update session with fresh metadata
        session['flight_metadata'] = flight_metadata
        session.modified = True
        
        print(f"✓ Current flight metadata: {flight_metadata}")
        
        # Fallback defaults if extraction failed completely
        if not flight_metadata or flight_metadata.get('pic') == 'UNK':
            # Try to extract info from filename or use defaults
            print("⚠ Warning: No flight metadata found in session, using defaults")
            flight_metadata = {
                'flight_date': date.today(),
                'pic': 'UNK',
                'sic': 'UNK',
                'fe': 'UNK',
                'sortie': 1,
                'aircraft_id': 3,  # last-resort literal: PIC extraction failed entirely, nothing to look up
                'call_sign': 'UNK'
            }
        else:
            # Ensure call_sign exists (might be from old session)
            flight_metadata.setdefault('call_sign', 'UNK')
            print(f"Flight metadata: {flight_metadata}")
        
        # Analyze the flight WITHOUT auto-actions (UPDATED)
        if DATABASE_ENABLED:
            results = flight_analyzer.analyze_flight(
                excel_path=excel_path,
                sheet_name=sheet_name,
                flight_metadata=flight_metadata,
                interactive=False,
                auto_add_to_training=False,  # NEW: Don't auto-add to training
                auto_save_to_db=False         # NEW: Don't auto-save to database
            )
        else:
            # Old FlightAnalyzer without database support
            results = flight_analyzer.analyze_flight(
                excel_path=excel_path,
                sheet_name=sheet_name,
                add_to_training=False  # Don't auto-add to training
            )
        
        if 'error' in results:
            return jsonify({'error': results['error']}), 500
        
        # Add analysis timestamp and status flags
        results['analysis_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        results['added_to_training'] = False  # NEW: User hasn't added yet
        results['saved_to_database'] = False  # NEW: User hasn't saved yet
        
        # Store in GLOBAL only (not session - too large for cookies!)
        # Session cookies have 4KB limit, visualization data is 100KB+
        latest_anomaly_report = results
        
        # Store only essential info in session (no visualization_data)
        session['analysis_metadata'] = {
            'flight_id': results.get('flight_id'),
            'analysis_date': results['analysis_date'],
            'added_to_training': False,
            'saved_to_database': False
        }
        
        # NEW: Store Excel path for later exceedances extraction
        session['excel_path'] = excel_path
        session['excel_filename'] = excel_filename
        
        session.modified = True
        
        print(f"✓ Analysis complete for Flight {results.get('flight_id', 'unknown')}")
        print(f"{'='*60}\n")
        
        # Return success - frontend will redirect to /anomaly_report
        return jsonify({'success': True})
        
    except Exception as e:
        print(f"❌ Error in flight anomaly analysis: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500


@app.route("/anomaly_report", methods=["GET"])
def view_anomaly_report():
    """
    UPDATED: Display the anomaly report with new action buttons.
    Uses global variable (not session - too large for cookies).
    """
    global latest_anomaly_report
    
    # Use global variable for main data
    report_data = latest_anomaly_report
    
    if not report_data:
        return "No anomaly report available. Please run an analysis first.", 404
    
    # Get status flags from session if available
    session_metadata = session.get('analysis_metadata', {})
    
    # Update status flags from session (user actions)
    if session_metadata:
        report_data['added_to_training'] = session_metadata.get('added_to_training', False)
        report_data['saved_to_database'] = session_metadata.get('saved_to_database', False)
    
    # Get flight metadata from session
    flight_metadata = session.get('flight_metadata', None)
    
    # Prepare enhanced report data for template
    enhanced_report = {
        'flight_id': report_data.get('flight_id'),
        'analysis_date': report_data.get('analysis_date', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
        'total_data_points': report_data.get('total_data_points', 0),
        'anomaly_count': report_data.get('total_anomalies', 0),
        'anomaly_percentage': float(report_data.get('anomaly_percentage', 0.0)),  # Ensure it's a float
        'anomalies': report_data.get('anomalies', []),
        'anomalies_by_param_phase': report_data.get('anomalies_by_param_phase', {}),  # NEW
        'phases_summary': report_data.get('phases_summary', {}),
        'visualization_data': report_data.get('visualization_data', {}),
        'flight_metadata': flight_metadata,  # NEW
        'added_to_training': report_data.get('added_to_training', False),  # NEW
        'saved_to_database': report_data.get('saved_to_database', False),  # NEW
        'database_enabled': DATABASE_ENABLED,  # NEW - Show database button only if available
        'total_historical_flights': (
            flight_analyzer.historical_data['flight_id'].nunique()
            if hasattr(flight_analyzer, 'historical_data') and not flight_analyzer.historical_data.empty
            else 0
        )
    }
    
    # DEBUG: Print phases_summary structure to help diagnose issues
    print("\n🔍 DEBUG - phases_summary structure:")
    for phase, data in enhanced_report['phases_summary'].items():
        print(f"  Phase: {phase}")
        print(f"    Keys: {list(data.keys())}")
        print(f"    Values: {data}")
    
    return render_template("anomaly_report.html", report=enhanced_report)


@app.route("/", methods=["GET", "POST"])
def index():
    """
    UPDATED: Handle file upload and compliance checking.
    Now stores flight metadata in session and generates download URL.
    """
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

            # NEW: Extract flight metadata from Excel file and filename
            print("\n📋 Extracting flight metadata...")
            flight_metadata = extract_flight_metadata_from_excel(excel_file_path, excel_filename)
            
            # Override with form data if provided (optional form fields)
            if request.form.get('flight_date'):
                flight_metadata['flight_date'] = request.form.get('flight_date')
            if request.form.get('pic'):
                flight_metadata['pic'] = request.form.get('pic').upper()
            if request.form.get('sic'):
                flight_metadata['sic'] = request.form.get('sic').upper()
            if request.form.get('fe'):
                flight_metadata['fe'] = request.form.get('fe').upper()
            if request.form.get('sortie'):
                flight_metadata['sortie'] = int(request.form.get('sortie'))
            if request.form.get('aircraft_id'):
                flight_metadata['aircraft_id'] = int(request.form.get('aircraft_id'))
                # Flag this as an explicit user override. Without this flag,
                # aircraft_id is always re-derived from the filename's call
                # sign and must NOT be inherited from session (see merges below).
                flight_metadata['aircraft_id_manual'] = True
            
            # Store in session
            session['flight_metadata'] = flight_metadata
            print(f"✓ Flight metadata: {flight_metadata}")

            # Audio Processing
            if len(saved_audio_paths) == 1:
                concatenated_audio_path = saved_audio_paths[0]
                print(f"Using single audio file: {os.path.basename(concatenated_audio_path)}")
            else:
                concatenated_audio_path = concatenate_audio_files(
                    saved_audio_paths, output_file_name, app.config["UPLOAD_FOLDER"]
                )
                if not concatenated_audio_path:
                    raise Exception("Audio concatenation failed.")
                print(f"Concatenated audio saved: {os.path.basename(concatenated_audio_path)}")

            cleaned_audio_path = preprocess_audio(concatenated_audio_path)
            if not cleaned_audio_path:
                raise Exception("Audio preprocessing failed.")

            # Transcribe audio
            print("Transcribing audio with Whisper...")
            transcript = transcribe_audio(cleaned_audio_path, output_file_name)
            print(f"Transcription complete: {len(transcript.split())} words")

            # Load Checklist and Check Compliance
            df, checklist, row_positions = load_checklist(excel_file_path, sheet_name)
            print(f"Checking compliance against {len(checklist)} checklist items...")
            results = check_compliance(transcript, checklist, threshold)
            
            # Map results to Excel row positions for database storage
            results_with_positions = []
            for i, (status, item, score, matched_text) in enumerate(results):
                excel_row = row_positions.get(i, i + 2)  # Fallback to i+2 if missing
                results_with_positions.append((status, item, score, matched_text, excel_row))
            
            print(f"  ✓ Mapped {len(results_with_positions)} results to Excel row positions")

            # Calculate compliance statistics
            passed_count = sum(1 for r in results if r[0] == "PASS")
            total_checks = len(results)
            compliance_percent = round((passed_count / total_checks) * 100, 1) if total_checks else 0
            not_complied_count = total_checks - passed_count

            # Update Excel with compliance results
            final_excel_output_path = update_excel(
                excel_file_path, results, sheet_name, not_complied_count, compliance_percent
            )

            # Save compliance report
            save_compliance_report(results, output_file_name)

            # FIXED: Re-extract metadata from the UPDATED Excel file
            # This ensures we get the correct compliance data that was just written
            print("\n📋 Re-extracting flight metadata from updated Excel file...")
            flight_metadata_updated = extract_flight_metadata_from_excel(
                final_excel_output_path, 
                os.path.basename(final_excel_output_path)
            )
            
            # Preserve any form overrides from the initial extraction.
            # 'aircraft_id' excluded on purpose: the updated file keeps the
            # same base filename, so the re-extraction above already resolved
            # it from the call sign. Only an explicit form override is kept.
            for key in ['flight_date', 'pic', 'sic', 'fe', 'sortie']:
                if key in flight_metadata and flight_metadata[key] != 'UNK':
                    flight_metadata_updated[key] = flight_metadata[key]

            if flight_metadata.get('aircraft_id_manual') and flight_metadata.get('aircraft_id'):
                flight_metadata_updated['aircraft_id'] = flight_metadata['aircraft_id']
                flight_metadata_updated['aircraft_id_manual'] = True
            
            # Update session with the corrected metadata
            flight_metadata = flight_metadata_updated
            session['flight_metadata'] = flight_metadata
            
            # Also store CVR results for later database save
            # Map sheet_name to checklist_type_id
            checklist_type_map = {
                'STARTING WITH AC-GPU CHECKLIST': 1,
                'STARTING WITH DC-GPU CHECKLIST': 2,
                'STARTING WITHOUT GPU CHECKLIST': 3
            }
            
            checklist_type_id = checklist_type_map.get(sheet_name, 1)
            
            session['cvr_results'] = {
                'results': [(r[0], r[1], r[2], r[3], r[4]) for r in results_with_positions],  # (status, item, score, matched_text, excel_row)
                'compliance_percent': compliance_percent,
                'not_complied_count': not_complied_count,
                'checklist_type_id': checklist_type_id,  # Store which checklist was used
                'sheet_name': sheet_name  # Store sheet name for reference
            }
            
            print(f"✓ Updated flight metadata with compliance data from Excel:")
            print(f"  - Compliance: {compliance_percent}%")
            print(f"  - Checks failed: {not_complied_count}")
            print(f"  - Checklist type: {sheet_name} (ID: {checklist_type_id})")

            # Return results with Excel filename for later anomaly analysis
            updated_excel_filename = os.path.basename(final_excel_output_path)
            
            # NEW: Generate download URL for the updated Excel file
            download_url = url_for('download_updated_excel', filename=updated_excel_filename, _external=True)

            print(f"\n✓ Compliance check complete:")
            print(f"  - Overall compliance: {compliance_percent}%")
            print(f"  - Checks failed: {not_complied_count}")
            print(f"  - Updated Excel: {updated_excel_filename}\n")

            return jsonify({
                "results": results,
                "compliance_percent": compliance_percent,
                "not_complied_count": not_complied_count,
                "excel_updated": True,
                "updated_excel_filename": updated_excel_filename,
                "download_excel_url": download_url,  # NEW
                "sheet_name": sheet_name
            })

        except Exception as e:
            print(f"❌ An error occurred: {e}")
            import traceback
            traceback.print_exc()
            return jsonify({"error": f"Error processing files: {e}"}), 500
        finally:
            # Clean up temporary directory
            if os.path.exists(temp_request_dir):
                shutil.rmtree(temp_request_dir)

    return render_template("index.html")



# ============================================================================
# NEW: DASHBOARD ROUTES
# ============================================================================

@app.route('/dashboard')
def dashboard():
    """
    NEW ROUTE: Main Dashboard with KPIs and visualizations.
    Displays flight safety metrics filtered by date range.
    """
    # Get filter parameters from request
    date_range_type = request.args.get('date_range', 'this_quarter')
    compare = request.args.get('compare', 'false') == 'true'
    custom_start = request.args.get('custom_start')
    custom_end = request.args.get('custom_end')
    
    # Calculate date ranges
    start_date, end_date, prev_start_date, prev_end_date = get_date_range(
        date_range_type, custom_start, custom_end
    )
    
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    
    # ===== FETCH KPIs =====
    kpis = {}
    
    # 1. Total Flights
    cursor.execute("""
        SELECT COUNT(*) as total_flights
        FROM flights
        WHERE flight_date BETWEEN %s AND %s
    """, (start_date, end_date))
    kpis['total_flights'] = cursor.fetchone()['total_flights']
    
    if compare:
        cursor.execute("""
            SELECT COUNT(*) as prev_total_flights
            FROM flights
            WHERE flight_date BETWEEN %s AND %s
        """, (prev_start_date, prev_end_date))
        prev_flights = cursor.fetchone()['prev_total_flights']
        kpis['flights_change'] = calculate_percentage_change(kpis['total_flights'], prev_flights)
    
    # 2. Total Exceedances
    cursor.execute("""
        SELECT SUM(continuous_exceedances + discrete_exceedances) as total_exceedances
        FROM flights
        WHERE flight_date BETWEEN %s AND %s
    """, (start_date, end_date))
    kpis['total_exceedances'] = cursor.fetchone()['total_exceedances'] or 0
    
    if compare:
        cursor.execute("""
            SELECT SUM(continuous_exceedances + discrete_exceedances) as prev_total_exceedances
            FROM flights
            WHERE flight_date BETWEEN %s AND %s
        """, (prev_start_date, prev_end_date))
        prev_exc = cursor.fetchone()['prev_total_exceedances'] or 0
        kpis['exceedances_change'] = calculate_percentage_change(kpis['total_exceedances'], prev_exc)
    
    # 3. Total Anomalies
    cursor.execute("""
        SELECT SUM(anomalies) as total_anomalies
        FROM flights
        WHERE flight_date BETWEEN %s AND %s
    """, (start_date, end_date))
    kpis['total_anomalies'] = cursor.fetchone()['total_anomalies'] or 0
    
    if compare:
        cursor.execute("""
            SELECT SUM(anomalies) as prev_total_anomalies
            FROM flights
            WHERE flight_date BETWEEN %s AND %s
        """, (prev_start_date, prev_end_date))
        prev_anom = cursor.fetchone()['prev_total_anomalies'] or 0
        kpis['anomalies_change'] = calculate_percentage_change(kpis['total_anomalies'], prev_anom)
    
    # 4. Average Compliance Rate
    cursor.execute("""
        SELECT AVG(compliance_percentage) as avg_compliance
        FROM flights
        WHERE flight_date BETWEEN %s AND %s
        AND compliance_percentage IS NOT NULL
    """, (start_date, end_date))
    result = cursor.fetchone()
    kpis['avg_compliance'] = round(result['avg_compliance'], 1) if result['avg_compliance'] else 0
    
    if compare:
        cursor.execute("""
            SELECT AVG(compliance_percentage) as prev_avg_compliance
            FROM flights
            WHERE flight_date BETWEEN %s AND %s
            AND compliance_percentage IS NOT NULL
        """, (prev_start_date, prev_end_date))
        prev_result = cursor.fetchone()
        prev_comp = round(prev_result['prev_avg_compliance'], 1) if prev_result['prev_avg_compliance'] else 0
        kpis['compliance_change'] = calculate_percentage_change(kpis['avg_compliance'], prev_comp)
    
    # ===== FETCH RECENT FLIGHTS =====
    cursor.execute("""
        SELECT 
            f.id,
            f.flight_date,
            f.PIC,
            f.SIC,
            a.call_sign as aircraft_name,
            f.sortie,
            f.compliance_percentage,
            f.checks_not_complied,
            (f.continuous_exceedances + f.discrete_exceedances) as exceedance_count,
            f.anomalies as anomaly_count,
            CASE 
                WHEN (f.continuous_exceedances + f.discrete_exceedances) = 0 
                     AND f.anomalies = 0 
                     AND (f.compliance_percentage >= 95 OR f.compliance_percentage IS NULL) 
                THEN 'green'
                WHEN (f.continuous_exceedances + f.discrete_exceedances) > 0 
                     OR f.anomalies > 3 
                     OR (f.compliance_percentage < 90 AND f.compliance_percentage IS NOT NULL) 
                THEN 'red'
                ELSE 'yellow'
            END as status_color
        FROM flights f
        LEFT JOIN aircrafts a ON f.aircraft_id = a.id
        WHERE f.flight_date BETWEEN %s AND %s
        ORDER BY f.flight_date DESC, f.sortie DESC
        LIMIT 20
    """, (start_date, end_date))
    recent_flights = cursor.fetchall()
    
    # ===== FETCH TREND DATA =====
    cursor.execute("""
        SELECT 
            DATE_FORMAT(f.flight_date, '%%Y-%%m') as month,
            COUNT(f.id) as total_flights,
            AVG(COALESCE(f.compliance_percentage, 0)) as avg_compliance,
            SUM(CASE WHEN (f.continuous_exceedances + f.discrete_exceedances) > 0 THEN 1 ELSE 0 END) as flights_with_exceedances
        FROM flights f
        WHERE f.flight_date BETWEEN %s AND %s
        GROUP BY DATE_FORMAT(f.flight_date, '%%Y-%%m')
        ORDER BY month
    """, (start_date, end_date))
    trend_data = cursor.fetchall()
    
    # Prepare trend data for charts
    months = [row['month'] for row in trend_data]
    total_flights_trend = [row['total_flights'] for row in trend_data]
    avg_compliance_trend = [round(row['avg_compliance'], 1) for row in trend_data]
    flights_with_issues = [row['flights_with_exceedances'] for row in trend_data]
    
    cursor.close()
    conn.close()
    
    return render_template('dashboard.html',
        kpis=kpis,
        recent_flights=recent_flights,
        start_date=start_date,
        end_date=end_date,
        prev_start_date=prev_start_date,
        prev_end_date=prev_end_date,
        compare=compare,
        date_range_type=date_range_type,
        months=months,
        total_flights_trend=total_flights_trend,
        avg_compliance_trend=avg_compliance_trend,
        flights_with_issues=flights_with_issues
    )


@app.route('/flights_list')
def flights_list():
    """
    NEW ROUTE: Full flight list with pagination and export capabilities.
    Allows browsing all flights with filtering and data export.
    """
    date_range_type = request.args.get('date_range', 'this_quarter')
    custom_start = request.args.get('custom_start')
    custom_end = request.args.get('custom_end')
    page = int(request.args.get('page', 1))
    page_size = int(request.args.get('page_size', 50))
    export_format = request.args.get('export')
    
    start_date, end_date, _, _ = get_date_range(date_range_type, custom_start, custom_end)
    
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    
    # Count total flights for pagination
    cursor.execute("""
        SELECT COUNT(*) as total_count
        FROM flights
        WHERE flight_date BETWEEN %s AND %s
    """, (start_date, end_date))
    total_count = cursor.fetchone()['total_count']
    total_pages = (total_count + page_size - 1) // page_size
    
    # Fetch flights for current page
    offset = (page - 1) * page_size
    cursor.execute("""
        SELECT 
            f.id,
            f.flight_date,
            f.PIC,
            f.SIC,
            f.FE,
            a.call_sign as aircraft_name,
            f.sortie,
            f.compliance_percentage,
            f.checks_not_complied,
            (f.continuous_exceedances + f.discrete_exceedances) as exceedance_count,
            f.anomalies as anomaly_count,
            CASE 
                WHEN (f.continuous_exceedances + f.discrete_exceedances) = 0 
                     AND f.anomalies = 0 
                     AND (f.compliance_percentage >= 95 OR f.compliance_percentage IS NULL) 
                THEN 'green'
                WHEN (f.continuous_exceedances + f.discrete_exceedances) > 0 
                     OR f.anomalies > 3 
                     OR (f.compliance_percentage < 90 AND f.compliance_percentage IS NOT NULL) 
                THEN 'red'
                ELSE 'yellow'
            END as status_color
        FROM flights f
        LEFT JOIN aircrafts a ON f.aircraft_id = a.id
        WHERE f.flight_date BETWEEN %s AND %s
        ORDER BY f.flight_date DESC, f.sortie DESC
        LIMIT %s OFFSET %s
    """, (start_date, end_date, page_size, offset))
    flights = cursor.fetchall()
    
    cursor.close()
    conn.close()
    
    # Handle export requests
    if export_format in ['excel', 'csv']:
        df = pd.DataFrame(flights)
        
        if export_format == 'excel':
            filename = f'flights_{start_date}_to_{end_date}.xlsx'
            filepath = os.path.join(tempfile.gettempdir(), filename)
            df.to_excel(filepath, index=False, engine='openpyxl')
            return send_file(filepath, as_attachment=True, download_name=filename)
        
        elif export_format == 'csv':
            filename = f'flights_{start_date}_to_{end_date}.csv'
            filepath = os.path.join(tempfile.gettempdir(), filename)
            df.to_csv(filepath, index=False)
            return send_file(filepath, as_attachment=True, download_name=filename)
    
    # Render template for normal page view
    return render_template('flights_list.html',
        flights=flights,
        start_date=start_date,
        end_date=end_date,
        page=page,
        total_pages=total_pages,
        total_count=total_count,
        page_size=page_size,
        date_range_type=date_range_type
    )


if __name__ == "__main__":
    print("\n" + "="*60)
    print("🚁 MI-17 Flight Analysis System Starting...")
    print("="*60)
    print(f"Database support: {'✓ Enabled' if DATABASE_ENABLED else '✗ Disabled'}")
    print(f"Upload folder: {UPLOAD_FOLDER}")
    print(f"Flight data folder: {FLIGHT_DATA_FOLDER}")
    print(f"Historical flights loaded: {flight_analyzer.historical_data['flight_id'].nunique() if hasattr(flight_analyzer, 'historical_data') and not flight_analyzer.historical_data.empty else 0}")
    print(f"NEW: Dashboard available at /dashboard")
    print(f"NEW: Flights list available at /flights_list")
    print("="*60 + "\n")
    
    app.run(debug=True, host='0.0.0.0', port=5000)