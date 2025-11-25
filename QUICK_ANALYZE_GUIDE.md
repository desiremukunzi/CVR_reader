# 🚀 QUICK ANALYZE - BYPASS AUDIO COMPLIANCE

## What This Does

The `quick_analyze.py` script lets you **directly analyze a flight Excel file for anomalies** without going through the audio compliance check process.

**Skip:**
- ❌ Audio file upload
- ❌ Excel checklist upload
- ❌ Audio compliance processing
- ❌ Web interface navigation

**Go straight to:**
- ✅ Anomaly detection
- ✅ Anomaly report
- ✅ Automatic browser display

## Quick Start

### Step 1: Configure the Script

Open `quick_analyze.py` and set your file path:

```python
# Line 28 - Path to your flight Excel file:
FLIGHT_FILE_PATH = r'A:\Onedrive\RAF-61504\JULY\UNO-561P_04-07-25_1.xlsm'

# Line 31 - Excel sheet name:
SHEET_NAME = 'Clean Data'

# Line 34 - Add to training database?
ADD_TO_TRAINING = False  # False = just analyze, True = also add to training
```

### Step 2: Run the Analysis

```bash
cd A:\CVR_reader
python quick_analyze.py
```

### Step 3: Report Opens Automatically!

The anomaly report will automatically open in your default browser with:
- ✅ Summary statistics
- ✅ Anomalies by phase
- ✅ Detected anomalies table
- ✅ Parameter breakdown

## Example Output

```
======================================================================
  MI-17 FLIGHT ANOMALY ANALYSIS - QUICK ANALYSIS
======================================================================

📁 Flight File: UNO-561P_04-07-25_1.xlsm
📊 Sheet Name: Clean Data
🎯 Add to Training: No
🌐 Auto-open Browser: Yes
======================================================================

✅ File validation passed

🚀 Initializing Flight Analyzer...
✅ Loaded 27 trained models
   Historical database: 15 flights, 178,234 data points

======================================================================
  ANALYZING FLIGHT DATA
======================================================================

📊 Analyzing: UNO-561P_04-07-25_1.xlsm
   Sheet: Clean Data
   Add to training: No

Statistical Filter: Checking 3 Fcp anomalies in 'after landing' phase...
  Historical stats: mean=2.09, std=0.50
  Acceptance range: [-0.39, 4.58] (3.5σ)
  ✅ All 3 anomalies are real (none filtered)

Statistical Filter: Checking 3 Xcpl anomalies in 'after landing' phase...
  Historical stats: mean=3.93, std=1.23
  Acceptance range: [-2.23, 10.09] (3.5σ)
  ✅ All 3 anomalies are real (none filtered)

Statistical Filter: Checking 2 NZ anomalies in 'when airborne' phase...
  Historical stats: mean=0.07, std=0.23
  Acceptance range: [-1.07, 1.20] (3.5σ)
  ✅ All 2 anomalies are real (none filtered)

======================================================================
  ANALYSIS RESULTS
======================================================================

✅ Analysis complete!

📈 Flight Statistics:
   - Flight ID: 21
   - Total data points: 10,234
   - Anomalies detected: 8
   - Anomaly rate: 0.08%

📊 Anomalies by Phase:
   - Before takeoff: 0 anomalies (0.0%) out of 1,234 points
   - When airborne: 2 anomalies (0.02%) out of 8,567 points
   - After landing: 6 anomalies (1.37%) out of 433 points

🎯 Anomalies by Parameter:
   ✅ Fcp: 3 anomaly(ies)
   ✅ Xcpl: 3 anomaly(ies)
   ⚪ Pedals: No anomalies
   ⚪ X_lat: No anomalies
   ⚪ X_long: No anomalies
   ⚪ PITCH: No anomalies
   ✅ NZ: 2 anomaly(ies)
   ⚪ T1: No anomalies
   ⚪ T2: No anomalies

======================================================================
  GENERATING REPORT
======================================================================

📄 Basic report saved: reports\anomaly_report_flight_21_20251121_223045.html

🌐 Opening report in browser...
✅ Report opened in default browser

======================================================================
  🎉 ANALYSIS COMPLETE!
======================================================================
```

## Configuration Options

### Required Settings:

```python
# Path to your flight Excel file
FLIGHT_FILE_PATH = r'A:\path\to\your\flight.xlsm'

# Sheet name in Excel file
SHEET_NAME = 'Clean Data'

# Add to training database?
ADD_TO_TRAINING = False  # Default: False (just analyze)

# Auto-open report in browser?
AUTO_OPEN_BROWSER = True  # Default: True
```

### ADD_TO_TRAINING Options:

**False (Default)** - Quick Analysis:
- Just analyze the flight
- Don't add to historical database
- Models stay unchanged
- Use for: Quick checks, testing, anomaly investigation

**True** - Analyze and Train:
- Analyze the flight
- Add to historical database
- Retrain models with new data
- Use for: Building training database, adding normal flights

## Generated Report

### Report Location:
```
A:\CVR_reader\reports\
└── anomaly_report_flight_21_20251121_223045.html
```

### Report Contents:

1. **Header Section**
   - Flight ID
   - Filename
   - Generation timestamp

2. **Summary Statistics**
   - Total data points
   - Anomalies detected
   - Anomaly rate

3. **Phase Breakdown Table**
   - Data points per phase
   - Anomalies per phase
   - Anomaly rate per phase

4. **Detected Anomalies Table**
   - Parameter name
   - Phase
   - Time
   - Value

## Use Cases

### Use Case 1: Quick Anomaly Check
```python
FLIGHT_FILE_PATH = r'A:\Flights\new_flight.xlsm'
ADD_TO_TRAINING = False
```
**Result:** Analyze only, don't modify training database

### Use Case 2: Analyze and Add to Training
```python
FLIGHT_FILE_PATH = r'A:\Flights\normal_flight.xlsm'
ADD_TO_TRAINING = True
```
**Result:** Analyze + add to database + retrain models

### Use Case 3: Batch Analysis
```python
# Analyze multiple flights
flights = [
    r'A:\Flights\flight_01.xlsm',
    r'A:\Flights\flight_02.xlsm',
    r'A:\Flights\flight_03.xlsm'
]

for flight in flights:
    FLIGHT_FILE_PATH = flight
    ADD_TO_TRAINING = False
    # Run analysis
```

## Common Issues & Solutions

### Issue 1: "No trained models found"
```
❌ ERROR: No trained models found!
   You need to train models first using quick_train.py
```

**Solution:**
```bash
# First train models:
python quick_train.py

# Then analyze:
python quick_analyze.py
```

### Issue 2: "File does not exist"
```
❌ ERROR: File does not exist: A:\path\to\flight.xlsm
```

**Solution:**
- Check file path is correct
- Use `r'...'` for Windows paths
- Ensure file extension is correct

### Issue 3: "Sheet not found"
```
❌ ERROR: Sheet 'Clean Data' not found in Excel file
```

**Solution:**
- Open Excel file
- Check sheet name (case-sensitive!)
- Update `SHEET_NAME` in script

### Issue 4: Report doesn't open
```
📄 Report saved: reports\anomaly_report_flight_21.html
```

**Solution:**
- Manually navigate to `reports/` folder
- Double-click the HTML file
- Or set `AUTO_OPEN_BROWSER = True` in script

## Workflow Examples

### Workflow 1: Investigate Specific Flight
```bash
# 1. Edit quick_analyze.py
FLIGHT_FILE_PATH = r'A:\Flights\suspicious_flight.xlsm'
ADD_TO_TRAINING = False

# 2. Run analysis
python quick_analyze.py

# 3. Review report in browser
# 4. Check anomalies
```

### Workflow 2: Build Training Database
```bash
# 1. Edit quick_analyze.py
FLIGHT_FILE_PATH = r'A:\Flights\normal_flight_01.xlsm'
ADD_TO_TRAINING = True

# 2. Run for each normal flight
python quick_analyze.py

# Change file path, run again
FLIGHT_FILE_PATH = r'A:\Flights\normal_flight_02.xlsm'
python quick_analyze.py

# Repeat for all normal flights
```

### Workflow 3: Daily Flight Checks
```bash
# Morning routine:
# 1. Copy latest flight to A:\Flights\today.xlsm
# 2. Edit quick_analyze.py
FLIGHT_FILE_PATH = r'A:\Flights\today.xlsm'
ADD_TO_TRAINING = False

# 3. Run quick check
python quick_analyze.py

# 4. Review anomalies
# 5. If flight is normal, add to training:
ADD_TO_TRAINING = True
python quick_analyze.py
```

## Comparison: Web App vs Quick Analyze

### Web Application:
1. Start Flask server ⏱️
2. Open browser
3. Generate compliance report (audio check)
4. Click "Analyze for Anomalies"
5. Wait for processing
6. View report

**Time: ~5 minutes**

### Quick Analyze Script:
1. Edit file path in script
2. Run `python quick_analyze.py`
3. Report opens automatically

**Time: ~30 seconds** ⚡

## Advanced Usage

### Disable Auto-Open Browser
```python
AUTO_OPEN_BROWSER = False
```
Report still generated, but won't open automatically.

### Change Model Folder
```python
MODEL_DATA_FOLDER = 'custom_models_folder'
```
Use different models than default.

### Programmatic Analysis
```python
from quick_analyze import analyze_flight_direct

results = analyze_flight_direct(
    file_path=r'A:\Flights\test.xlsm',
    add_to_training=False,
    auto_open=False
)

if results:
    print(f"Anomalies: {results['anomaly_count']}")
```

## Tips

### ✅ Do:
- Use for quick anomaly checks
- Disable auto-open if analyzing many flights
- Set ADD_TO_TRAINING=True for normal flights only
- Review reports before adding to training

### ❌ Don't:
- Add anomalous flights to training (pollutes database)
- Forget to train models first (quick_train.py)
- Use without checking report output
- Analyze without historical models

## Prerequisites

1. **Trained models must exist:**
   ```bash
   python quick_train.py  # Run once to create models
   ```

2. **Excel file requirements:**
   - Must have "Clean Data" sheet (or specified sheet name)
   - Must have flight parameters (Fcp, Xcpl, PITCH, etc.)
   - Must have `iWOW`, `_time` columns

3. **Python environment:**
   - All dependencies installed (pandas, openpyxl, etc.)
   - flight_analyzer.py in same directory

## Output Files

```
A:\CVR_reader\
├── quick_analyze.py          ← The script
├── reports\                  ← Generated reports
│   ├── anomaly_report_flight_21_20251121_223045.html
│   └── anomaly_report_flight_22_20251121_224512.html
└── flight_data\              ← Models (if ADD_TO_TRAINING=True)
    ├── trained_anomaly_models.joblib
    └── all_flights_data.parquet
```

## Summary

**Quick analysis workflow:**

```bash
# 1. Set file path in quick_analyze.py
FLIGHT_FILE_PATH = r'A:\Flights\your_flight.xlsm'

# 2. Run analysis
python quick_analyze.py

# 3. Report opens automatically
# Done! ✅
```

**Benefits:**
- ⚡ 30 seconds vs 5 minutes
- 🚀 Direct to results
- 📊 Automatic report generation
- 🌐 Auto-open in browser
- 💾 Optional training database update

**Perfect for:**
- Quick daily flight checks
- Investigating specific flights
- Batch analysis runs
- Building training database
- Bypassing audio compliance check

**Your anomaly analysis is now just one command away!** 🎯🚀
