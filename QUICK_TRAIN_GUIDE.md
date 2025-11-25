# 🚀 QUICK MODEL TRAINING GUIDE

## What This Does

The `quick_train.py` script lets you **rapidly train anomaly detection models** by pointing to a folder with your historical flight data Excel files.

**Instead of:**
- Manually uploading each flight through web interface ❌
- Clicking "Add to training" for each one ❌
- Waiting for models to retrain after each flight ❌

**You get:**
- Point to folder with all Excel files ✅
- Script processes all flights at once ✅
- Models trained in one go ✅

## Quick Start

### Step 1: Configure the Script

Open `quick_train.py` and set your folder path:

```python
# Line 28 - Set your training data folder:
TRAINING_DATA_FOLDER = r'A:\Onedrive\RAF-61504\JUNE\FLIGHTS\FOR_REPORT'

# Line 31 - Set your Excel sheet name (if different):
SHEET_NAME = 'Clean Data'
```

### Step 2: Run the Script

```bash
cd A:\CVR_reader
python quick_train.py
```

### Step 3: Watch the Progress

```
======================================================================
  MI-17 FLIGHT ANOMALY DETECTION - QUICK MODEL TRAINING
======================================================================

📁 Training Data Folder: A:\Onedrive\RAF-61504\JUNE\FLIGHTS\FOR_REPORT
📊 Sheet Name: Clean Data
💾 Models will be saved to: flight_data
======================================================================

✅ Found 15 Excel file(s) for training

🚀 Initializing Flight Analyzer...

======================================================================
  LOADING FLIGHT DATA
======================================================================

[1/15] Processing: UNO-561P_01-06-25.xlsm
   ✅ Added Flight 1 - 12453 data points

[2/15] Processing: UNO-561P_02-06-25.xlsm
   ✅ Added Flight 2 - 11892 data points

...

======================================================================
  TRAINING ANOMALY DETECTION MODELS
======================================================================

📊 Training data summary:
   - Total flights: 15
   - Total data points: 178,234
   - Parameters to analyze: Fcp, Xcpl, Pedals, X_lat, X_long, PITCH, NZ, T1, T2

Training Isolation Forest models...
  - Trained model for Fcp in 'before takeoff'
  - Trained model for Fcp in 'when airborne'
  - Trained model for Fcp in 'after landing'
  ...

Trained 27 models.

======================================================================
  TRAINING RESULTS
======================================================================

✅ Successfully processed: 15 flight(s)

📈 Training Statistics:
   - Models trained: 27
   - Unique flights: 15
   - Total data points: 178,234

📊 Data distribution by phase:
   - Before takeoff: 12,345 points (6.9%)
   - When airborne: 156,789 points (88.0%)
   - After landing: 9,100 points (5.1%)

🎯 Models trained per parameter:
   ✅ Fcp: 3 model(s)
   ✅ Xcpl: 3 model(s)
   ✅ Pedals: 3 model(s)
   ✅ X_lat: 3 model(s)
   ✅ X_long: 3 model(s)
   ✅ PITCH: 3 model(s)
   ✅ NZ: 3 model(s)
   ✅ T1: 3 model(s)
   ✅ T2: 3 model(s)

💾 Models saved to: A:\CVR_reader\flight_data\trained_anomaly_models.joblib
💾 Historical data saved to: A:\CVR_reader\flight_data\all_flights_data.parquet

======================================================================
  🎉 TRAINING COMPLETE!
======================================================================

✅ Your models are ready for use in the web application!
   Simply restart Flask and start analyzing flights.
```

### Step 4: Use Your Models

```bash
# Restart Flask
python app.py
```

Now when you analyze new flights, they'll be compared against your trained models!

## Features

### 1. Automatic Flight Detection
- Scans folder for all `.xlsx`, `.xlsm`, `.xls` files
- Processes them in order
- Shows progress for each file

### 2. Error Handling
- Skips empty sheets
- Reports missing sheets
- Continues even if some files fail

### 3. Existing Model Management
If models already exist, you'll see:
```
⚠️  WARNING: Existing models found!
   - 27 models already trained
   - 10 flights in database
   - 125,678 total data points

   Continue and ADD to existing data? (y/n):
```

**Type `y`** to add more training data
**Type `n`** to cancel

### 4. Detailed Statistics
Shows:
- ✅ Successful flights processed
- ❌ Failed flights (with reasons)
- 📊 Data distribution by phase
- 🎯 Models trained per parameter
- 💾 File locations

## Configuration Options

### Required Settings:

```python
# Path to your training data folder
TRAINING_DATA_FOLDER = r'A:\Onedrive\RAF-61504\JUNE\FLIGHTS\FOR_REPORT'

# Sheet name in Excel files (must be same across all files)
SHEET_NAME = 'Clean Data'

# Where to save models (default is 'flight_data')
MODEL_DATA_FOLDER = 'flight_data'
```

### Folder Structure Example:

```
A:\Onedrive\RAF-61504\JUNE\FLIGHTS\FOR_REPORT\
├── UNO-561P_01-06-25.xlsm
├── UNO-561P_02-06-25.xlsm
├── UNO-561P_03-06-25.xlsm
├── UNO-561P_04-06-25.xlsm
└── ... (more files)
```

Each file must have a sheet named `Clean Data` (or whatever you set in `SHEET_NAME`).

## Common Issues & Solutions

### Issue 1: "No Excel files found"
```
❌ ERROR: No Excel files found in: A:\path\to\folder
```

**Solution:**
- Check folder path is correct
- Ensure files have `.xlsx`, `.xlsm`, or `.xls` extension
- Check files aren't in a subfolder

### Issue 2: "Sheet not found"
```
[3/15] Processing: flight_03.xlsm
   ❌ Error: Sheet 'Clean Data' not found in file
```

**Solution:**
- Open the Excel file
- Check the sheet name (case-sensitive!)
- Update `SHEET_NAME` in the script to match

### Issue 3: "Empty sheet"
```
[5/15] Processing: flight_05.xlsm
   ⚠️  Warning: Sheet 'Clean Data' is empty - Skipping
```

**Solution:**
- File has no data in the specified sheet
- Script will skip it and continue
- Remove empty files from folder or fix them

### Issue 4: Script crashes with error

**Solution:**
- Check the error message printed
- Common causes:
  - Wrong column names in Excel
  - Corrupted Excel file
  - Missing `iWOW` column
- Fix the problematic file and run again

## Advanced Usage

### Clear Existing Models First

If you want to start fresh (delete old models):

```bash
cd A:\CVR_reader\flight_data
del trained_anomaly_models.joblib
del all_flights_data.parquet
```

Then run `quick_train.py` to rebuild from scratch.

### Train in Batches

**Batch 1 (Historical flights):**
```python
TRAINING_DATA_FOLDER = r'A:\Flights\2024\Historical'
python quick_train.py
```

**Batch 2 (Add more recent flights):**
```python
TRAINING_DATA_FOLDER = r'A:\Flights\2025\Recent'
python quick_train.py
# Answer 'y' when asked to add to existing data
```

### Check Training Data

After training, check what's in the database:

```python
# In Python console:
from flight_analyzer import FlightAnalyzer

analyzer = FlightAnalyzer()
print(f"Flights: {analyzer.historical_data['flight_id'].nunique()}")
print(f"Points: {len(analyzer.historical_data)}")
print(f"Models: {len(analyzer.trained_models)}")
```

## Output Files

After training, you'll have:

```
A:\CVR_reader\flight_data\
├── trained_anomaly_models.joblib     ← Trained ML models
└── all_flights_data.parquet          ← Historical flight data
```

**These files are used by the web application automatically!**

## Workflow

### Initial Training Setup:

1. **Collect historical flights** → Put all Excel files in one folder
2. **Set folder path** → Edit `TRAINING_DATA_FOLDER` in `quick_train.py`
3. **Run training** → `python quick_train.py`
4. **Start Flask** → `python app.py`
5. **Analyze new flights** → Use web interface

### Adding More Training Data:

1. **Put new Excel files in folder**
2. **Run training again** → `python quick_train.py`
3. **Answer 'y'** when asked to add to existing data
4. **Restart Flask** → Models updated automatically

### Starting Fresh:

1. **Delete old models** → Remove `.joblib` and `.parquet` files
2. **Run training** → `python quick_train.py`
3. **Restart Flask**

## Tips

### ✅ Do:
- Use at least 5-10 historical flights for good training
- Ensure all Excel files have same sheet name
- Check files open in Excel before training
- Keep a backup of working model files

### ❌ Don't:
- Mix different aircraft types in training
- Include incomplete or corrupted flights
- Forget to restart Flask after retraining
- Delete model files unless you want to retrain

## Expected Performance

### Small dataset (5 flights, ~50k points):
- Loading: ~10 seconds
- Training: ~20 seconds
- Total: ~30 seconds

### Medium dataset (15 flights, ~150k points):
- Loading: ~30 seconds
- Training: ~60 seconds
- Total: ~90 seconds

### Large dataset (50 flights, ~500k points):
- Loading: ~2 minutes
- Training: ~3 minutes
- Total: ~5 minutes

## Summary

**Quick training workflow:**

```bash
# 1. Set folder path in quick_train.py
# 2. Run training
python quick_train.py

# 3. Wait for completion
# 4. Restart Flask
python app.py

# 5. Analyze flights!
```

**That's it! Your models are trained and ready to detect anomalies! 🚀**
