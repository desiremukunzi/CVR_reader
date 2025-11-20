# ✅ INSTALLATION COMPLETE - Final Steps

## Files Created Successfully

✅ **flight_analyzer.py** - Created at `A:\CVR_reader\flight_analyzer.py`
✅ **ROUTE_TO_ADD.txt** - Created at `A:\CVR_reader\ROUTE_TO_ADD.txt`
✅ **index.html** - Already correct, no changes needed!

## ONE MANUAL STEP REQUIRED

You need to add the route to your `app.py` file. Here's how:

### Step 1: Open your `app.py` file
Location: `A:\CVR_reader\app.py`

### Step 2: Find this section (around line 246):
```python
def save_compliance_report(results, output_file_name):
    """Save compliance results to text file."""
    # ... rest of the function ...
    print(f"Compliance report saved to: {report_path}")


@app.route("/", methods=["GET", "POST"])
def index():
```

### Step 3: Copy the entire content from `ROUTE_TO_ADD.txt` 
Open: `A:\CVR_reader\ROUTE_TO_ADD.txt`

### Step 4: Paste it between the two functions
So it becomes:
```python
def save_compliance_report(results, output_file_name):
    """Save compliance results to text file."""
    # ... rest of the function ...
    print(f"Compliance report saved to: {report_path}")


# PASTE THE ROUTE HERE (from ROUTE_TO_ADD.txt)


@app.route("/", methods=["GET", "POST"])
def index():
```

### Step 5: Save `app.py`

### Step 6: Install dependencies
```bash
pip install joblib pyarrow
```

### Step 7: Restart your application
```bash
cd A:\CVR_reader
python app.py
```

## How to Test

1. Navigate to http://localhost:5000
2. Upload your Excel file (with both checklist sheet and "Clean Data" sheet)
3. Upload audio files
4. Click "Generate Compliance Report"
5. ✅ After report shows, click "Analyze Flight for Anomalies"
6. ✅ Check console - should see:
   - "Analyzing flight data from: ..."
   - "Sheet name: Clean Data"
   - "Training Isolation Forest models..."
   - "Trained X models" (where X > 0)

## Your Excel File Must Have

- **Sheet 1**: Checklist sheet (e.g., "STARTING WITH AC-GPU CHECKLIST")
- **Sheet 2**: "Clean Data" with columns:
  - `_time` (or `_Time`)
  - `iWOW` (0 = airborne, 1 = on ground)
  - `Fcp`, `Xcpl`, `Pedals`, `X_lat`, `X_long`, `PITCH`, `NZ`, `T1`, `T2`

## If You Get Errors

### "Warning: 'iWOW' column not found"
- Check that your "Clean Data" sheet has `iWOW` column (exact spelling, case-sensitive)

### "Trained 0 models"
- Make sure you have at least 100 rows of flight data
- Check that parameter values vary (not all the same number)
- Verify column names match exactly

### "Sheet 'Clean Data' not found"
- Your Excel file doesn't have this sheet
- Either add the sheet or change line in the route where it says `sheet_name = data.get('sheet_name', 'Clean Data')`

## Summary

You're 99% done! Just need to:
1. Copy content from `ROUTE_TO_ADD.txt`
2. Paste it into `app.py` at the right location
3. Save and restart

That's it! 🎉
