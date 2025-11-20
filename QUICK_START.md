# Quick Start Guide

## ✅ What I've Done

1. Created `flight_analyzer.py` - The anomaly detection engine
2. Created `ROUTE_TO_ADD.txt` - The Flask route you need to add
3. Your `index.html` is already perfect!

## 🔧 What You Need To Do (5 minutes)

### 1. Add the Route to app.py

**Open:** `A:\CVR_reader\app.py`

**Find line ~246:**
```python
    print(f"Compliance report saved to: {report_path}")


@app.route("/", methods=["GET", "POST"])
```

**Insert between them:** Copy entire content from `ROUTE_TO_ADD.txt` and paste

**Result:**
```python
    print(f"Compliance report saved to: {report_path}")


@app.route("/analyze_flight_anomalies", methods=["POST"])
def analyze_flight_anomalies():
    # ... the whole function from ROUTE_TO_ADD.txt ...


@app.route("/", methods=["GET", "POST"])
```

### 2. Install Packages
```bash
pip install joblib pyarrow
```

### 3. Restart App
```bash
python app.py
```

## ✅ Test It

1. Upload Excel file (must have both sheets)
2. Upload audio
3. Generate report
4. Click "Analyze Flight for Anomalies"
5. New window opens with charts!

## 📋 Excel File Requirements

Your Excel file needs **TWO sheets**:

**Sheet 1:** "STARTING WITH AC-GPU CHECKLIST" (or DC-GPU/WITHOUT GPU)
- Column A: Checklist items

**Sheet 2:** "Clean Data"
- Columns: `_time`, `iWOW`, `Fcp`, `Xcpl`, `Pedals`, `X_lat`, `X_long`, `PITCH`, `NZ`, `T1`, `T2`
- `iWOW`: 0 = airborne, 1 = on ground

## 🎯 That's It!

Just add the route and you're done!
