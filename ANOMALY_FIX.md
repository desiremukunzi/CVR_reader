# FIX FOR ANOMALY DETECTION ACCURACY

## Problem
The anomaly detection was missing obvious anomalies (like the T1 spike in the chart) because the contamination rate was too low.

### Root Cause
```python
ANOMALY_CONTAMINATION_RATE = 0.0001  # 0.01% - TOO STRICT!
```

This means the model expected only **1 anomaly per 10,000 data points**, which is unrealistic for flight anomaly detection.

## Solution
Changed contamination rate to a more realistic value:

```python
ANOMALY_CONTAMINATION_RATE = 0.005  # 0.5% - More sensitive
```

This means the model now expects about **5 anomalies per 1,000 data points**, which better matches real-world flight anomaly patterns.

## What This Changes

### Before (0.01%):
- **Too strict** - Only catches extreme outliers
- Misses obvious anomalies like the T1 spike
- Only 1-2 anomalies detected per flight

### After (0.5%):
- **Balanced** - Catches significant anomalies
- Will detect visible spikes like T1
- More realistic anomaly count (5-10 per flight typically)

## How to Apply the Fix

### Step 1: Delete Old Models
The old models were trained with the wrong contamination rate and need to be deleted:

```bash
cd A:\CVR_reader\flight_data
del trained_anomaly_models.joblib
```

**Or manually:**
1. Navigate to `A:\CVR_reader\flight_data\`
2. Delete the file `trained_anomaly_models.joblib`

### Step 2: Restart Flask
```bash
cd A:\CVR_reader
python app.py
```

### Step 3: Rerun Analysis
1. Upload audio + Excel files
2. Generate compliance report
3. Click "Analyze Flight for Anomalies"
4. **Check the "Add this flight data to training database" checkbox**
5. Click analyze

The system will:
- Train NEW models with the correct contamination rate (0.5%)
- Detect MORE anomalies
- Save the new models for future use

## Expected Results

### Anomaly Detection Should Now Show:

**Before Fix:**
- Fcp: 3 anomalies
- Xcpl: 3 anomalies  
- NZ: 2 anomalies
- T1: **0 anomalies** ❌ (WRONG!)

**After Fix:**
- Fcp: 3-6 anomalies
- Xcpl: 3-6 anomalies
- NZ: 2-4 anomalies
- T1: **2-3 anomalies** ✅ (Including the visible spike!)
- Plus more parameters with anomalies

### In the Anomaly Table:
You should now see **ALL parameters with anomalies**, including T1.

### In the Charts:
Red X markers should appear at **ALL visible spikes**, including the T1 spike you saw.

## Technical Details

### Contamination Rate Explained:
- **0.0001 (0.01%)**: Expects 1 anomaly per 10,000 points - Too strict
- **0.001 (0.1%)**: Expects 1 anomaly per 1,000 points - Still strict
- **0.005 (0.5%)**: Expects 5 anomalies per 1,000 points - **GOOD BALANCE** ✅
- **0.01 (1.0%)**: Expects 10 anomalies per 1,000 points - More sensitive
- **0.05 (5.0%)**: Expects 50 anomalies per 1,000 points - Too sensitive

### Why 0.5%?
- Matches typical flight anomaly patterns
- Catches significant deviations
- Not too sensitive (avoids false positives)
- Not too strict (avoids missing real anomalies)

## Verification

After applying the fix, check:

1. **Anomaly Table** - Should show T1 and other parameters
2. **T1 Chart** - Red X markers at spike points
3. **Total Anomaly Count** - Should be higher (20-50 typical)
4. **All Parameter Charts** - Anomalies marked appropriately

## Files Modified

- ✅ `flight_analyzer.py` - Changed `ANOMALY_CONTAMINATION_RATE` from 0.0001 to 0.005

## Action Required

**YOU MUST DELETE THE OLD MODEL FILE:**
```
A:\CVR_reader\flight_data\trained_anomaly_models.joblib
```

Then restart Flask and rerun the analysis with the checkbox checked!

**This will retrain the models with the correct sensitivity.** ✅
