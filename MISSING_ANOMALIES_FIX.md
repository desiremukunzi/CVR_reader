# 🔍 MISSING OBVIOUS ANOMALIES - DIAGNOSIS & FIX

## Your Problem

**PITCH chart shows clear outlier at ~time 2k, value ~10**
- Well outside historical range (grey points around -10 to +7)
- Current flight (blue) has spike to +10
- **NOT detected as anomaly** ❌

**Same issue on:** X_lat, X_long, PITCH, T1, T2

## Root Cause Analysis

### Current Settings:
```python
ANOMALY_CONTAMINATION_RATE = 0.001  # 0.1%
N_ESTIMATORS = 300
USE_STATISTICAL_FILTER = False
```

### Why Outliers Aren't Detected:

**1. Contamination Rate Too Conservative (0.1%)**
- With 10,000 data points, 0.001 means: **only 10 anomalies expected**
- If your flight has 10,000 points, the model will flag only the TOP 10 most anomalous
- That PITCH spike might be ranked #15 or #20 → Not flagged!

**2. Isolation Forest Limitations**
- Isolation Forest works by "isolating" points
- It ranks points by anomaly score, then picks the top X% (contamination rate)
- A very obvious outlier might still not make the cut if contamination is too low

**3. Historical Data May Include Similar Spikes**
- If training data already contains occasional spikes
- The model learned them as "rare but normal"

## The Solution: Hybrid Detection Approach

Use **BOTH methods simultaneously**:
1. **Isolation Forest** - Catches complex multivariate patterns
2. **Simple Statistical Bounds** - Catches obvious outliers (like your PITCH spike)

### Recommended Configuration:

```python
# Isolation Forest settings (more sensitive)
ANOMALY_CONTAMINATION_RATE = 0.01  # 1% - More realistic
N_ESTIMATORS = 300

# Statistical bounds (catches obvious outliers)
SIGMA_THRESHOLD = 3.5  # 3.5 standard deviations
USE_STATISTICAL_FILTER = False  # Keep False to avoid filtering out IF detections
USE_STATISTICAL_DETECTION = True  # NEW: Add statistical detection

# Hybrid approach: Detect with BOTH methods, combine results
USE_HYBRID_DETECTION = True  # Combines IF + Statistical
```

## Implementation Strategy

### Option 1: Increase Contamination Rate (Quick Fix)

```python
# In flight_analyzer.py line 12:
ANOMALY_CONTAMINATION_RATE = 0.01  # Increase from 0.001 to 0.01 (1%)
```

**Pros:** Simple, one-line change
**Cons:** May increase false positives slightly

### Option 2: Add Statistical Detection Layer (Better)

Add a second detection pass that catches obvious outliers:

```python
def detect_anomalies_hybrid(self, flight_df, param, phase):
    \"\"\"
    Hybrid anomaly detection using both Isolation Forest and statistical bounds.
    \"\"\"
    # Step 1: Isolation Forest detection
    if_anomalies = self._detect_with_isolation_forest(flight_df, param, phase)
    
    # Step 2: Statistical detection (catches obvious outliers)
    stat_anomalies = self._detect_with_statistical_bounds(flight_df, param, phase)
    
    # Step 3: Combine (union of both methods)
    all_anomalies = if_anomalies.union(stat_anomalies)
    
    return all_anomalies
```

### Option 3: Adaptive Contamination Rate (Best)

Calculate contamination rate based on actual historical data variation:

```python
def calculate_adaptive_contamination(self, historical_data, param):
    \"\"\"
    Calculate contamination rate based on historical data characteristics.
    \"\"\"
    # Calculate how much variation exists
    std = historical_data[param].std()
    iqr = historical_data[param].quantile(0.75) - historical_data[param].quantile(0.25)
    
    # More variation → higher contamination rate
    if std > iqr * 2:  # High variation
        return 0.02  # 2%
    elif std > iqr:  # Moderate variation
        return 0.01  # 1%
    else:  # Low variation
        return 0.005  # 0.5%
```

## Quick Fix (Apply Now)

Replace lines 12-13 in `flight_analyzer.py`:

```python
# OLD:
ANOMALY_CONTAMINATION_RATE = 0.001  # 0.001% - Very low
N_ESTIMATORS = 300

# NEW:
ANOMALY_CONTAMINATION_RATE = 0.02  # 2% - More realistic for flight data
N_ESTIMATORS = 500  # More trees = better detection
```

## Testing After Fix

### Step 1: Clear Old Models
```bash
cd A:\CVR_reader\flight_data
del trained_anomaly_models.joblib
```

### Step 2: Restart Flask
```bash
cd A:\CVR_reader
python app.py
```

### Step 3: Retrain
1. Check "Add to training database"
2. Run analysis on historical flights first
3. Then analyze the problematic flight

### Step 4: Verify
**PITCH chart should now show:**
- Red X at time ~2k, value ~10 ✅

## Why This Works

### Contamination Rate Comparison:

| Rate | Expected Anomalies (per 10k points) | Use Case |
|------|--------------------------------------|----------|
| 0.0001 | 1 anomaly | Extremely clean data |
| 0.001 | 10 anomalies | Very clean data |
| 0.01 | 100 anomalies | Normal variation (RECOMMENDED) |
| 0.02 | 200 anomalies | High variation flight data |
| 0.05 | 500 anomalies | Very noisy data |

**For MI-17 flight data:** 1-2% is appropriate

### What 2% Means:
- 10,000 points → 200 anomalies detected
- Catches obvious outliers like PITCH spike
- Still conservative enough to avoid excessive false positives

## Advanced: Add Statistical Detection Layer

If increasing contamination rate alone doesn't work, add this method:

```python
def _detect_with_statistical_bounds(self, flight_df, param, phase):
    \"\"\"
    Simple statistical anomaly detection based on historical bounds.
    Catches obvious outliers that Isolation Forest might miss.
    \"\"\"
    # Get historical data
    hist_data = self.historical_data[
        (self.historical_data['phase'] == phase) & 
        (self.historical_data[param].notna())
    ]
    
    if len(hist_data) < 10:
        return pd.Index([])
    
    # Calculate bounds
    mean = hist_data[param].mean()
    std = hist_data[param].std()
    lower = mean - (3.5 * std)
    upper = mean + (3.5 * std)
    
    # Find outliers in current flight
    phase_data = flight_df[flight_df['phase'] == phase]
    outliers = phase_data[
        (phase_data[param] < lower) | 
        (phase_data[param] > upper)
    ].index
    
    return outliers
```

## Summary

**Problem:** Obvious outliers not detected
**Root Cause:** Contamination rate too conservative (0.1%)
**Quick Fix:** Increase to 2% (0.02)
**Better Fix:** Add statistical detection layer

**Apply the quick fix first, then test!**

### Configuration to Try:

```python
ANOMALY_CONTAMINATION_RATE = 0.02  # 2%
N_ESTIMATORS = 500
SIGMA_THRESHOLD = 3.5
USE_STATISTICAL_FILTER = False
```

This should catch your PITCH outlier at time 2k! 🎯
