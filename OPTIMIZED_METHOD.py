"""
OPTIMIZED _prepare_historical_data() METHOD
============================================

Replace the existing _prepare_historical_data() method in flight_analyzer.py
(around line 500) with this optimized version.

This fixes the browser performance issue with large datasets (180+ flights).
"""

def _prepare_historical_data(self):
    """
    Prepare historical data for overlay on charts with intelligent sampling.
    Samples large datasets to prevent browser performance issues.
    """
    # Configuration (add these at top of flight_analyzer.py if not present)
    MAX_HISTORICAL_POINTS_PER_PHASE = 2000  # Maximum points per phase
    HISTORICAL_DATA_SAMPLING = True  # Enable sampling
    
    historical_viz = {}
    phases = ['before takeoff', 'when airborne', 'after landing']
    
    print(f"\n📊 Preparing historical data for visualization...")
    total_hist_points = len(self.historical_data)
    print(f"   Total historical data points: {total_hist_points:,}")
    
    for param in PARAMETERS_TO_ANALYZE:
        if param not in self.historical_data.columns:
            continue
        
        historical_viz[param] = {}
        
        for phase in phases:
            phase_data = self.historical_data[
                (self.historical_data['phase'] == phase) & 
                (self.historical_data[param].notna())
            ].copy()
            
            if not phase_data.empty:
                # Smart sampling if data exceeds threshold
                if HISTORICAL_DATA_SAMPLING and len(phase_data) > MAX_HISTORICAL_POINTS_PER_PHASE:
                    # Calculate sampling ratio
                    sample_size = MAX_HISTORICAL_POINTS_PER_PHASE
                    sampling_ratio = sample_size / len(phase_data)
                    
                    print(f"   Sampling {param} '{phase}': {len(phase_data):,} → {sample_size} points ({sampling_ratio*100:.1f}%)")
                    
                    # Stratified sampling: preserve distribution
                    # Sort by value to ensure we capture min/max/distribution
                    phase_data_sorted = phase_data.sort_values(by=param)
                    
                    # Take every Nth point to maintain distribution
                    step = len(phase_data_sorted) // sample_size
                    if step < 1:
                        step = 1
                    
                    sampled_data = phase_data_sorted.iloc[::step].head(sample_size)
                    
                    # Ensure we include absolute min and max (outliers)
                    min_idx = phase_data[param].idxmin()
                    max_idx = phase_data[param].idxmax()
                    
                    if min_idx not in sampled_data.index:
                        sampled_data = pd.concat([sampled_data, phase_data.loc[[min_idx]]])
                    if max_idx not in sampled_data.index:
                        sampled_data = pd.concat([sampled_data, phase_data.loc[[max_idx]]])
                    
                    # Sort by time for proper plotting
                    sampled_data = sampled_data.sort_values(by='_time')
                    
                    phase_data = sampled_data
                
                historical_viz[param][phase] = {
                    'time': [float(x) for x in phase_data['_time'].tolist()],
                    'values': [float(x) for x in phase_data[param].tolist()]
                }
    
    print(f"   ✅ Historical data prepared for browser\n")
    return historical_viz
