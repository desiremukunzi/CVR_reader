import pandas as pd
import numpy as np
from sklearn.ensemble import IsolationForest
import joblib
import os
from datetime import datetime

# Define the parameters for anomaly detection
PARAMETERS_TO_ANALYZE = ['Fcp', 'Xcpl', 'Pedals', 'X_lat', 'X_long', 'PITCH', 'NZ', 'T1', 'T2']

# Anomaly detection configuration
ANOMALY_CONTAMINATION_RATE = 0.01  # 1% - Very low rate to reduce false positives
N_ESTIMATORS = 500  # More trees = better detection of outliers

# Statistical anomaly filtering - eliminates false positives
# Points must be outside mean ± (SIGMA_THRESHOLD * std_dev) of historical data
SIGMA_THRESHOLD = 5  # 5 standard deviations (99.99994% confidence)
USE_STATISTICAL_FILTER = True  # Set to True to filter out borderline anomalies

# Visualization optimization
MAX_HISTORICAL_POINTS_PER_PHASE = 20000
HISTORICAL_DATA_SAMPLING = True

# File paths for persistence
MODEL_FILENAME = 'trained_anomaly_models.joblib'
HISTORICAL_DATA_FILENAME = 'all_flights_data.parquet'

class FlightAnalyzer:
    def __init__(self, data_folder='flight_data'):
        """
        Initialize the Flight Analyzer with persistent storage.
        
        Args:
            data_folder (str): Folder to store models and historical data
        """
        self.data_folder = data_folder
        os.makedirs(self.data_folder, exist_ok=True)
        
        self.model_path = os.path.join(self.data_folder, MODEL_FILENAME)
        self.historical_data_path = os.path.join(self.data_folder, HISTORICAL_DATA_FILENAME)
        
        self.trained_models = {}
        self.historical_data = pd.DataFrame()
        self.flight_counter = 0
        
        # Try to load existing models and data
        self._load_persistent_data()
    
    def _load_persistent_data(self):
        """Load previously saved models and historical data if they exist."""
        try:
            if os.path.exists(self.model_path) and os.path.exists(self.historical_data_path):
                self.trained_models = joblib.load(self.model_path)
                self.historical_data = pd.read_parquet(self.historical_data_path)
                
                if not self.historical_data.empty:
                    self.flight_counter = int(self.historical_data['flight_id'].max())
                
                print(f"Loaded {len(self.trained_models)} models and {len(self.historical_data)} historical data points.")
            else:
                print("No existing models or historical data found. Starting fresh.")
        except Exception as e:
            print(f"Error loading persistent data: {e}. Starting fresh.")
            self.trained_models = {}
            self.historical_data = pd.DataFrame()
            self.flight_counter = 0
    
    def _save_persistent_data(self):
        """Save models and historical data to disk."""
        try:
            if self.trained_models:
                joblib.dump(self.trained_models, self.model_path)
                print(f"Models saved to {self.model_path}")
            
            if not self.historical_data.empty:
                self.historical_data.to_parquet(self.historical_data_path)
                print(f"Historical data saved to {self.historical_data_path}")
        except Exception as e:
            print(f"Error saving persistent data: {e}")
    
    def _convert_to_native_types(self, obj):
        """
        Convert numpy/pandas types to native Python types for JSON serialization.
        """
        if isinstance(obj, (np.integer, np.int64, np.int32)):
            return int(obj)
        elif isinstance(obj, (np.floating, np.float64, np.float32)):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        elif isinstance(obj, dict):
            return {key: self._convert_to_native_types(value) for key, value in obj.items()}
        elif isinstance(obj, list):
            return [self._convert_to_native_types(item) for item in obj]
        elif pd.isna(obj):
            return None
        else:
            return obj
    
    def process_flight_data(self, df, flight_id):
        """
        Process a single flight's data and segment it into phases.
        
        Args:
            df (pd.DataFrame): Raw flight data
            flight_id (int): Unique flight identifier
            
        Returns:
            pd.DataFrame: Processed data with 'phase' and 'flight_id' columns
        """
        df_copy = df.copy()
        
        # Handle '_Time' vs '_time' column name
        if '_time' not in df_copy.columns and '_Time' in df_copy.columns:
            df_copy.rename(columns={'_Time': '_time'}, inplace=True)
        elif '_time' not in df_copy.columns and '_Time' not in df_copy.columns:
            # Assign dummy time if missing
            df_copy['_time'] = range(len(df_copy))
        
        # Ensure '_time' is numeric
        df_copy['_time'] = pd.to_numeric(df_copy['_time'], errors='coerce')
        
        df_copy['phase'] = 'unknown'
        
        # Check if 'iWOW' column exists
        if 'iWOW' not in df_copy.columns:
            print(f"Warning: 'iWOW' column not found. Assigning all to 'before takeoff'.")
            df_copy['phase'] = 'before takeoff'
            df_copy['flight_id'] = flight_id
            return df_copy
        
        # Identify airborne phases
        airborne_indices = df_copy[df_copy['iWOW'] == 0].index
        
        if airborne_indices.empty:
            df_copy['phase'] = 'before takeoff'
        else:
            first_airborne = airborne_indices.min()
            last_airborne = airborne_indices.max()
            
            df_copy.loc[df_copy.index < first_airborne, 'phase'] = 'before takeoff'
            df_copy.loc[df_copy['iWOW'] == 0, 'phase'] = 'when airborne'
            df_copy.loc[(df_copy.index > last_airborne) & (df_copy['iWOW'] == 1), 'phase'] = 'after landing'
        
        df_copy['flight_id'] = flight_id
        return df_copy
    
    def train_models(self, parameters=None):
        """
        Train Isolation Forest models on historical data.
        
        Args:
            parameters (list): Parameters to train models for. Uses PARAMETERS_TO_ANALYZE if None.
        """
        if parameters is None:
            parameters = PARAMETERS_TO_ANALYZE
        
        if self.historical_data.empty:
            print("No historical data available for training.")
            return
        
        phases = ['before takeoff', 'when airborne', 'after landing']
        self.trained_models.clear()
        models_trained = 0
        
        print("\nTraining Isolation Forest models...")
        
        for param in parameters:
            if param not in self.historical_data.columns:
                continue
            
            for phase in phases:
                phase_data = self.historical_data[
                    (self.historical_data['phase'] == phase) & 
                    (self.historical_data[param].notna())
                ]
                
                if len(phase_data) > 10 and phase_data[param].nunique() > 1:
                    X = phase_data[[param]].values
                    model = IsolationForest(
                        n_estimators=N_ESTIMATORS,
                        contamination=ANOMALY_CONTAMINATION_RATE,
                        random_state=42
                    )
                    model.fit(X)
                    self.trained_models[(param, phase)] = model
                    models_trained += 1
                    print(f"  - Trained model for {param} in '{phase}'")
        
        print(f"\nTrained {models_trained} models.")
        self._save_persistent_data()
    
    def detect_anomalies(self, flight_df, parameters=None):
        """
        Detect anomalies in flight data using trained models with optional statistical filtering.
        Creates parameter-specific anomaly tracking to prevent cross-contamination.
        
        Args:
            flight_df (pd.DataFrame): Flight data to analyze
            parameters (list): Parameters to check. Uses PARAMETERS_TO_ANALYZE if None.
            
        Returns:
            tuple: (flight_df with parameter-specific anomaly columns, list of anomaly details)
        """
        if parameters is None:
            parameters = PARAMETERS_TO_ANALYZE
        
        if not self.trained_models:
            print("Warning: No trained models available. Cannot detect anomalies.")
            flight_df['is_anomaly'] = False
            return flight_df, []
        
        anomalies_detected = []
        flight_df['is_anomaly'] = False  # Keep for backward compatibility
        
        # Create parameter-specific anomaly columns to prevent cross-contamination
        for param in parameters:
            if param in flight_df.columns:
                flight_df[f'is_anomaly_{param}'] = False
        
        for param in parameters:
            if param not in flight_df.columns:
                continue
            
            for phase in flight_df['phase'].unique():
                model_key = (param, phase)
                if model_key in self.trained_models:
                    model = self.trained_models[model_key]
                    phase_data = flight_df[
                        (flight_df['phase'] == phase) & 
                        (flight_df[param].notna())
                    ].copy()
                    
                    if not phase_data.empty:
                        X = phase_data[[param]].values
                        predictions = model.predict(X)
                        anomaly_indices = phase_data.index[predictions == -1]
                        
                        # Apply statistical filter if enabled
                        if USE_STATISTICAL_FILTER and not self.historical_data.empty:
                            anomaly_indices = self._apply_statistical_filter(
                                anomaly_indices, flight_df, param, phase
                            )
                        
                        # Mark anomalies ONLY for this specific parameter
                        flight_df.loc[anomaly_indices, f'is_anomaly_{param}'] = True
                        flight_df.loc[anomaly_indices, 'is_anomaly'] = True  # General flag for any anomaly
                        
                        for idx in anomaly_indices:
                            anomalies_detected.append({
                                'flight_id': int(flight_df.loc[idx, 'flight_id']),
                                'parameter': str(param),
                                'phase': str(flight_df.loc[idx, 'phase']),
                                'time': float(flight_df.loc[idx, '_time']),
                                'value': float(flight_df.loc[idx, param])
                            })
        
        return flight_df, anomalies_detected
    
    def _apply_statistical_filter(self, anomaly_indices, flight_df, param, phase):
        """
        Filter anomalies using statistical thresholds based on historical data.
        Only keeps anomalies that are outside mean ± (SIGMA_THRESHOLD * std_dev).
        This eliminates false positives that are within normal statistical range.
        
        Args:
            anomaly_indices: Indices flagged as anomalies by Isolation Forest
            flight_df: Flight dataframe
            param: Parameter name
            phase: Flight phase
            
        Returns:
            Filtered anomaly indices
        """
        if len(anomaly_indices) == 0:
            return anomaly_indices
        
        print(f"\n  Statistical Filter: Checking {len(anomaly_indices)} {param} anomalies in '{phase}' phase...")
        
        # Get historical data for this parameter and phase
        historical_phase_data = self.historical_data[
            (self.historical_data['phase'] == phase) & 
            (self.historical_data[param].notna())
        ]
        
        if len(historical_phase_data) < 10:
            print(f"  ⚠ Skipping filter: Not enough historical data ({len(historical_phase_data)} points, need 10+)")
            return anomaly_indices
        
        # Calculate statistical bounds from historical data
        hist_mean = historical_phase_data[param].mean()
        hist_std = historical_phase_data[param].std()
        
        if hist_std == 0:
            print(f"  ⚠ Skipping filter: No variation in historical data")
            return anomaly_indices
        
        # Define acceptable range
        lower_bound = hist_mean - (SIGMA_THRESHOLD * hist_std)
        upper_bound = hist_mean + (SIGMA_THRESHOLD * hist_std)
        
        print(f"  Historical stats: mean={hist_mean:.2f}, std={hist_std:.2f}")
        print(f"  Acceptance range: [{lower_bound:.2f}, {upper_bound:.2f}] ({SIGMA_THRESHOLD}σ)")
        
        # Filter: only keep anomalies that are truly outside the statistical range
        filtered_indices = []
        filtered_count = 0
        
        for idx in anomaly_indices:
            value = flight_df.loc[idx, param]
            time = flight_df.loc[idx, '_time']
            
            if value < lower_bound or value > upper_bound:
                filtered_indices.append(idx)
            else:
                # This was a false positive - value is within normal range
                filtered_count += 1
                print(f"  ❌ Filtered: {param}={value:.2f} at t={time:.0f} (within normal range)")
        
        if filtered_count > 0:
            print(f"  ✅ Filtered {filtered_count} false positives, kept {len(filtered_indices)} real anomalies")
        else:
            print(f"  ✅ All {len(filtered_indices)} anomalies are real (none filtered)")
        
        return pd.Index(filtered_indices)
    
    def add_to_training_data(self, flight_df):
        """
        Add flight data to historical training database.
        
        Args:
            flight_df (pd.DataFrame): Processed flight data to add
        """
        self.historical_data = pd.concat([self.historical_data, flight_df], ignore_index=True)
        print(f"Flight added to historical data. Total flights: {self.historical_data['flight_id'].nunique()}")
        self._save_persistent_data()
    
    def analyze_flight(self, excel_path, sheet_name='Clean Data', add_to_training=False):
        """
        Complete flight analysis workflow.
        
        Args:
            excel_path (str): Path to Excel file with flight data
            sheet_name (str): Sheet name to read
            add_to_training (bool): Whether to add this flight to training data
            
        Returns:
            dict: Analysis results including anomalies and statistics
        """
        try:
            # Read flight data
            df_raw = pd.read_excel(excel_path, sheet_name=sheet_name)
            
            if df_raw.empty:
                return {'error': f"Sheet '{sheet_name}' is empty"}
            
            # Process flight
            self.flight_counter += 1
            processed_df = self.process_flight_data(df_raw, self.flight_counter)
            
            # Detect anomalies - use the actual anomaly list from detection
            analyzed_df, anomalies = self.detect_anomalies(processed_df.copy())
            
            # Prepare visualization data
            viz_data = self._prepare_visualization_data(analyzed_df)
            
            # Use the anomalies list directly from detect_anomalies
            # This ensures accuracy - only parameters that were actually flagged are included
            
            # Add to training if requested
            if add_to_training:
                self.add_to_training_data(processed_df)
                self.train_models()  # Retrain with new data
            
            # Prepare results with explicit type conversion
            results = {
                'flight_id': int(self.flight_counter),
                'total_data_points': int(len(analyzed_df)),
                'anomaly_count': int(len(anomalies)),
                'anomaly_percentage': float(round((len(anomalies) / len(analyzed_df)) * 100, 2)) if len(analyzed_df) > 0 else 0.0,
                'anomalies': anomalies,
                'visualization_data': viz_data,
                'phases_summary': self._get_phases_summary(analyzed_df, anomalies),
                'added_to_training': bool(add_to_training),
                'total_historical_flights': int(self.historical_data['flight_id'].nunique()) if not self.historical_data.empty else 0
            }
            
            # Convert entire results dict to native types
            results = self._convert_to_native_types(results)
            
            return results
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {'error': str(e)}
    
    def _extract_anomalies_from_dataframe(self, flight_df):
        """
        Extract anomaly list from the analyzed dataframe.
        This ensures the anomaly table matches exactly what's plotted in the charts.
        
        Args:
            flight_df (pd.DataFrame): Flight data with 'is_anomaly' column
            
        Returns:
            list: List of anomaly dictionaries
        """
        anomalies = []
        
        # Get all rows marked as anomalies
        anomaly_rows = flight_df[flight_df['is_anomaly'] == True]
        
        # For each parameter that has data
        for param in PARAMETERS_TO_ANALYZE:
            if param not in flight_df.columns:
                continue
            
            # Get anomalies for this parameter (where parameter value exists)
            param_anomalies = anomaly_rows[anomaly_rows[param].notna()]
            
            for idx in param_anomalies.index:
                anomalies.append({
                    'flight_id': int(flight_df.loc[idx, 'flight_id']),
                    'parameter': str(param),
                    'phase': str(flight_df.loc[idx, 'phase']),
                    'time': float(flight_df.loc[idx, '_time']),
                    'value': float(flight_df.loc[idx, param])
                })
        
        # Sort by time for better readability
        anomalies.sort(key=lambda x: (x['parameter'], x['time']))
        
        return anomalies
    
    def _prepare_visualization_data(self, flight_df):
        """
        Prepare data for web visualization (organized by parameter and phase).
        Uses parameter-specific 'is_anomaly_{param}' flags to prevent cross-contamination.
        
        Args:
            flight_df (pd.DataFrame): Flight data with parameter-specific anomaly columns
            
        Returns:
            dict: Data organized for Plotly.js charts
        """
        viz_data = {}
        phases = ['before takeoff', 'when airborne', 'after landing']
        
        for param in PARAMETERS_TO_ANALYZE:
            if param not in flight_df.columns:
                continue
            
            viz_data[param] = {}
            
            # Check if parameter-specific anomaly column exists
            param_anomaly_col = f'is_anomaly_{param}'
            has_param_anomaly_col = param_anomaly_col in flight_df.columns
            
            for phase in phases:
                phase_data = flight_df[flight_df['phase'] == phase].copy()
                
                if not phase_data.empty and param in phase_data.columns:
                    # Separate normal and anomaly points using parameter-specific flag
                    if has_param_anomaly_col:
                        # Use parameter-specific anomaly detection (prevents cross-contamination)
                        normal_data = phase_data[phase_data[param_anomaly_col] == False]
                        anomaly_data = phase_data[phase_data[param_anomaly_col] == True]
                    else:
                        # Fallback to general flag (backward compatibility)
                        normal_data = phase_data[phase_data['is_anomaly'] == False]
                        anomaly_data = phase_data[phase_data['is_anomaly'] == True]
                    
                    # Convert to native Python types
                    viz_data[param][phase] = {
                        'normal': {
                            'time': [float(x) for x in normal_data['_time'].tolist()],
                            'values': [float(x) for x in normal_data[param].tolist()]
                        },
                        'anomalies': {
                            'time': [float(x) for x in anomaly_data['_time'].tolist()],
                            'values': [float(x) for x in anomaly_data[param].tolist()]
                        },
                        'has_data': bool(len(phase_data) > 0)
                    }
                else:
                    viz_data[param][phase] = {'has_data': False}
        
        # Add historical data for comparison if available
        if not self.historical_data.empty:
            viz_data['historical'] = self._prepare_historical_data()
        
        return viz_data
    
    def _prepare_historical_data(self):
        """
        Prepare historical data for overlay on charts with intelligent sampling.
        Samples large datasets to prevent browser performance issues.
        """
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
    
    def _get_phases_summary(self, flight_df, anomalies):
        """Get summary statistics for each phase."""
        phases = ['before takeoff', 'when airborne', 'after landing']
        summary = {}
        
        for phase in phases:
            phase_data = flight_df[flight_df['phase'] == phase]
            phase_anomalies = [a for a in anomalies if a['phase'] == phase]
            
            summary[phase] = {
                'total_points': int(len(phase_data)),
                'anomaly_count': int(len(phase_anomalies)),
                'anomaly_percentage': float(round((len(phase_anomalies) / len(phase_data)) * 100, 2)) if len(phase_data) > 0 else 0.0
            }
        
        return summary
