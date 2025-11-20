import pandas as pd
import numpy as np
from sklearn.ensemble import IsolationForest
import joblib
import os
from datetime import datetime

# Define the parameters for anomaly detection
PARAMETERS_TO_ANALYZE = ['Fcp', 'Xcpl', 'Pedals', 'X_lat', 'X_long', 'PITCH', 'NZ', 'T1', 'T2']

# Anomaly detection configuration
ANOMALY_CONTAMINATION_RATE = 0.0001
N_ESTIMATORS = 300

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
        Detect anomalies in flight data using trained models.
        
        Args:
            flight_df (pd.DataFrame): Flight data to analyze
            parameters (list): Parameters to check. Uses PARAMETERS_TO_ANALYZE if None.
            
        Returns:
            tuple: (flight_df with 'is_anomaly' column, list of anomaly details)
        """
        if parameters is None:
            parameters = PARAMETERS_TO_ANALYZE
        
        if not self.trained_models:
            print("Warning: No trained models available. Cannot detect anomalies.")
            flight_df['is_anomaly'] = False
            return flight_df, []
        
        anomalies_detected = []
        flight_df['is_anomaly'] = False
        
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
                        
                        flight_df.loc[anomaly_indices, 'is_anomaly'] = True
                        
                        for idx in anomaly_indices:
                            anomalies_detected.append({
                                'flight_id': int(flight_df.loc[idx, 'flight_id']),
                                'parameter': str(param),
                                'phase': str(phase),
                                'time': float(flight_df.loc[idx, '_time']),
                                'value': float(flight_df.loc[idx, param])
                            })
        
        return flight_df, anomalies_detected
    
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
            
            # Detect anomalies
            analyzed_df, anomalies = self.detect_anomalies(processed_df.copy())
            
            # Prepare visualization data
            viz_data = self._prepare_visualization_data(analyzed_df, anomalies)
            
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
                'anomalies': anomalies,  # Already converted in detect_anomalies
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
    
    def _prepare_visualization_data(self, flight_df, anomalies):
        """
        Prepare data for web visualization (organized by parameter and phase).
        
        Returns:
            dict: Data organized for Plotly.js charts
        """
        viz_data = {}
        phases = ['before takeoff', 'when airborne', 'after landing']
        
        for param in PARAMETERS_TO_ANALYZE:
            if param not in flight_df.columns:
                continue
            
            viz_data[param] = {}
            
            for phase in phases:
                phase_data = flight_df[flight_df['phase'] == phase].copy()
                
                if not phase_data.empty and param in phase_data.columns:
                    # Separate normal and anomaly points
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
        """Prepare historical data for overlay on charts."""
        historical_viz = {}
        phases = ['before takeoff', 'when airborne', 'after landing']
        
        for param in PARAMETERS_TO_ANALYZE:
            if param not in self.historical_data.columns:
                continue
            
            historical_viz[param] = {}
            
            for phase in phases:
                phase_data = self.historical_data[self.historical_data['phase'] == phase]
                
                if not phase_data.empty and param in phase_data.columns:
                    historical_viz[param][phase] = {
                        'time': [float(x) for x in phase_data['_time'].tolist()],
                        'values': [float(x) for x in phase_data[param].tolist()]
                    }
        
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
