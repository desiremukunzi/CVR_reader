"""
FlightAnalyzer class - Wrapper for load.py functions
This provides a class-based API compatible with the new app.py
"""

import pandas as pd
import os
import sys
from datetime import datetime

# Import functions from your existing load.py
try:
    # Add current directory to path to find load.py
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    
    from load import (
        process_flight_data,
        train_anomaly_models,
        detect_anomalies,
        all_flights_data,
        flight_counter,
        trained_anomaly_models,
        PARAMETERS_TO_ANALYZE
    )
    import load  # Import the module itself to access globals
    LOAD_AVAILABLE = True
except ImportError as e:
    print(f"Warning: Could not import from load.py: {e}")
    LOAD_AVAILABLE = False
    PARAMETERS_TO_ANALYZE = ['Fcp', 'Xcpl', 'Pedals', 'X_lat', 'X_long', 'PITCH', 'NZ', 'T1', 'T2']


class FlightAnalyzer:
    """
    FlightAnalyzer class compatible with both old and new app.py
    Wraps existing load.py functionality
    """
    
    def __init__(self, data_folder='flight_data', enable_database=False):
        """Initialize the FlightAnalyzer"""
        self.data_folder = data_folder
        self.enable_database = enable_database
        
        if not LOAD_AVAILABLE:
            print("⚠ Warning: load.py not available. Creating empty analyzer.")
            self.historical_data = pd.DataFrame()
            self.flight_counter = 0
            self.trained_models = {}
            return
        
        # Reference the global data from load.py
        self.historical_data = load.all_flights_data
        self.flight_counter = load.flight_counter
        self.trained_models = load.trained_anomaly_models
    
    def load_historical_from_folder(self, folder_path, sheet_name='Clean Data'):
        """Load historical flights from a folder"""
        if not LOAD_AVAILABLE:
            print("⚠ Warning: load.py not available")
            return
        
        print(f"Loading historical data from: {folder_path}")
        
        # This function updates the globals in load.py
        from load import load_historical_flights_from_folder
        load_historical_flights_from_folder(folder_path, sheet_name)
        
        # Update our references
        self.historical_data = load.all_flights_data
        self.flight_counter = load.flight_counter
        self.trained_models = load.trained_anomaly_models
    
    def analyze_flight(self, excel_path, sheet_name='Clean Data', 
                      flight_metadata=None, interactive=False,
                      auto_add_to_training=False, auto_save_to_db=False,
                      add_to_training=None):
        """
        Analyze a flight for anomalies
        
        Compatible with both old and new calling conventions
        """
        if not LOAD_AVAILABLE:
            return {'error': 'load.py module not available'}
        
        try:
            print(f"\n📊 Analyzing: {os.path.basename(excel_path)}")
            print(f"   Sheet: {sheet_name}")
            
            # Read the flight data
            df_raw = pd.read_excel(excel_path, sheet_name=sheet_name)
            
            if df_raw.empty:
                return {'error': 'Sheet is empty'}
            
            # Increment flight counter
            load.flight_counter += 1
            current_flight_id = load.flight_counter
            
            # Process the flight data (adds phase column, flight_id)
            processed_df = process_flight_data(df_raw, current_flight_id)
            
            # Detect anomalies in this flight
            flight_with_anomalies, detected_anomalies = detect_anomalies(
                processed_df.copy(), 
                PARAMETERS_TO_ANALYZE
            )
            
            # Add to historical data
            load.all_flights_data = pd.concat(
                [load.all_flights_data, processed_df], 
                ignore_index=True
            )
            
            # Update our reference
            self.historical_data = load.all_flights_data
            self.flight_counter = load.flight_counter
            
            # Prepare results for web app
            results = self._format_results(
                flight_with_anomalies, 
                detected_anomalies, 
                current_flight_id,
                flight_metadata
            )
            
            # If auto_add_to_training is True, retrain models
            if auto_add_to_training or add_to_training:
                self.train_models()
                results['added_to_training'] = True
            else:
                results['added_to_training'] = False
            
            results['saved_to_database'] = False  # Not supported
            
            print(f"✓ Found {results['total_anomalies']} anomalies")
            
            return results
            
        except Exception as e:
            print(f"❌ Error: {e}")
            import traceback
            traceback.print_exc()
            return {'error': str(e)}
    
    def _format_results(self, flight_data, anomalies_list, flight_id, metadata):
        """Format results for the web app"""
        
        # Count total anomalies
        total_anomalies = len(anomalies_list)
        total_points = len(flight_data)
        
        # Group anomalies by parameter and phase
        anomalies_by_param_phase = {}
        for anomaly in anomalies_list:
            param = anomaly['parameter']
            phase = anomaly['phase']
            key = f"{param}_{phase.replace(' ', '_')}"  # Use underscore for JSON
            anomalies_by_param_phase[key] = anomalies_by_param_phase.get(key, 0) + 1
        
        # Calculate phase summaries
        phases_summary = {}
        for phase in ['before takeoff', 'when airborne', 'after landing']:
            phase_data = flight_data[flight_data['phase'] == phase]
            phase_anomalies = [a for a in anomalies_list if a['phase'] == phase]
            
            if not phase_data.empty:
                phase_total = len(phase_data)
                phase_anom_count = len(phase_anomalies)
                percentage = float((phase_anom_count / phase_total * 100) if phase_total > 0 else 0.0)
                
                phases_summary[phase] = {
                    'total_points': phase_total,
                    'anomalies': phase_anom_count,
                    'percentage': percentage
                }
                
                print(f"  Phase '{phase}': {phase_total} points, {phase_anom_count} anomalies ({percentage:.2f}%)")
        
        # Prepare visualization data
        visualization_data = {}
        for param in PARAMETERS_TO_ANALYZE:
            if param in flight_data.columns:
                for phase in ['before takeoff', 'when airborne', 'after landing']:
                    phase_data = flight_data[flight_data['phase'] == phase]
                    
                    if not phase_data.empty and param in phase_data.columns:
                        param_data = phase_data[param].dropna()
                        
                        if not param_data.empty:
                            key = f"{param}_{phase.replace(' ', '_')}"
                            
                            # Get time values
                            if '_time' in phase_data.columns:
                                x_values = phase_data['_time'].tolist()
                            else:
                                x_values = list(range(len(phase_data)))
                            
                            visualization_data[key] = {
                                'x': x_values,
                                'y': param_data.tolist(),
                                'name': f"{param} - {phase}"
                            }
        
        # Build results
        results = {
            'flight_id': flight_id,
            'total_data_points': total_points,
            'total_anomalies': total_anomalies,
            'anomaly_percentage': float((total_anomalies / total_points * 100) if total_points > 0 else 0.0),  # Ensure float
            'anomalies': anomalies_list,
            'anomalies_by_param_phase': anomalies_by_param_phase,
            'phases_summary': phases_summary,
            'visualization_data': visualization_data,
            'flight_metadata': metadata or {},
            'added_to_training': False,
            'saved_to_database': False
        }
        
        return results
    
    def train_models(self):
        """Retrain anomaly detection models"""
        if not LOAD_AVAILABLE:
            print("⚠ Warning: load.py not available")
            return
        
        print("🔄 Retraining models...")
        train_anomaly_models(load.all_flights_data, PARAMETERS_TO_ANALYZE)
        
        # Update reference
        self.trained_models = load.trained_anomaly_models
        print("✓ Models retrained")
    
    def _save_to_database(self, flight_metadata, anomalies_dict):
        """Save to database (not supported in this version)"""
        print("⚠ Database save not supported")
        print("  Use flight_analyzer_with_db.py for database features")
        return False