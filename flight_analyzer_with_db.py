import pandas as pd
import numpy as np
from sklearn.ensemble import IsolationForest
import joblib
import os
from datetime import datetime, date
import mysql.connector
from mysql.connector import Error, pooling
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

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

# Phase of flight mapping for database
PHASE_ID_MAP = {
    'before takeoff': 1,
    'when airborne': 2,
    'after landing': 3
}


class DatabaseManager:
    """Manages database connections and operations for anomaly data"""
    
    def __init__(self):
        self.connection_pool = None
        self.initialize_pool()
    
    def initialize_pool(self):
        """Initialize connection pool"""
        try:
            self.connection_pool = pooling.MySQLConnectionPool(
                pool_name="anomaly_pool",
                pool_size=5,
                pool_reset_session=True,
                host=os.getenv('DB_HOST', 'localhost'),
                port=int(os.getenv('DB_PORT', 3306)),
                database=os.getenv('DB_NAME', 'fdap'),
                user=os.getenv('DB_USER'),
                password=os.getenv('DB_PASSWORD'),
                charset='utf8mb4',
                collation='utf8mb4_general_ci'
            )
            print("✓ Database connection pool initialized")
        except Error as e:
            print(f"✗ Database connection error: {e}")
            raise
    
    def get_connection(self):
        """Get connection from pool"""
        try:
            return self.connection_pool.get_connection()
        except Error as e:
            print(f"✗ Error getting connection: {e}")
            raise
    
    def get_or_create_flight(self, flight_date, pic, sic, fe, sortie=1, aircraft_id=3, compliance_percentage=None, checks_not_complied=None):
        """
        Get flight_id if exists, or create new flight record.
        Updates compliance_percentage and checks_not_complied if provided and flight exists.
        """
        connection = None
        cursor = None
        
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            
            # Check if flight exists
            check_query = """
                SELECT id FROM flights 
                WHERE flight_date = %s AND PIC = %s AND SIC = %s AND FE = %s AND sortie = %s
            """
            cursor.execute(check_query, (flight_date, pic, sic, fe, sortie))
            existing = cursor.fetchone()
            
            if existing:
                flight_id = existing[0]
                print(f"  ✓ Found existing flight ID: {flight_id}")
                
                # Update compliance_percentage and checks_not_complied if provided
                if compliance_percentage is not None or checks_not_complied is not None:
                    update_parts = []
                    update_values = []
                    
                    if compliance_percentage is not None:
                        update_parts.append("compliance_percentage = %s")
                        update_values.append(compliance_percentage)
                    
                    if checks_not_complied is not None:
                        update_parts.append("checks_not_complied = %s")
                        update_values.append(checks_not_complied)
                    
                    if update_parts:
                        update_query = f"""
                            UPDATE flights 
                            SET {', '.join(update_parts)}
                            WHERE id = %s
                        """
                        update_values.append(flight_id)
                        cursor.execute(update_query, tuple(update_values))
                        connection.commit()
                        
                        if compliance_percentage is not None:
                            print(f"  ✓ Updated compliance: {compliance_percentage}%")
                        if checks_not_complied is not None:
                            print(f"  ✓ Updated checks not complied: {checks_not_complied}")
                
                return flight_id
            else:
                # Create new flight record
                insert_query = """
                    INSERT INTO flights 
                    (flight_date, aircraft_id, PIC, SIC, FE, flight_type_id, sortie, compliance_percentage, checks_not_complied)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                """
                
                # DEBUG: Show exactly what we're about to insert
                print(f"  🔍 DEBUG: About to INSERT into flights table:")
                print(f"     flight_date={flight_date}")
                print(f"     aircraft_id={aircraft_id}")
                print(f"     PIC={pic}, SIC={sic}, FE={fe}")
                print(f"     flight_type_id=1")
                print(f"     sortie={sortie}")
                print(f"     compliance_percentage={compliance_percentage} (type: {type(compliance_percentage).__name__})")
                print(f"     checks_not_complied={checks_not_complied} (type: {type(checks_not_complied).__name__})")
                
                cursor.execute(insert_query, (flight_date, aircraft_id, pic, sic, fe, 1, sortie, compliance_percentage, checks_not_complied))
                flight_id = cursor.lastrowid
                connection.commit()
                print(f"  ✓ Created new flight ID: {flight_id}")
                if compliance_percentage is not None:
                    print(f"  ✓ Set compliance: {compliance_percentage}%")
                else:
                    print(f"  ⚠️ compliance_percentage was None, saved as NULL")
                if checks_not_complied is not None:
                    print(f"  ✓ Set checks not complied: {checks_not_complied}")
                else:
                    print(f"  ⚠️ checks_not_complied was None, saved as NULL")
                return flight_id
                
        except Error as e:
            if connection:
                connection.rollback()
            print(f"  ✗ Database error: {e}")
            return None
        finally:
            if cursor:
                cursor.close()
            if connection:
                connection.close()
    
    def delete_flight_anomalies(self, flight_id):
        """Delete existing anomalies for a flight"""
        connection = None
        cursor = None
        
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            
            delete_query = "DELETE FROM anomalies WHERE flight_id = %s"
            cursor.execute(delete_query, (flight_id,))
            deleted_count = cursor.rowcount
            connection.commit()
            
            if deleted_count > 0:
                print(f"  ✓ Deleted {deleted_count} existing anomaly records")
            
        except Error as e:
            if connection:
                connection.rollback()
            print(f"  ✗ Error deleting anomalies: {e}")
        finally:
            if cursor:
                cursor.close()
            if connection:
                connection.close()
    
    def save_anomalies(self, flight_id, anomalies_summary):
        """
        Save anomaly summary to database
        
        Args:
            flight_id: The flight ID from flights table
            anomalies_summary: Dict with structure {(parameter, phase): count}
        """
        if not anomalies_summary:
            print("  ℹ No anomalies to save")
            return True
        
        connection = None
        cursor = None
        
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            
            insert_query = """
                INSERT INTO anomalies 
                (flight_id, parameter_MI_17V_5_name, phase_of_flight_id, total_anomalies)
                VALUES (%s, %s, %s, %s)
            """
            
            anomaly_data = []
            for (param, phase), count in anomalies_summary.items():
                phase_id = PHASE_ID_MAP.get(phase)
                if phase_id is None:
                    print(f"  ⚠ Unknown phase '{phase}', skipping")
                    continue
                anomaly_data.append((flight_id, param, phase_id, count))
            
            if anomaly_data:
                cursor.executemany(insert_query, anomaly_data)
                connection.commit()
                print(f"  ✓ Saved {len(anomaly_data)} anomaly records")
                return True
            
            return False
                
        except Error as e:
            if connection:
                connection.rollback()
            print(f"  ✗ Error saving anomalies: {e}")
            return False
        finally:
            if cursor:
                cursor.close()
            if connection:
                connection.close()
    
    def delete_flight_exceedances(self, flight_id):
        """Delete existing exceedances for a flight"""
        connection = None
        cursor = None
        
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            
            delete_query = "DELETE FROM exceedances WHERE flight_id = %s"
            cursor.execute(delete_query, (flight_id,))
            deleted_count = cursor.rowcount
            connection.commit()
            
            if deleted_count > 0:
                print(f"  ✓ Deleted {deleted_count} existing exceedance records")
            
        except Error as e:
            if connection:
                connection.rollback()
            print(f"  ✗ Error deleting exceedances: {e}")
        finally:
            if cursor:
                cursor.close()
            if connection:
                connection.close()
    
    def save_exceedances(self, flight_id, exceedances_list):
        """
        Save exceedance summary to database
        
        Args:
            flight_id: The flight ID from flights table
            exceedances_list: List of dicts [{'parameter': 'IAS', 'count': 5}, ...]
        """
        if not exceedances_list:
            print("  ℹ No exceedances to save")
            return True
        
        connection = None
        cursor = None
        
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            
            insert_query = """
                INSERT INTO exceedances 
                (flight_id, parameter_MI_17V_5_name, number_of_exceedances)
                VALUES (%s, %s, %s)
            """
            
            exceedance_data = [
                (flight_id, exc['parameter'], exc['count'])
                for exc in exceedances_list
            ]
            
            if exceedance_data:
                cursor.executemany(insert_query, exceedance_data)
                connection.commit()
                print(f"  ✓ Saved {len(exceedance_data)} exceedance records")
                return True
            
            return False
                
        except Error as e:
            if connection:
                connection.rollback()
            print(f"  ✗ Error saving exceedances: {e}")
            return False
        finally:
            if cursor:
                cursor.close()
            if connection:
                connection.close()
    
    def save_missed_checks(self, flight_id, missed_checks_list, checklist_type_id):
        """
        Save missed checklist items to database
        
        Args:
            flight_id: The flight ID from flights table
            missed_checks_list: List of tuples (checklist_item, match_score, excel_row)
                               Only items with status='FAIL' should be passed
            checklist_type_id: ID of the checklist type (1=AC-GPU, 2=DC-GPU, 3=WITHOUT GPU)
        
        Schema: missed_checks(flight_id, checklist_type_id, checklist_item_position)
        where checklist_item_position is the Excel row number (2-180)
        """
        if not missed_checks_list:
            print("  ℹ No missed checks to save")
            return True
        
        connection = None
        cursor = None
        
        try:
            connection = self.get_connection()
            cursor = connection.cursor()
            
            insert_query = """
                INSERT INTO missed_checks 
                (flight_id, checklist_type_id, checklist_item_position)
                VALUES (%s, %s, %s)
            """
            
            missed_check_data = []
            for item, score, excel_row in missed_checks_list:
                # excel_row is the actual Excel row number (2-180)
                missed_check_data.append((flight_id, checklist_type_id, excel_row))
            
            if missed_check_data:
                cursor.executemany(insert_query, missed_check_data)
                connection.commit()
                print(f"  ✓ Saved {len(missed_check_data)} missed check records")
                print(f"    - Checklist type ID: {checklist_type_id}")
                print(f"    - Excel row positions: {[row for _, _, row in missed_check_data[:5]]}{'...' if len(missed_check_data) > 5 else ''}")
                return True
            
            return False
                
        except Error as e:
            if connection:
                connection.rollback()
            print(f"  ✗ Error saving missed checks: {e}")
            return False
        finally:
            if cursor:
                cursor.close()
            if connection:
                connection.close()


class FlightAnalyzer:
    def __init__(self, data_folder='flight_data', enable_database=False):
        """
        Initialize the Flight Analyzer with persistent storage.
        
        Args:
            data_folder (str): Folder to store models and historical data
            enable_database (bool): Enable database integration
        """
        self.data_folder = data_folder
        os.makedirs(self.data_folder, exist_ok=True)
        
        self.model_path = os.path.join(self.data_folder, MODEL_FILENAME)
        self.historical_data_path = os.path.join(self.data_folder, HISTORICAL_DATA_FILENAME)
        
        self.trained_models = {}
        self.historical_data = pd.DataFrame()
        self.flight_counter = 0
        
        # Database integration
        self.enable_database = enable_database
        self.db_manager = None
        if self.enable_database:
            try:
                self.db_manager = DatabaseManager()
                print("✓ Database integration enabled")
            except Exception as e:
                print(f"⚠ Database initialization failed: {e}")
                print("  Continuing without database integration")
                self.enable_database = False
        
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
        
        print("\n" + "="*60)
        print("TRAINING ISOLATION FOREST MODELS")
        print("="*60)
        
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
                    print(f"  ✓ Trained {param} in '{phase}' ({len(phase_data)} points)")
        
        if models_trained == 0:
            print("\n⚠ WARNING: No models trained!")
        else:
            print(f"\n✓ Successfully trained {models_trained} models")
        
        self._save_persistent_data()
    
    def detect_anomalies(self, flight_df, parameters=None):
        """
        Detect anomalies in flight data using trained models.
        
        Args:
            flight_df (pd.DataFrame): Flight data to analyze
            parameters (list): Parameters to check. Uses PARAMETERS_TO_ANALYZE if None.
            
        Returns:
            tuple: (flight_df with anomaly flags, list of detected anomalies)
        """
        if parameters is None:
            parameters = PARAMETERS_TO_ANALYZE
        
        anomalies_detected = []
        
        # Add general anomaly flag
        flight_df['is_anomaly'] = False
        
        # Add parameter-specific anomaly flags
        for param in parameters:
            flight_df[f'is_anomaly_{param}'] = False
        
        for param in parameters:
            if param not in flight_df.columns:
                continue
            
            # Get historical statistics for this parameter
            if not self.historical_data.empty and param in self.historical_data.columns:
                hist_mean = self.historical_data[param].mean()
                hist_std = self.historical_data[param].std()
            else:
                hist_mean = None
                hist_std = None
            
            for phase in flight_df['phase'].unique():
                model_key = (param, phase)
                if model_key not in self.trained_models:
                    continue
                
                model = self.trained_models[model_key]
                phase_data = flight_df[
                    (flight_df['phase'] == phase) & 
                    (flight_df[param].notna())
                ].copy()
                
                if phase_data.empty:
                    continue
                
                X = phase_data[[param]].values
                predictions = model.predict(X)
                
                # Get indices where anomalies detected
                anomaly_indices = phase_data.index[predictions == -1]
                
                # Apply statistical filter if enabled
                if USE_STATISTICAL_FILTER and hist_mean is not None and hist_std is not None:
                    filtered_indices = []
                    for idx in anomaly_indices:
                        value = flight_df.loc[idx, param]
                        z_score = abs((value - hist_mean) / hist_std) if hist_std > 0 else 0
                        if z_score >= SIGMA_THRESHOLD:
                            filtered_indices.append(idx)
                    anomaly_indices = filtered_indices
                
                # Mark anomalies with both general and parameter-specific flags
                flight_df.loc[anomaly_indices, 'is_anomaly'] = True
                flight_df.loc[anomaly_indices, f'is_anomaly_{param}'] = True
                
                # Record anomaly details
                for idx in anomaly_indices:
                    anomalies_detected.append({
                        'flight_id': int(flight_df.loc[idx, 'flight_id']),
                        'parameter': param,
                        'phase': phase,
                        'time': float(flight_df.loc[idx, '_time']),
                        'value': float(flight_df.loc[idx, param])
                    })
        
        return flight_df, anomalies_detected
    
    def load_historical_from_folder(self, folder_path, sheet_name='Clean Data'):
        """
        Load historical flight data from a folder of Excel files.
        
        Args:
            folder_path (str): Path to folder containing Excel files
            sheet_name (str): Name of sheet to read from each file
        """
        print(f"\n{'='*60}")
        print(f"LOADING HISTORICAL FLIGHT DATA")
        print(f"Folder: {folder_path}")
        print("="*60)
        
        if not os.path.isdir(folder_path):
            print(f"✗ Error: Folder '{folder_path}' not found")
            return
        
        excel_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xlsm'))]
        
        if not excel_files:
            print(f"✗ No Excel files found")
            return
        
        for filename in excel_files:
            filepath = os.path.join(folder_path, filename)
            try:
                print(f"  Loading: {filename}")
                df_raw = pd.read_excel(filepath, sheet_name=sheet_name)
                
                if df_raw.empty:
                    print(f"    ⚠ Empty sheet, skipping")
                    continue
                
                self.flight_counter += 1
                processed_df = self.process_flight_data(df_raw, self.flight_counter)
                self.historical_data = pd.concat([self.historical_data, processed_df], ignore_index=True)
                print(f"    ✓ Added Flight {self.flight_counter}")
            except Exception as e:
                print(f"    ✗ Error: {e}")
        
        print(f"\n✓ Loaded {self.historical_data['flight_id'].nunique()} flights")
        print(f"✓ Total data points: {len(self.historical_data):,}")
        
        # Train models on loaded data
        if not self.historical_data.empty:
            self.train_models()
    
    def analyze_flight(self, excel_path, sheet_name='Clean Data', 
                       flight_metadata=None, interactive=True, 
                       auto_add_to_training=None, auto_save_to_db=None):
        """
        Analyze a flight for anomalies.
        
        Args:
            excel_path (str): Path to Excel file with flight data
            sheet_name (str): Sheet name to read
            flight_metadata (dict): Optional metadata for database saving
                Required keys: flight_date, pic, sic, fe
                Optional keys: sortie (default=1), aircraft_id (default=3)
            interactive (bool): If True, prompt user for decisions. If False, use auto_* parameters
            auto_add_to_training (bool): Auto decision for adding to training (used when interactive=False)
            auto_save_to_db (bool): Auto decision for database saving (used when interactive=False)
            
        Returns:
            dict: Analysis results with anomalies, statistics, and visualization data
        """
        try:
            print(f"\n{'='*60}")
            print(f"ANALYZING FLIGHT")
            print(f"File: {os.path.basename(excel_path)}")
            print("="*60)
            
            # Check if models are trained
            if not self.trained_models:
                print("\n✗ Error: No trained models found!")
                print("  Please run load_historical_from_folder() first to train models.")
                return {'error': 'No trained models available'}
            
            # Load flight data
            try:
                current_df_raw = pd.read_excel(excel_path, sheet_name=sheet_name)
            except Exception as e:
                return {'error': f'Failed to read Excel file: {e}'}
            
            if current_df_raw.empty:
                return {'error': 'Excel sheet is empty'}
            
            # Process flight data
            self.flight_counter += 1
            flight_df = self.process_flight_data(current_df_raw, self.flight_counter)
            
            print(f"✓ Flight {self.flight_counter} loaded ({len(flight_df):,} data points)")
            
            # Detect anomalies
            flight_df_with_anomalies, detected_anomalies = self.detect_anomalies(flight_df)
            
            # Summarize anomalies by parameter and phase
            anomalies_summary = {}
            for anomaly in detected_anomalies:
                key = (anomaly['parameter'], anomaly['phase'])
                anomalies_summary[key] = anomalies_summary.get(key, 0) + 1
            
            # Display summary
            print(f"\n{'='*60}")
            print(f"ANOMALY DETECTION SUMMARY - FLIGHT {self.flight_counter}")
            print("="*60)
            
            if detected_anomalies:
                print(f"✗ Detected {len(detected_anomalies)} anomalies:\n")
                print(f"{'Parameter':<15} {'Phase':<18} {'Count':<10}")
                print("-" * 45)
                for (param, phase), count in sorted(anomalies_summary.items()):
                    print(f"{param:<15} {phase:<18} {count:<10}")
            else:
                print("✓ No anomalies detected")
            
            # Prepare results
            results = {
                'flight_id': int(self.flight_counter),
                'total_anomalies': len(detected_anomalies),
                'anomalies': [self._convert_to_native_types(a) for a in detected_anomalies],
                'anomalies_by_param_phase': {
                    f"{k[0]}_{k[1]}": int(v) for k, v in anomalies_summary.items()
                },
                'phases_summary': self._get_phases_summary(flight_df_with_anomalies, detected_anomalies),
                'visualization_data': self._prepare_visualization_data(flight_df_with_anomalies),
                'timestamp': datetime.now().isoformat()
            }
            
            # Add to accumulated data (always)
            self.historical_data = pd.concat(
                [self.historical_data, flight_df], 
                ignore_index=True
            )
            print(f"\n✓ Flight added to accumulated data")
            print(f"  Total flights: {self.historical_data['flight_id'].nunique()}")
            print(f"  Total data points: {len(self.historical_data):,}")
            
            # Post-analysis decisions
            print(f"\n{'='*60}")
            print("POST-ANALYSIS OPTIONS")
            print("="*60)
            
            # Decision 1: Add to training data
            add_to_training = False
            if interactive:
                response = input("\n🔄 Retrain models with this flight? (y/n): ").strip().lower()
                add_to_training = (response == 'y')
            else:
                add_to_training = auto_add_to_training if auto_add_to_training is not None else False
            
            if add_to_training:
                print("\n  Retraining models with new flight data...")
                self.train_models()
                results['added_to_training'] = True
            else:
                print("\n  ℹ Flight kept in accumulated data but models not retrained")
                results['added_to_training'] = False
            
            # Decision 2: Save to database
            save_to_db = False
            if self.enable_database and flight_metadata:
                if interactive:
                    response = input("\n💾 Save anomaly results to database? (y/n): ").strip().lower()
                    save_to_db = (response == 'y')
                else:
                    save_to_db = auto_save_to_db if auto_save_to_db is not None else False
                
                if save_to_db:
                    success = self._save_to_database(flight_metadata, anomalies_summary)
                    results['saved_to_database'] = success
                else:
                    print("\n  ℹ Results not saved to database")
                    results['saved_to_database'] = False
            elif self.enable_database and not flight_metadata:
                print("\n  ⚠ Flight metadata not provided. Cannot save to database.")
                print("    To enable database saving, provide flight_metadata dict with:")
                print("    - flight_date (date object)")
                print("    - pic (str, 3-char code)")
                print("    - sic (str, 3-char code)")
                print("    - fe (str, 3-char code)")
                results['saved_to_database'] = False
            else:
                results['saved_to_database'] = False
            
            print(f"\n{'='*60}")
            print(f"ANALYSIS COMPLETE - FLIGHT {self.flight_counter}")
            print("="*60 + "\n")
            
            return results
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {'error': str(e)}
    
    def _save_to_database(self, flight_metadata, anomalies_summary, cvr_results=None, exceedances_list=None):
        """
        Save flight, anomalies, exceedances, and missed checks to database.
        
        Args:
            flight_metadata: Dict with flight info (date, crew, sortie, etc.)
            anomalies_summary: Dict with anomaly counts by (parameter, phase)
            cvr_results: Dict with CVR analysis results (optional)
                {
                    'results': [(status, item, score, matched_text, excel_row), ...],
                    'compliance_percent': float,
                    'not_complied_count': int,
                    'checklist_type_id': int,
                    'sheet_name': str
                }
            exceedances_list: List of dicts with exceedance data (optional)
                [{'parameter': 'IAS', 'count': 5}, ...]
        """
        print(f"\n{'='*60}")
        print("SAVING TO DATABASE")
        print("="*60)
        
        try:
            # Extract metadata
            flight_date = flight_metadata.get('flight_date')
            pic = flight_metadata.get('pic')
            sic = flight_metadata.get('sic')
            fe = flight_metadata.get('fe')
            sortie = flight_metadata.get('sortie', 1)
            aircraft_id = flight_metadata.get('aircraft_id', 3)
            
            # NEW: Extract compliance data if available
            compliance_percentage = None
            checks_not_complied = None
            if cvr_results:
                print(f"  🔍 DEBUG: cvr_results keys: {list(cvr_results.keys())}")
                compliance_percentage = cvr_results.get('compliance_percent')
                checks_not_complied = cvr_results.get('not_complied_count')
                print(f"  🔍 DEBUG: Extracted from cvr_results:")
                print(f"     compliance_percent key → {compliance_percentage}")
                print(f"     not_complied_count key → {checks_not_complied}")
            else:
                print(f"  ⚠️ WARNING: cvr_results is None or empty!")
            
            # Validate
            if not all([flight_date, pic, sic, fe]):
                print("  ✗ Error: Missing required flight metadata")
                return False
            
            print(f"  Flight Date: {flight_date}")
            print(f"  Crew: PIC={pic}, SIC={sic}, FE={fe}")
            print(f"  Sortie: {sortie}, Aircraft ID: {aircraft_id}")
            if compliance_percentage is not None:
                print(f"  Compliance: {compliance_percentage}%")
            else:
                print(f"  ⚠️ Compliance percentage is None")
            if checks_not_complied is not None:
                print(f"  Checks Not Complied: {checks_not_complied}")
            else:
                print(f"  ⚠️ Checks not complied is None")
            
            # Get or create flight record (now with compliance data)
            flight_id = self.db_manager.get_or_create_flight(
                flight_date=flight_date,
                pic=pic,
                sic=sic,
                fe=fe,
                sortie=sortie,
                aircraft_id=aircraft_id,
                compliance_percentage=compliance_percentage,  # NEW: Pass compliance data
                checks_not_complied=checks_not_complied  # NEW: Pass checks not complied count
            )
            
            if not flight_id:
                print("  ✗ Error: Could not get/create flight record")
                return False
            
            # Delete existing anomalies, exceedances, and missed checks
            self.db_manager.delete_flight_anomalies(flight_id)
            self.db_manager.delete_flight_exceedances(flight_id)
            
            # Save anomalies
            anomaly_success = self.db_manager.save_anomalies(flight_id, anomalies_summary)
            
            # NEW: Save exceedances if provided
            exceedance_success = True
            if exceedances_list:
                print(f"  💾 Saving {len(exceedances_list)} exceedances...")
                exceedance_success = self.db_manager.save_exceedances(flight_id, exceedances_list)
            
            # NEW: Save missed checks if CVR results provided
            missed_checks_success = True
            if cvr_results and cvr_results.get('results'):
                # Extract checklist_type_id from CVR results
                checklist_type_id = cvr_results.get('checklist_type_id', 1)
                
                # Extract failed checks (status='FAIL')
                # Format: (status, item, score, matched_text, excel_row)
                missed_checks = [
                    (item, score, excel_row) 
                    for status, item, score, matched_text, excel_row in cvr_results['results']
                    if status == 'FAIL'
                ]
                
                if missed_checks:
                    print(f"  💾 Saving {len(missed_checks)} missed checks...")
                    print(f"    - Checklist type: {cvr_results.get('sheet_name', 'Unknown')}")
                    missed_checks_success = self.db_manager.save_missed_checks(
                        flight_id, missed_checks, checklist_type_id
                    )
            
            # Check overall success - at least flight record must be saved
            if flight_id:
                # Count what was saved
                saved_items = []
                if anomaly_success and anomalies_summary:
                    saved_items.append("anomalies")
                if exceedance_success and exceedances_list:
                    saved_items.append("exceedances")
                if missed_checks_success and cvr_results:
                    saved_items.append("missed checks")
                
                if saved_items:
                    print(f"\n  ✓ Successfully saved to database (Flight ID: {flight_id})")
                    print(f"    Saved: {', '.join(saved_items)}")
                else:
                    print(f"\n  ✓ Flight record saved (Flight ID: {flight_id})")
                    print(f"    ℹ️ No anomalies, exceedances, or missed checks to save")
                
                return True
            else:
                print("\n  ✗ Failed to create flight record")
                return False
                
        except Exception as e:
            print(f"\n  ✗ Database error: {e}")
            import traceback
            traceback.print_exc()
            return False
    
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
            
            # FIXED: Match frontend expectations (anomalies_report.html uses 'anomalies' and 'percentage')
            summary[phase] = {
                'total_points': int(len(phase_data)),
                'anomalies': int(len(phase_anomalies)),  # Changed from 'anomaly_count'
                'percentage': float(round((len(phase_anomalies) / len(phase_data)) * 100, 2)) if len(phase_data) > 0 else 0.0  # Changed from 'anomaly_percentage'
            }
        
        return summary


# Example usage
if __name__ == "__main__":
    from datetime import date
    
    # Initialize analyzer with database integration
    analyzer = FlightAnalyzer(data_folder='flight_data', enable_database=True)
    
    # Load historical flights for training
    historical_folder = r'A:\Onedrive\RAF-61504\JUNE\FLIGHTS\FOR_REPORT'
    analyzer.load_historical_from_folder(historical_folder, sheet_name='Clean Data')
    
    # Analyze a new flight with metadata
    flight_metadata = {
        'flight_date': date(2025, 7, 4),
        'pic': 'JDO',
        'sic': 'JSM',
        'fe': 'RBR',
        'sortie': 1,
        'aircraft_id': 3
    }
    
    results = analyzer.analyze_flight(
        excel_path=r'A:\Onedrive\RAF-61504\JULY\UNO-561P_04-07-25_1.xlsm',
        sheet_name='Clean Data',
        flight_metadata=flight_metadata,
        interactive=True  # Set to False for automated processing
    )
    
    print("\nAnalysis Results:")
    print(f"  Total Anomalies: {results.get('total_anomalies', 0)}")
    print(f"  Added to Training: {results.get('added_to_training', False)}")
    print(f"  Saved to Database: {results.get('saved_to_database', False)}")