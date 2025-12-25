import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from sklearn.ensemble import IsolationForest
from matplotlib.backends.backend_pdf import PdfPages
import joblib
import os
import mysql.connector
from mysql.connector import Error, pooling
from datetime import datetime, date
from dotenv import load_dotenv
import sys

# Load environment variables
load_dotenv()

# Global DataFrame to accumulate all flight data
all_flights_data = pd.DataFrame()
flight_counter = 0

# Dictionary to store trained Isolation Forest models
trained_anomaly_models = {}

# Define the parameters for which we want to detect anomalies globally
PARAMETERS_TO_ANALYZE = ['Fcp', 'Xcpl', 'Pedals', 'X_lat', 'X_long', 'PITCH', 'NZ', 'T1', 'T2']

# Define the contamination rate and number of estimators
ANOMALY_CONTAMINATION_RATE = 0.005
N_ESTIMATORS = 300

# Define filenames for saving/loading
MODEL_FILENAME = 'trained_anomaly_models.joblib'
HISTORICAL_DATA_FILENAME = 'all_flights_data.parquet'

# Phase of flight mapping (adjust IDs based on your database)
PHASE_ID_MAP = {
    'before takeoff': 1,
    'when airborne': 2,
    'after landing': 3
}


# ============================================================================
# DATABASE MANAGEMENT
# ============================================================================

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
            print("✓ Database connection pool initialized successfully")
        except Error as e:
            print(f"✗ Error initializing database pool: {e}")
            raise
    
    def get_connection(self):
        """Get connection from pool"""
        try:
            return self.connection_pool.get_connection()
        except Error as e:
            print(f"✗ Error getting connection from pool: {e}")
            raise
    
    def get_or_create_flight(self, flight_date, pic, sic, fe, sortie=1, aircraft_id=3):
        """
        Get flight_id if exists, or create new flight record
        Returns flight_id or None
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
                print(f"✓ Found existing flight ID: {flight_id}")
                return flight_id
            else:
                # Create new flight record
                insert_query = """
                    INSERT INTO flights 
                    (flight_date, aircraft_id, PIC, SIC, FE, flight_type_id, sortie)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                """
                cursor.execute(insert_query, (flight_date, aircraft_id, pic, sic, fe, 1, sortie))
                flight_id = cursor.lastrowid
                connection.commit()
                print(f"✓ Created new flight ID: {flight_id}")
                return flight_id
                
        except Error as e:
            if connection:
                connection.rollback()
            print(f"✗ Database error in get_or_create_flight: {e}")
            return None
        finally:
            if cursor:
                cursor.close()
            if connection:
                connection.close()
    
    def delete_flight_anomalies(self, flight_id):
        """Delete existing anomalies for a flight (for re-analysis scenario)"""
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
                print(f"✓ Deleted {deleted_count} existing anomaly records for flight ID {flight_id}")
            
        except Error as e:
            if connection:
                connection.rollback()
            print(f"✗ Error deleting anomalies: {e}")
        finally:
            if cursor:
                cursor.close()
            if connection:
                connection.close()
    
    def save_anomalies_to_db(self, flight_id, anomalies_summary):
        """
        Save anomaly summary to database
        
        Args:
            flight_id: The flight ID from flights table
            anomalies_summary: Dict with structure {(parameter, phase): count}
        
        Returns:
            bool: True if successful, False otherwise
        """
        if not anomalies_summary:
            print("ℹ No anomalies to save")
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
            
            # Prepare data for batch insert
            anomaly_data = []
            for (param, phase), count in anomalies_summary.items():
                phase_id = PHASE_ID_MAP.get(phase)
                if phase_id is None:
                    print(f"⚠ Warning: Unknown phase '{phase}', skipping")
                    continue
                anomaly_data.append((flight_id, param, phase_id, count))
            
            if anomaly_data:
                cursor.executemany(insert_query, anomaly_data)
                connection.commit()
                print(f"✓ Saved {len(anomaly_data)} anomaly records to database for flight ID {flight_id}")
                return True
            else:
                print("⚠ No valid anomaly data to insert")
                return False
                
        except Error as e:
            if connection:
                connection.rollback()
            print(f"✗ Error saving anomalies to database: {e}")
            return False
        finally:
            if cursor:
                cursor.close()
            if connection:
                connection.close()


# ============================================================================
# FLIGHT DATA PROCESSING FUNCTIONS
# ============================================================================

def process_flight_data(df, flight_id):
    """
    Processes a single flight's data, segments it into distinct phases
    """
    df_copy = df.copy()

    # Handle '_Time' vs '_time' column name
    if '_time' not in df_copy.columns and '_Time' in df_copy.columns:
        df_copy.rename(columns={'_Time': '_time'}, inplace=True)
        print(f"  - Renamed '_Time' to '_time' for Flight {flight_id}.")
    elif '_time' not in df_copy.columns and '_Time' not in df_copy.columns:
        print(f"Error: Neither '_time' nor '_Time' column found in Flight {flight_id} data.")
        df_copy['_time'] = range(len(df_copy))

    df_copy['_time'] = pd.to_numeric(df_copy['_time'])
    df_copy['phase'] = 'unknown'

    if 'iWOW' not in df_copy.columns:
        print(f"Warning: 'iWOW' column not found in Flight {flight_id} data. Assigning all to 'before takeoff'.")
        df_copy['phase'] = 'before takeoff'
        df_copy['flight_id'] = flight_id
        return df_copy

    airborne_indices = df_copy[df_copy['iWOW'] == 0].index

    if airborne_indices.empty:
        df_copy['phase'] = 'before takeoff'
        print(f"Warning: Flight {flight_id} did not show an airborne phase. All data assigned to 'before takeoff'.")
    else:
        first_airborne_start_idx = airborne_indices.min()
        last_airborne_end_idx = airborne_indices.max()

        df_copy.loc[df_copy.index < first_airborne_start_idx, 'phase'] = 'before takeoff'
        df_copy.loc[df_copy['iWOW'] == 0, 'phase'] = 'when airborne'
        df_copy.loc[(df_copy.index > last_airborne_end_idx) & (df_copy['iWOW'] == 1), 'phase'] = 'after landing'

    df_copy['flight_id'] = flight_id
    return df_copy


def train_anomaly_models(data_df, parameters):
    """
    Trains Isolation Forest models for each parameter and phase
    """
    global trained_anomaly_models
    phases = ['before takeoff', 'when airborne', 'after landing']
    
    print("\n" + "="*60)
    print("TRAINING ISOLATION FOREST MODELS")
    print("="*60)
    
    trained_anomaly_models.clear()
    models_trained_count = 0
    
    for param in parameters:
        if param not in data_df.columns:
            print(f"  ⚠ Parameter '{param}' not found in data. Skipping.")
            continue
            
        for phase in phases:
            phase_data = data_df[(data_df['phase'] == phase) & (data_df[param].notna())]
            
            if len(phase_data) > 10 and phase_data[param].nunique() > 1:
                X = phase_data[[param]].values
                model = IsolationForest(
                    n_estimators=N_ESTIMATORS, 
                    contamination=ANOMALY_CONTAMINATION_RATE, 
                    random_state=42
                )
                model.fit(X)
                trained_anomaly_models[(param, phase)] = model
                print(f"  ✓ Trained model for {param} in '{phase}' phase ({len(phase_data)} points)")
                models_trained_count += 1
            else:
                print(f"  ✗ Insufficient data for {param} in '{phase}' phase ({len(phase_data)} points)")
    
    if models_trained_count == 0:
        print("\n⚠ WARNING: No models were trained!")
    else:
        print(f"\n✓ Successfully trained {models_trained_count} models")
        try:
            joblib.dump(trained_anomaly_models, MODEL_FILENAME)
            print(f"✓ Models saved to '{MODEL_FILENAME}'")
        except Exception as e:
            print(f"✗ Error saving models: {e}")
        
        try:
            data_df.to_parquet(HISTORICAL_DATA_FILENAME)
            print(f"✓ Historical data saved to '{HISTORICAL_DATA_FILENAME}'")
        except Exception as e:
            print(f"✗ Error saving historical data: {e}")


def detect_anomalies(flight_df_to_analyze, parameters):
    """
    Detects anomalies using pre-trained models
    """
    anomalies_detected = []
    flight_df_to_analyze['is_anomaly'] = False

    for param in parameters:
        if param not in flight_df_to_analyze.columns:
            print(f"  ⚠ Parameter '{param}' not found in flight data. Skipping.")
            continue

        for phase in flight_df_to_analyze['phase'].unique():
            model_key = (param, phase)
            if model_key in trained_anomaly_models:
                model = trained_anomaly_models[model_key]
                phase_data_to_analyze = flight_df_to_analyze[
                    (flight_df_to_analyze['phase'] == phase) & 
                    (flight_df_to_analyze[param].notna())
                ].copy()
                
                if not phase_data_to_analyze.empty:
                    X_to_analyze = phase_data_to_analyze[[param]].values
                    predictions = model.predict(X_to_analyze)
                    anomaly_indices = phase_data_to_analyze.index[predictions == -1]
                    flight_df_to_analyze.loc[anomaly_indices, 'is_anomaly'] = True
                    
                    for idx in anomaly_indices:
                        anomalies_detected.append({
                            'flight_id': flight_df_to_analyze.loc[idx, 'flight_id'],
                            'parameter': param,
                            'phase': phase,
                            'time': flight_df_to_analyze.loc[idx, '_time'],
                            'value': flight_df_to_analyze.loc[idx, param]
                        })
    
    return flight_df_to_analyze, anomalies_detected


def plot_parameters(historical_df, current_flight_df=None, parameters=None):
    """
    Generates plots for parameters
    """
    if parameters is None:
        parameters = PARAMETERS_TO_ANALYZE

    phases = ['before takeoff', 'when airborne', 'after landing']
    pdf_filename_suffix = f"flight_{current_flight_df['flight_id'].iloc[0]}" if current_flight_df is not None and not current_flight_df.empty else "historical_training_data"
    pdf_filename = f"flight_analysis_{pdf_filename_suffix}.pdf"

    print(f"\n📊 Generating plots and saving to '{pdf_filename}'...")
    
    with PdfPages(pdf_filename) as pdf:
        for param in parameters:
            param_in_historical = param in historical_df.columns
            param_in_current_flight = current_flight_df is not None and param in current_flight_df.columns

            if not param_in_historical and not param_in_current_flight:
                print(f"  ⚠ Parameter '{param}' not found in any data. Skipping.")
                continue
                
            for phase in phases:
                fig, ax = plt.subplots(figsize=(12, 6))
                ax.set_title(f'{param} - {phase.replace("_", " ").title()}')
                ax.set_xlabel('Time')
                ax.set_ylabel(param)
                ax.grid(True, linestyle='--', alpha=0.7)

                # Plot historical data
                if param_in_historical:
                    phase_historical_data = historical_df[historical_df['phase'] == phase]
                    if not phase_historical_data.empty and param in phase_historical_data.columns:
                        ax.scatter(
                            phase_historical_data['_time'], 
                            phase_historical_data[param],
                            label='Historical Data',
                            alpha=0.6, s=10, color='darkgrey'
                        )

                # Plot current flight data
                if current_flight_df is not None and param_in_current_flight:
                    phase_current_flight_data = current_flight_df[current_flight_df['phase'] == phase]
                    if not phase_current_flight_data.empty:
                        # Normal points
                        normal_points = phase_current_flight_data[phase_current_flight_data['is_anomaly'] == False]
                        if not normal_points.empty:
                            ax.scatter(
                                normal_points['_time'], 
                                normal_points[param],
                                label=f'Current Flight (ID: {current_flight_df["flight_id"].iloc[0]})',
                                color='blue', marker='o', s=10, zorder=1
                            )
                        
                        # Anomaly points
                        anomaly_points = phase_current_flight_data[phase_current_flight_data['is_anomaly'] == True]
                        if not anomaly_points.empty:
                            ax.scatter(
                                anomaly_points['_time'], 
                                anomaly_points[param],
                                label=f'Current Flight Anomaly (ID: {current_flight_df["flight_id"].iloc[0]})',
                                color='red', marker='o', s=5, zorder=1, linewidth=1
                            )

                ax.legend(loc='center left', bbox_to_anchor=(1, 0.5), borderaxespad=0.)
                plt.tight_layout(rect=[0, 0, 0.9, 1])
                pdf.savefig(fig)
                plt.close(fig)

    print(f"✓ All plots saved to '{pdf_filename}'")


def load_historical_flights_from_folder(folder_path, sheet_name='Clean Data'):
    """
    Loads all historical flight data from a folder
    """
    global all_flights_data, flight_counter
    print(f"\n" + "="*60)
    print(f"LOADING HISTORICAL FLIGHT DATA")
    print(f"Folder: {folder_path}")
    print("="*60)
    
    all_flights_data = pd.DataFrame()
    flight_counter = 0

    if not os.path.isdir(folder_path):
        print(f"✗ Error: Folder '{folder_path}' not found")
        return

    files_in_folder = os.listdir(folder_path)
    excel_files = [f for f in files_in_folder if f.endswith(('.xlsx', '.xlsm'))]

    if not excel_files:
        print(f"✗ No Excel files found in '{folder_path}'")
        return

    for filename in excel_files:
        filepath = os.path.join(folder_path, filename)
        try:
            print(f"  Loading: {filename}")
            df_raw = pd.read_excel(filepath, sheet_name=sheet_name)
            
            if df_raw.empty:
                print(f"    ⚠ Empty sheet, skipping")
                continue

            flight_counter += 1
            processed_df = process_flight_data(df_raw, flight_counter)
            all_flights_data = pd.concat([all_flights_data, processed_df], ignore_index=True)
            print(f"    ✓ Added Flight {flight_counter}")
        except Exception as e:
            print(f"    ✗ Error: {e}")
    
    print(f"\n✓ Loaded {all_flights_data['flight_id'].nunique()} historical flights")
    print(f"✓ Total data points: {all_flights_data.shape[0]:,}")

    if not all_flights_data.empty:
        train_anomaly_models(all_flights_data, PARAMETERS_TO_ANALYZE)
    else:
        print("✗ No historical data loaded")


def analyze_current_flight(filepath, sheet_name='Clean Data', flight_metadata=None):
    """
    Analyzes a current flight and optionally saves to database
    
    Args:
        filepath: Path to the Excel file
        sheet_name: Sheet name to read
        flight_metadata: Dict with keys: flight_date, pic, sic, fe, sortie, aircraft_id
    
    Returns:
        tuple: (flight_df_with_anomalies, detected_anomalies, anomalies_summary)
    """
    global all_flights_data, flight_counter
    
    print(f"\n" + "="*60)
    print(f"ANALYZING CURRENT FLIGHT")
    print(f"File: {os.path.basename(filepath)}")
    print("="*60)
    
    if not trained_anomaly_models:
        print("✗ Error: No trained models found. Run load_historical_flights_from_folder first.")
        return None, None, None

    try:
        current_df_raw = pd.read_excel(filepath, sheet_name=sheet_name)
        
        if current_df_raw.empty:
            print(f"✗ Error: Empty sheet")
            return None, None, None

        flight_counter += 1
        processed_current_df = process_flight_data(current_df_raw, flight_counter)
        print(f"✓ Flight {flight_counter} loaded and processed")

        # Detect anomalies
        current_flight_with_anomalies, detected_anomalies = detect_anomalies(
            processed_current_df.copy(), 
            PARAMETERS_TO_ANALYZE
        )

        # Generate plots
        plot_parameters(
            all_flights_data.copy(), 
            current_flight_df=current_flight_with_anomalies, 
            parameters=PARAMETERS_TO_ANALYZE
        )

        # Summarize anomalies by parameter and phase
        anomalies_summary = {}
        for anomaly in detected_anomalies:
            key = (anomaly['parameter'], anomaly['phase'])
            anomalies_summary[key] = anomalies_summary.get(key, 0) + 1

        # Display summary
        print(f"\n" + "="*60)
        print(f"ANOMALY DETECTION SUMMARY - FLIGHT {flight_counter}")
        print("="*60)
        
        if detected_anomalies:
            print(f"✗ Detected {len(detected_anomalies)} anomalies:")
            print(f"\n{'Parameter':<15} {'Phase':<18} {'Count':<10}")
            print("-" * 45)
            for (param, phase), count in sorted(anomalies_summary.items()):
                print(f"{param:<15} {phase:<18} {count:<10}")
        else:
            print("✓ No anomalies detected in this flight")
        
        # Add to historical data
        all_flights_data = pd.concat([all_flights_data, processed_current_df], ignore_index=True)
        print(f"\n✓ Flight {flight_counter} added to historical data")
        print(f"✓ Total flights: {all_flights_data['flight_id'].nunique()}")
        print(f"✓ Total data points: {all_flights_data.shape[0]:,}")

        # Ask user about next steps
        print(f"\n" + "="*60)
        print("POST-ANALYSIS OPTIONS")
        print("="*60)
        
        # Option 1: Add to training data
        response = input("\n🔄 Add this flight to training data? (y/n): ").strip().lower()
        if response == 'y':
            train_anomaly_models(all_flights_data, PARAMETERS_TO_ANALYZE)
            print("✓ Models retrained with new flight data")
        else:
            print("ℹ Flight added to accumulated data but models not retrained")
        
        # Option 2: Save to database
        if flight_metadata:
            response = input("\n💾 Save anomaly results to database? (y/n): ").strip().lower()
            if response == 'y':
                save_flight_anomalies_to_database(
                    flight_metadata=flight_metadata,
                    anomalies_summary=anomalies_summary
                )
        else:
            print("\nℹ Flight metadata not provided. Cannot save to database.")
            print("  To enable database saving, provide flight_metadata dict with:")
            print("  - flight_date (date object)")
            print("  - pic (str, 3-char code)")
            print("  - sic (str, 3-char code)")
            print("  - fe (str, 3-char code)")
            print("  - sortie (int, default=1)")
            print("  - aircraft_id (int, default=3)")

        print(f"\n" + "="*60)
        print(f"ANALYSIS COMPLETE - FLIGHT {flight_counter}")
        print("="*60 + "\n")
        
        return current_flight_with_anomalies, detected_anomalies, anomalies_summary

    except FileNotFoundError:
        print(f"✗ Error: File '{filepath}' not found")
        return None, None, None
    except Exception as e:
        print(f"✗ Error processing flight: {e}")
        return None, None, None


def save_flight_anomalies_to_database(flight_metadata, anomalies_summary):
    """
    Save flight and its anomalies to database
    
    Args:
        flight_metadata: Dict with flight information
        anomalies_summary: Dict with {(parameter, phase): count}
    """
    print("\n" + "="*60)
    print("SAVING TO DATABASE")
    print("="*60)
    
    try:
        db_manager = DatabaseManager()
        
        # Extract metadata
        flight_date = flight_metadata.get('flight_date')
        pic = flight_metadata.get('pic')
        sic = flight_metadata.get('sic')
        fe = flight_metadata.get('fe')
        sortie = flight_metadata.get('sortie', 1)
        aircraft_id = flight_metadata.get('aircraft_id', 3)
        
        # Validate required fields
        if not all([flight_date, pic, sic, fe]):
            print("✗ Error: Missing required flight metadata")
            return False
        
        print(f"Flight Date: {flight_date}")
        print(f"Crew: PIC={pic}, SIC={sic}, FE={fe}")
        print(f"Sortie: {sortie}, Aircraft ID: {aircraft_id}")
        
        # Get or create flight record
        flight_id = db_manager.get_or_create_flight(
            flight_date=flight_date,
            pic=pic,
            sic=sic,
            fe=fe,
            sortie=sortie,
            aircraft_id=aircraft_id
        )
        
        if not flight_id:
            print("✗ Error: Could not get/create flight record")
            return False
        
        # Delete existing anomalies for this flight (if any)
        db_manager.delete_flight_anomalies(flight_id)
        
        # Save anomalies
        success = db_manager.save_anomalies_to_db(flight_id, anomalies_summary)
        
        if success:
            print(f"\n✓ Successfully saved anomalies to database for flight ID {flight_id}")
            return True
        else:
            print("\n✗ Failed to save anomalies to database")
            return False
            
    except Exception as e:
        print(f"\n✗ Database error: {e}")
        return False


# ============================================================================
# MAIN EXECUTION
# ============================================================================

if __name__ == "__main__":
    print("\n" + "="*80)
    print(" " * 20 + "FLIGHT ANOMALY DETECTION SYSTEM")
    print("="*80 + "\n")
    
    # Try to load previously saved data
    try:
        print("Checking for saved models and historical data...")
        if os.path.exists(HISTORICAL_DATA_FILENAME) and os.path.exists(MODEL_FILENAME):
            all_flights_data = pd.read_parquet(HISTORICAL_DATA_FILENAME)
            trained_anomaly_models = joblib.load(MODEL_FILENAME)
            if not all_flights_data.empty:
                flight_counter = all_flights_data['flight_id'].max()
            print(f"✓ Loaded saved data: {all_flights_data['flight_id'].nunique()} flights, {all_flights_data.shape[0]:,} points\n")
        else:
            print("No saved data found. Loading from folder...\n")
            historical_data_folder = r'A:\Onedrive\RAF-61504\JUNE\FLIGHTS\FOR_REPORT'
            load_historical_flights_from_folder(historical_data_folder, sheet_name='Clean Data')
    except Exception as e:
        print(f"Error loading data: {e}\n")
        historical_data_folder = r'A:\Onedrive\RAF-61504\JUNE\FLIGHTS\FOR_REPORT'
        load_historical_flights_from_folder(historical_data_folder, sheet_name='Clean Data')

    # Analyze current flight with metadata
    print("\n" + "="*80)
    print("CURRENT FLIGHT ANALYSIS")
    print("="*80 + "\n")
    
    # Example with metadata for database saving
    flight_metadata = {
        'flight_date': date(2025, 7, 4),  # Adjust as needed
        'pic': 'ABC',  # Replace with actual crew codes
        'sic': 'DEF',
        'fe': 'GHI',
        'sortie': 1,
        'aircraft_id': 3
    }
    
    analyze_current_flight(
        filepath=r'A:\Onedrive\RAF-61504\JULY\UNO-561P_04-07-25_1.xlsm',
        sheet_name='Clean Data',
        flight_metadata=flight_metadata
    )