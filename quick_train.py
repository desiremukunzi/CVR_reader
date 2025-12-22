"""
Quick Model Training Script for MI-17 Flight Anomaly Detection
===============================================================

This script allows you to quickly train anomaly detection models by pointing
to a folder containing your historical flight data Excel files.

Usage:
    1. Set the TRAINING_DATA_FOLDER path below to your folder with Excel files
    2. Run: python quick_train.py
    3. Models will be saved and ready for use in the web application

Author: Flight Safety Analysis System
"""

import os
import sys
import pandas as pd
from datetime import datetime
from pathlib import Path

# Add the current directory to the path so we can import flight_analyzer
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flight_analyzer import FlightAnalyzer, PARAMETERS_TO_ANALYZE

# ============================================================================
# CONFIGURATION - MODIFY THESE SETTINGS
# ============================================================================

# Path to folder containing your training Excel files (.xlsx, .xlsm)
TRAINING_DATA_FOLDER = r'A:\Onedrive\RAF-61504\Data'

# Excel sheet name containing clean flight data
SHEET_NAME = 'Clean Data'

# Data folder where models and historical data will be saved
MODEL_DATA_FOLDER = 'flight_data'

# ============================================================================
# DO NOT MODIFY BELOW THIS LINE (unless you know what you're doing!)
# ============================================================================

def print_header():
    """Print a nice header for the training script."""
    print("\n" + "="*70)
    print("  MI-17 FLIGHT ANOMALY DETECTION - QUICK MODEL TRAINING")
    print("="*70)
    print(f"\n📁 Training Data Folder: {TRAINING_DATA_FOLDER}")
    print(f"📊 Sheet Name: {SHEET_NAME}")
    print(f"💾 Models will be saved to: {MODEL_DATA_FOLDER}")
    print("="*70 + "\n")

def validate_folder(folder_path):
    """
    Validate that the training data folder exists and contains Excel files.
    
    Args:
        folder_path (str): Path to the training data folder
        
    Returns:
        list: List of Excel file paths found in the folder
    """
    if not os.path.exists(folder_path):
        print(f"❌ ERROR: Folder does not exist: {folder_path}")
        print("   Please check the path and try again.")
        return None
    
    if not os.path.isdir(folder_path):
        print(f"❌ ERROR: Path is not a directory: {folder_path}")
        return None
    
    # Find all Excel files
    excel_files = []
    for file in os.listdir(folder_path):
        if file.endswith(('.xlsx', '.xlsm', '.xls')):
            excel_files.append(os.path.join(folder_path, file))
    
    if not excel_files:
        print(f"❌ ERROR: No Excel files found in: {folder_path}")
        print("   Please add some .xlsx or .xlsm files and try again.")
        return None
    
    return excel_files

def train_models_from_folder(folder_path, sheet_name='Clean Data', data_folder='flight_data'):
    """
    Train anomaly detection models from all Excel files in a folder.
    
    Args:
        folder_path (str): Path to folder containing Excel training files
        sheet_name (str): Name of the sheet to read from each Excel file
        data_folder (str): Folder to save models and historical data
        
    Returns:
        bool: True if training successful, False otherwise
    """
    print_header()
    
    # Validate folder
    excel_files = validate_folder(folder_path)
    if excel_files is None:
        return False
    
    print(f"✅ Found {len(excel_files)} Excel file(s) for training\n")
    
    # Initialize the flight analyzer
    print("🚀 Initializing Flight Analyzer...")
    analyzer = FlightAnalyzer(data_folder=data_folder)
    
    # Check if models already exist
    if analyzer.trained_models:
        print(f"\n⚠️  WARNING: Existing models found!")
        print(f"   - {len(analyzer.trained_models)} models already trained")
        print(f"   - {analyzer.historical_data['flight_id'].nunique()} flights in database")
        print(f"   - {len(analyzer.historical_data)} total data points")
        
        response = input("\n   Continue and ADD to existing data? (y/n): ").strip().lower()
        if response != 'y':
            print("\n❌ Training cancelled by user.")
            return False
        print()
    
    # Process each Excel file
    successful_flights = 0
    failed_flights = []
    
    print(f"\n{'='*70}")
    print(f"  LOADING FLIGHT DATA")
    print(f"{'='*70}\n")
    
    for idx, excel_file in enumerate(excel_files, 1):
        filename = os.path.basename(excel_file)
        print(f"[{idx}/{len(excel_files)}] Processing: {filename}")
        
        try:
            # Read Excel file
            df_raw = pd.read_excel(excel_file, sheet_name=sheet_name)
            
            if df_raw.empty:
                print(f"   ⚠️  Warning: Sheet '{sheet_name}' is empty - Skipping")
                failed_flights.append((filename, "Empty sheet"))
                continue
            
            # Process flight data
            analyzer.flight_counter += 1
            processed_df = analyzer.process_flight_data(df_raw, analyzer.flight_counter)
            
            # Add to training data (without detecting anomalies)
            analyzer.add_to_training_data(processed_df)
            
            print(f"   ✅ Added Flight {analyzer.flight_counter} - {len(processed_df)} data points")
            successful_flights += 1
            
        except ValueError as ve:
            if "Worksheet named" in str(ve):
                print(f"   ❌ Error: Sheet '{sheet_name}' not found in file")
                failed_flights.append((filename, f"Sheet '{sheet_name}' not found"))
            else:
                print(f"   ❌ Error: {ve}")
                failed_flights.append((filename, str(ve)))
        except Exception as e:
            print(f"   ❌ Error: {e}")
            failed_flights.append((filename, str(e)))
        
        print()
    
    # Train models on accumulated data
    if successful_flights > 0:
        print(f"\n{'='*70}")
        print(f"  TRAINING ANOMALY DETECTION MODELS")
        print(f"{'='*70}\n")
        
        print(f"📊 Training data summary:")
        print(f"   - Total flights: {analyzer.historical_data['flight_id'].nunique()}")
        print(f"   - Total data points: {len(analyzer.historical_data)}")
        print(f"   - Parameters to analyze: {', '.join(PARAMETERS_TO_ANALYZE)}")
        print()
        
        analyzer.train_models(PARAMETERS_TO_ANALYZE)
        
        # Print results
        print(f"\n{'='*70}")
        print(f"  TRAINING RESULTS")
        print(f"{'='*70}\n")
        
        print(f"✅ Successfully processed: {successful_flights} flight(s)")
        
        if failed_flights:
            print(f"❌ Failed: {len(failed_flights)} flight(s)")
            print(f"\n   Failed files:")
            for filename, error in failed_flights:
                print(f"   - {filename}: {error}")
        
        print(f"\n📈 Training Statistics:")
        print(f"   - Models trained: {len(analyzer.trained_models)}")
        print(f"   - Unique flights: {analyzer.historical_data['flight_id'].nunique()}")
        print(f"   - Total data points: {len(analyzer.historical_data)}")
        
        # Show per-phase data counts
        print(f"\n📊 Data distribution by phase:")
        for phase in ['before takeoff', 'when airborne', 'after landing']:
            phase_count = len(analyzer.historical_data[analyzer.historical_data['phase'] == phase])
            percentage = (phase_count / len(analyzer.historical_data)) * 100
            print(f"   - {phase.capitalize()}: {phase_count:,} points ({percentage:.1f}%)")
        
        # Show per-parameter model counts
        print(f"\n🎯 Models trained per parameter:")
        param_counts = {}
        for (param, phase) in analyzer.trained_models.keys():
            param_counts[param] = param_counts.get(param, 0) + 1
        
        for param in PARAMETERS_TO_ANALYZE:
            count = param_counts.get(param, 0)
            status = "✅" if count > 0 else "❌"
            print(f"   {status} {param}: {count} model(s)")
        
        print(f"\n💾 Models saved to: {os.path.abspath(analyzer.model_path)}")
        print(f"💾 Historical data saved to: {os.path.abspath(analyzer.historical_data_path)}")
        
        print(f"\n{'='*70}")
        print(f"  🎉 TRAINING COMPLETE!")
        print(f"{'='*70}\n")
        
        print("✅ Your models are ready for use in the web application!")
        print("   Simply restart Flask and start analyzing flights.\n")
        
        return True
    else:
        print(f"\n{'='*70}")
        print(f"  ❌ TRAINING FAILED")
        print(f"{'='*70}\n")
        print("No flights were successfully processed.")
        print("Please check the errors above and try again.\n")
        return False

def main():
    """Main entry point for the training script."""
    try:
        success = train_models_from_folder(
            folder_path=TRAINING_DATA_FOLDER,
            sheet_name=SHEET_NAME,
            data_folder=MODEL_DATA_FOLDER
        )
        
        if success:
            sys.exit(0)
        else:
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\n\n❌ Training interrupted by user (Ctrl+C)")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
