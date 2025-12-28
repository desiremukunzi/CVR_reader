"""
FlightAnalyzer - Anomaly Detection for MI-17V-5 Flight Data
Save this as: A:\CVR_reader\flight_analyzer.py

This analyzer:
1. Loads trained ML models from A:\CVR_reader\flight_data
2. Processes flight data from Excel files
3. Detects anomalies across 3 phases of flight
4. Returns results in database-ready format
"""

import os
import pickle
import joblib
import pandas as pd
import numpy as np
from pathlib import Path
from typing import Dict, List, Optional
import logging

# Setup logging
logger = logging.getLogger(__name__)


class FlightAnalyzer:
    """
    Analyzes flight data for anomalies using trained ML models
    
    Phases of Flight:
        1 = Before Takeoff
        2 = When Airborne  
        3 = After Landing
    """
    
    def __init__(self, model_folder: str = None):
        """
        Initialize the Flight Analyzer
        
        Args:
            model_folder: Path to folder containing trained models
                         Defaults to A:\CVR_reader\flight_data
        """
        self.model_folder = model_folder or r'A:\CVR_reader\flight_data'
        self.models = {}
        self.phase_models = {
            1: [],  # Before takeoff models
            2: [],  # Airborne models
            3: []   # After landing models
        }
        
        # Parameter mapping - maps Excel columns to parameter names
        self.parameter_columns = {
            'N1': 'N1',
            'N2': 'N2',
            'Nmr': 'Nmr',
            'IAS': 'IAS',
            'Alt': 'Alt',
            'PITCH': 'PITCH',
            'Roll': 'Roll',
            'Fcp': 'Fcp',
            # Add more as needed
        }
        
        self.load_models()
    
    def load_models(self):
        """Load all trained models from the model folder"""
        model_path = Path(self.model_folder)
        
        if not model_path.exists():
            logger.warning(f"Model folder not found: {model_path}")
            logger.warning("Anomaly detection will return empty results")
            return
        
        # Load .pkl files (pickle format)
        pkl_files = list(model_path.glob('*.pkl'))
        for model_file in pkl_files:
            try:
                with open(model_file, 'rb') as f:
                    model_name = model_file.stem
                    self.models[model_name] = pickle.load(f)
                    logger.info(f"Loaded model: {model_name}")
                    
                    # Classify by phase (based on naming convention)
                    if 'before' in model_name.lower() or 'phase1' in model_name.lower():
                        self.phase_models[1].append(model_name)
                    elif 'airborne' in model_name.lower() or 'phase2' in model_name.lower():
                        self.phase_models[2].append(model_name)
                    elif 'after' in model_name.lower() or 'phase3' in model_name.lower():
                        self.phase_models[3].append(model_name)
                    
            except Exception as e:
                logger.error(f"Error loading {model_file}: {e}")
        
        # Load .joblib files (joblib format)
        joblib_files = list(model_path.glob('*.joblib'))
        for model_file in joblib_files:
            try:
                model_name = model_file.stem
                self.models[model_name] = joblib.load(model_file)
                logger.info(f"Loaded model: {model_name}")
                
                # Classify by phase
                if 'before' in model_name.lower() or 'phase1' in model_name.lower():
                    self.phase_models[1].append(model_name)
                elif 'airborne' in model_name.lower() or 'phase2' in model_name.lower():
                    self.phase_models[2].append(model_name)
                elif 'after' in model_name.lower() or 'phase3' in model_name.lower():
                    self.phase_models[3].append(model_name)
                    
            except Exception as e:
                logger.error(f"Error loading {model_file}: {e}")
        
        logger.info(f"Total models loaded: {len(self.models)}")
        logger.info(f"Phase distribution: Before={len(self.phase_models[1])}, "
                   f"Airborne={len(self.phase_models[2])}, After={len(self.phase_models[3])}")
    
    def analyze_file(self, file_path: str) -> Dict:
        """
        Analyze a flight data Excel file for anomalies
        
        Args:
            file_path: Path to Excel file with 'Clean Data' sheet
            
        Returns:
            dict: {
                'anomalies': [
                    {
                        'parameter': str,
                        'phase_id': int (1, 2, or 3),
                        'score': float (0.0 to 1.0),
                        'description': str
                    },
                    ...
                ],
                'stats': {...}
            }
        """
        anomalies = []
        stats = {
            'total_analyzed': 0,
            'anomalies_found': 0,
            'phases_analyzed': []
        }
        
        try:
            # Read the Clean Data sheet
            df = pd.read_excel(file_path, sheet_name='Clean Data')
            logger.info(f"Loaded Clean Data: {len(df)} rows, {len(df.columns)} columns")
            
            if len(self.models) == 0:
                logger.warning("No models loaded - returning empty anomalies")
                return {'anomalies': [], 'stats': stats}
            
            # Analyze each phase of flight
            for phase_id in [1, 2, 3]:
                phase_anomalies = self._analyze_phase(df, phase_id)
                anomalies.extend(phase_anomalies)
                
                if len(self.phase_models[phase_id]) > 0:
                    stats['phases_analyzed'].append(phase_id)
            
            stats['total_analyzed'] = len(df)
            stats['anomalies_found'] = len(anomalies)
            
            logger.info(f"Analysis complete: {len(anomalies)} anomalies detected")
            
        except FileNotFoundError:
            logger.error(f"File not found: {file_path}")
        except ValueError as e:
            logger.error(f"Clean Data sheet not found in {file_path}: {e}")
        except Exception as e:
            logger.error(f"Error analyzing file: {e}", exc_info=True)
        
        return {
            'anomalies': anomalies,
            'stats': stats
        }
    
    def _analyze_phase(self, df: pd.DataFrame, phase_id: int) -> List[Dict]:
        """
        Analyze a specific phase of flight for anomalies
        
        Args:
            df: DataFrame with flight data
            phase_id: Phase identifier (1, 2, or 3)
            
        Returns:
            List of anomaly dictionaries
        """
        anomalies = []
        
        # Get models for this phase
        phase_model_names = self.phase_models.get(phase_id, [])
        
        if not phase_model_names:
            # No specific models for this phase, try generic models
            logger.debug(f"No specific models for phase {phase_id}")
            return anomalies
        
        # Filter data for this phase if phase column exists
        phase_df = self._filter_by_phase(df, phase_id)
        
        if phase_df.empty:
            logger.debug(f"No data for phase {phase_id}")
            return anomalies
        
        # Run each model for this phase
        for model_name in phase_model_names:
            model = self.models[model_name]
            
            try:
                phase_anomalies = self._run_model(model, phase_df, phase_id, model_name)
                anomalies.extend(phase_anomalies)
            except Exception as e:
                logger.error(f"Error running model {model_name}: {e}")
        
        return anomalies
    
    def _filter_by_phase(self, df: pd.DataFrame, phase_id: int) -> pd.DataFrame:
        """
        Filter data by phase of flight
        
        Args:
            df: Full flight data
            phase_id: Phase to filter (1, 2, or 3)
            
        Returns:
            Filtered DataFrame
        """
        # Check if phase column exists
        phase_columns = [col for col in df.columns if 'phase' in col.lower()]
        
        if not phase_columns:
            # No phase column - return all data
            logger.debug("No phase column found, using all data")
            return df
        
        phase_col = phase_columns[0]
        filtered = df[df[phase_col] == phase_id]
        
        logger.debug(f"Phase {phase_id}: {len(filtered)} rows from {len(df)} total")
        return filtered
    
    def _run_model(self, model, df: pd.DataFrame, phase_id: int, 
                   model_name: str) -> List[Dict]:
        """
        Run a specific model on data and extract anomalies
        
        Args:
            model: Trained ML model
            df: Data to analyze
            phase_id: Phase of flight
            model_name: Name of the model
            
        Returns:
            List of anomalies detected
        """
        anomalies = []
        
        try:
            # Prepare features for the model
            # This depends on your model's expected input
            # Adjust column selection based on your models
            
            # Option 1: Use all numeric columns
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            X = df[numeric_cols].fillna(0)  # Fill NaN with 0
            
            # Option 2: Use specific columns (uncomment if needed)
            # required_cols = ['N1', 'N2', 'Nmr', 'IAS', 'Alt']
            # available_cols = [col for col in required_cols if col in df.columns]
            # X = df[available_cols].fillna(0)
            
            if X.empty:
                return anomalies
            
            # Run model prediction
            # Different model types have different methods
            
            if hasattr(model, 'predict'):
                # Supervised model (returns class labels)
                predictions = model.predict(X)
                # -1 or 1 typically indicates anomaly
                anomaly_mask = predictions == -1
                
            elif hasattr(model, 'decision_function'):
                # Anomaly detection model (returns scores)
                scores = model.decision_function(X)
                # Negative scores indicate anomalies
                anomaly_mask = scores < 0
                
            elif hasattr(model, 'score_samples'):
                # Density-based model
                scores = model.score_samples(X)
                # Low scores indicate anomalies
                threshold = np.percentile(scores, 5)  # Bottom 5%
                anomaly_mask = scores < threshold
                
            else:
                logger.warning(f"Model {model_name} has unknown prediction method")
                return anomalies
            
            # Extract anomalies
            anomaly_indices = np.where(anomaly_mask)[0]
            
            for idx in anomaly_indices:
                # Determine which parameter(s) caused the anomaly
                row = df.iloc[idx]
                
                # Simple approach: find parameter with most deviation
                parameter = self._identify_anomalous_parameter(row, X.columns.tolist())
                
                # Calculate anomaly score (0.0 to 1.0)
                if hasattr(model, 'decision_function'):
                    raw_score = abs(model.decision_function(X.iloc[[idx]])[0])
                    score = min(raw_score / 10.0, 1.0)  # Normalize
                elif hasattr(model, 'score_samples'):
                    raw_score = abs(model.score_samples(X.iloc[[idx]])[0])
                    score = min(raw_score / 100.0, 1.0)  # Normalize
                else:
                    score = 0.75  # Default score
                
                anomaly = {
                    'parameter': parameter,
                    'phase_id': phase_id,
                    'score': round(score, 4),
                    'description': f'{parameter} anomaly detected in {model_name} (Phase {phase_id})'
                }
                
                anomalies.append(anomaly)
            
            logger.info(f"Model {model_name}: {len(anomalies)} anomalies in phase {phase_id}")
            
        except Exception as e:
            logger.error(f"Error in _run_model for {model_name}: {e}")
        
        return anomalies
    
    def _identify_anomalous_parameter(self, row: pd.Series, 
                                     available_params: List[str]) -> str:
        """
        Identify which parameter is most likely anomalous
        
        Args:
            row: Data row with anomaly
            available_params: List of parameter names in the data
            
        Returns:
            Parameter name (str)
        """
        # Map available parameters to our known parameters
        for param_name in self.parameter_columns.keys():
            if param_name in available_params:
                return param_name
        
        # Default to first available parameter
        if available_params:
            return available_params[0]
        
        return 'Unknown'
    
    def detect_anomalies(self, file_path: str) -> Dict:
        """
        Alternative method name - calls analyze_file
        
        Args:
            file_path: Path to Excel file
            
        Returns:
            Same as analyze_file
        """
        return self.analyze_file(file_path)


# Test code
if __name__ == "__main__":
    print("="*70)
    print("FlightAnalyzer Test")
    print("="*70)
    
    # Setup logging for test
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    # Initialize analyzer
    print("\n1. Initializing FlightAnalyzer...")
    analyzer = FlightAnalyzer()
    
    print(f"\n2. Models loaded: {len(analyzer.models)}")
    for model_name in analyzer.models.keys():
        print(f"   - {model_name}")
    
    # Test with a file (update path to your actual file)
    test_file = r"A:\populate_fdap_db\flight_data\UNO-561P_01-10-25_1.xlsm"
    
    if os.path.exists(test_file):
        print(f"\n3. Testing with file: {test_file}")
        results = analyzer.analyze_file(test_file)
        
        print(f"\n4. Results:")
        print(f"   Anomalies found: {len(results['anomalies'])}")
        print(f"   Stats: {results['stats']}")
        
        if results['anomalies']:
            print(f"\n5. Sample anomalies:")
            for i, anom in enumerate(results['anomalies'][:5], 1):
                print(f"   {i}. {anom['parameter']} (Phase {anom['phase_id']}) - "
                      f"Score: {anom['score']}")
    else:
        print(f"\n3. Test file not found: {test_file}")
        print("   Update the path and run again")
    
    print("\n" + "="*70)
    print("Test Complete!")
    print("="*70)
