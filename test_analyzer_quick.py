"""
Quick test for FlightAnalyzer
Run this to verify your setup before using with flight_processor.py
"""

import sys
import os

print("="*70)
print("FlightAnalyzer Quick Test")
print("="*70)

# Add path
sys.path.insert(0, r'A:\CVR_reader')

# Test 1: Can we import?
print("\n1. Testing import...")
try:
    from populate_db import FlightAnalyzer
    print("   ✓ FlightAnalyzer imported successfully!")
except ImportError as e:
    print(f"   ❌ Import failed: {e}")
    print("   Make sure flight_analyzer.py is in A:\\CVR_reader\\")
    exit(1)

# Test 2: Can we initialize?
print("\n2. Initializing FlightAnalyzer...")
try:
    analyzer = FlightAnalyzer()
    print(f"   ✓ FlightAnalyzer initialized!")
    print(f"   Models loaded: {len(analyzer.models)}")
except Exception as e:
    print(f"   ❌ Initialization failed: {e}")
    exit(1)

# Test 3: List models
print("\n3. Loaded models:")
if analyzer.models:
    for i, name in enumerate(analyzer.models.keys(), 1):
        print(f"   {i}. {name}")
    
    print(f"\n   Phase distribution:")
    print(f"   - Before Takeoff (Phase 1): {len(analyzer.phase_models[1])} models")
    print(f"   - Airborne (Phase 2): {len(analyzer.phase_models[2])} models")
    print(f"   - After Landing (Phase 3): {len(analyzer.phase_models[3])} models")
else:
    print("   ⚠️  No models loaded!")
    print("   Check A:\\CVR_reader\\flight_data\\ for .pkl or .joblib files")

# Test 4: Test with a file (if available)
print("\n4. Testing file analysis...")
test_file = r"A:\populate_fdap_db\flight_data\UNO-561P_01-10-25_1.xlsm"

if os.path.exists(test_file):
    try:
        results = analyzer.analyze_file(test_file)
        
        print(f"   ✓ Analysis complete!")
        print(f"   Anomalies detected: {len(results['anomalies'])}")
        print(f"   Stats: {results['stats']}")
        
        if results['anomalies']:
            print(f"\n   Sample anomalies:")
            for i, anom in enumerate(results['anomalies'][:3], 1):
                print(f"   {i}. {anom['parameter']} (Phase {anom['phase_id']}) - "
                      f"Score: {anom['score']:.4f}")
        else:
            print("   No anomalies detected (this is okay for test)")
            
    except Exception as e:
        print(f"   ❌ Analysis failed: {e}")
        import traceback
        traceback.print_exc()
else:
    print(f"   ⚠️  Test file not found: {test_file}")
    print("   Update the path to test with an actual file")

# Summary
print("\n" + "="*70)
print("Test Summary")
print("="*70)

if len(analyzer.models) > 0:
    print("✅ FlightAnalyzer is ready to use!")
    print("\nNext steps:")
    print("1. Make sure flight_processor.py has:")
    print("   analyzer = FlightAnalyzer()  # Line ~635")
    print("2. Run: python flight_processor.py")
    print("3. Check logs for 'Detected X anomalies'")
else:
    print("⚠️  No models loaded")
    print("\nTo fix:")
    print("1. Check A:\\CVR_reader\\flight_data\\ exists")
    print("2. Add .pkl or .joblib model files to that folder")
    print("3. Run this test again")

print("="*70)
