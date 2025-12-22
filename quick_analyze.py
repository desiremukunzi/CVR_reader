"""
Quick Anomaly Analysis Script - Bypass Audio Compliance Check
==============================================================

This script allows you to directly analyze a flight Excel file for anomalies
without going through the audio compliance check process.

Usage:
    1. Set the FLIGHT_FILE_PATH to your Excel file
    2. Set ADD_TO_TRAINING to True if you want to add this flight to training database
    3. Run: python quick_analyze.py
    4. Anomaly report will automatically open in your browser

Author: Flight Safety Analysis System
"""

import os
import sys
import webbrowser
from datetime import datetime
from pathlib import Path

# Add the current directory to the path so we can import flight_analyzer
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flight_analyzer import FlightAnalyzer, PARAMETERS_TO_ANALYZE

# ============================================================================
# CONFIGURATION - MODIFY THESE SETTINGS
# ============================================================================

# Path to the Excel file you want to analyze
FLIGHT_FILE_PATH = r'A:\Onedrive\RAF-61504\April\UNO-561P_24-04-25_1.xlsm'

# Excel sheet name containing clean flight data
SHEET_NAME = 'Clean Data'

# Add this flight to training database after analysis?
# True = Add to database (increases historical data for future analysis)
# False = Just analyze without adding (default)
ADD_TO_TRAINING = False

# Automatically open the report in browser after analysis?
AUTO_OPEN_BROWSER = True

# Data folder where models and historical data are stored
MODEL_DATA_FOLDER = 'flight_data'

# ============================================================================
# DO NOT MODIFY BELOW THIS LINE (unless you know what you're doing!)
# ============================================================================

def print_header():
    """Print a nice header for the analysis script."""
    print("\n" + "="*70)
    print("  MI-17 FLIGHT ANOMALY ANALYSIS - QUICK ANALYSIS")
    print("="*70)
    print(f"\n📁 Flight File: {os.path.basename(FLIGHT_FILE_PATH)}")
    print(f"📊 Sheet Name: {SHEET_NAME}")
    print(f"🎯 Add to Training: {'Yes' if ADD_TO_TRAINING else 'No'}")
    print(f"🌐 Auto-open Browser: {'Yes' if AUTO_OPEN_BROWSER else 'No'}")
    print("="*70 + "\n")

def validate_file(file_path):
    """
    Validate that the flight file exists and is an Excel file.
    
    Args:
        file_path (str): Path to the flight Excel file
        
    Returns:
        bool: True if valid, False otherwise
    """
    if not os.path.exists(file_path):
        print(f"❌ ERROR: File does not exist: {file_path}")
        print("   Please check the path and try again.")
        return False
    
    if not os.path.isfile(file_path):
        print(f"❌ ERROR: Path is not a file: {file_path}")
        return False
    
    if not file_path.endswith(('.xlsx', '.xlsm', '.xls')):
        print(f"❌ ERROR: File is not an Excel file: {file_path}")
        print("   Please use a .xlsx, .xlsm, or .xls file.")
        return False
    
    return True

def analyze_flight_direct(file_path, sheet_name='Clean Data', add_to_training=False, 
                          data_folder='flight_data', auto_open=True):
    """
    Directly analyze a flight Excel file for anomalies and generate a report.
    
    Args:
        file_path (str): Path to the Excel file containing flight data
        sheet_name (str): Name of the sheet to read from the Excel file
        add_to_training (bool): Whether to add this flight to training database
        data_folder (str): Folder where models and historical data are stored
        auto_open (bool): Whether to automatically open the report in browser
        
    Returns:
        dict: Analysis results or None if failed
    """
    print_header()
    
    # Validate file
    if not validate_file(file_path):
        return None
    
    print("✅ File validation passed\n")
    
    # Initialize the flight analyzer
    print("🚀 Initializing Flight Analyzer...")
    analyzer = FlightAnalyzer(data_folder=data_folder)
    
    # Check if models exist
    if not analyzer.trained_models:
        print("\n❌ ERROR: No trained models found!")
        print("   You need to train models first using quick_train.py")
        print("   Run: python quick_train.py")
        return None
    
    print(f"✅ Loaded {len(analyzer.trained_models)} trained models")
    print(f"   Historical database: {analyzer.historical_data['flight_id'].nunique()} flights, "
          f"{len(analyzer.historical_data):,} data points\n")
    
    # Analyze the flight
    print(f"{'='*70}")
    print(f"  ANALYZING FLIGHT DATA")
    print(f"{'='*70}\n")
    
    print(f"📊 Analyzing: {os.path.basename(file_path)}")
    print(f"   Sheet: {sheet_name}")
    print(f"   Add to training: {'Yes' if add_to_training else 'No'}\n")
    
    try:
        # Run the analysis
        results = analyzer.analyze_flight(
            excel_path=file_path,
            sheet_name=sheet_name,
            add_to_training=add_to_training
        )
        
        # Check for errors
        if 'error' in results:
            print(f"❌ Analysis failed: {results['error']}")
            return None
        
        # Print results summary
        print(f"\n{'='*70}")
        print(f"  ANALYSIS RESULTS")
        print(f"{'='*70}\n")
        
        flight_id = results['flight_id']
        total_points = results['total_data_points']
        anomaly_count = results['anomaly_count']
        anomaly_pct = results['anomaly_percentage']
        
        print(f"✅ Analysis complete!")
        print(f"\n📈 Flight Statistics:")
        print(f"   - Flight ID: {flight_id}")
        print(f"   - Total data points: {total_points:,}")
        print(f"   - Anomalies detected: {anomaly_count}")
        print(f"   - Anomaly rate: {anomaly_pct}%")
        
        if add_to_training:
            print(f"\n💾 Flight added to training database")
            print(f"   - Total flights in database: {results['total_historical_flights']}")
        
        # Print phase summary
        print(f"\n📊 Anomalies by Phase:")
        for phase, stats in results['phases_summary'].items():
            if stats['total_points'] > 0:
                print(f"   - {phase.capitalize()}: {stats['anomaly_count']} anomalies "
                      f"({stats['anomaly_percentage']:.1f}%) out of {stats['total_points']:,} points")
        
        # Print anomalies by parameter
        if anomaly_count > 0:
            print(f"\n🎯 Anomalies by Parameter:")
            param_counts = {}
            for anomaly in results['anomalies']:
                param = anomaly['parameter']
                param_counts[param] = param_counts.get(param, 0) + 1
            
            for param in PARAMETERS_TO_ANALYZE:
                count = param_counts.get(param, 0)
                if count > 0:
                    print(f"   ✅ {param}: {count} anomaly(ies)")
                else:
                    print(f"   ⚪ {param}: No anomalies")
        else:
            print(f"\n✅ No anomalies detected - Flight appears normal!")
        
        # Generate HTML report
        print(f"\n{'='*70}")
        print(f"  GENERATING REPORT")
        print(f"{'='*70}\n")
        
        report_path = generate_html_report(results, file_path)
        
        if report_path and auto_open:
            print(f"🌐 Opening report in browser...")
            webbrowser.open('file://' + os.path.abspath(report_path))
            print(f"✅ Report opened in default browser")
        elif report_path:
            print(f"📄 Report saved to: {report_path}")
            print(f"   Open manually in browser to view")
        
        print(f"\n{'='*70}")
        print(f"  🎉 ANALYSIS COMPLETE!")
        print(f"{'='*70}\n")
        
        return results
        
    except FileNotFoundError:
        print(f"❌ ERROR: File not found: {file_path}")
        return None
    except ValueError as ve:
        if "Worksheet named" in str(ve):
            print(f"❌ ERROR: Sheet '{sheet_name}' not found in Excel file")
            print(f"   Please check the sheet name and try again.")
        else:
            print(f"❌ ERROR: {ve}")
        return None
    except Exception as e:
        print(f"❌ ERROR: Unexpected error during analysis: {e}")
        import traceback
        traceback.print_exc()
        return None

def generate_html_report(results, flight_file_path):
    """
    Generate an HTML report from analysis results.
    
    Args:
        results (dict): Analysis results from FlightAnalyzer
        flight_file_path (str): Path to the analyzed flight file
        
    Returns:
        str: Path to the generated HTML report
    """
    from flask import render_template_string
    
    # Load the report template
    template_path = os.path.join(os.path.dirname(__file__), 'templates', 'anomaly_report.html')
    
    if not os.path.exists(template_path):
        print(f"⚠️  Warning: Report template not found at {template_path}")
        print(f"   Creating basic report...")
        return create_basic_report(results, flight_file_path)
    
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            template = f.read()
        
        # Render the template with results
        # Note: This is a simplified version - the actual Flask app does more processing
        html_content = render_template_string(template, **results)
        
        # Save to file
        report_filename = f"anomaly_report_flight_{results['flight_id']}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
        report_path = os.path.join('reports', report_filename)
        
        os.makedirs('reports', exist_ok=True)
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"📄 Report saved: {report_path}")
        return report_path
        
    except Exception as e:
        print(f"⚠️  Warning: Could not generate full report: {e}")
        print(f"   Creating basic report...")
        return create_basic_report(results, flight_file_path)

def create_basic_report(results, flight_file_path):
    """
    Create a basic HTML report if the full template is not available.
    
    Args:
        results (dict): Analysis results
        flight_file_path (str): Path to analyzed flight file
        
    Returns:
        str: Path to the generated basic report
    """
    flight_id = results['flight_id']
    anomaly_count = results['anomaly_count']
    total_points = results['total_data_points']
    anomaly_pct = results['anomaly_percentage']
    
    html = f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Anomaly Report - Flight {flight_id}</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                max-width: 1200px;
                margin: 0 auto;
                padding: 20px;
                background: #f5f5f5;
            }}
            .header {{
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
                padding: 30px;
                border-radius: 10px;
                margin-bottom: 20px;
            }}
            .card {{
                background: white;
                padding: 20px;
                border-radius: 10px;
                margin-bottom: 20px;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }}
            .stat {{
                display: inline-block;
                margin: 10px 20px;
                text-align: center;
            }}
            .stat-value {{
                font-size: 2em;
                font-weight: bold;
                color: #667eea;
            }}
            .stat-label {{
                color: #666;
                font-size: 0.9em;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
            }}
            th, td {{
                padding: 12px;
                text-align: left;
                border-bottom: 1px solid #ddd;
            }}
            th {{
                background: #667eea;
                color: white;
            }}
            tr:hover {{
                background: #f5f5f5;
            }}
            .anomaly {{
                color: #e74c3c;
                font-weight: bold;
            }}
        </style>
    </head>
    <body>
        <div class="header">
            <h1>🛩️ MI-17 Flight Anomaly Analysis Report</h1>
            <p><strong>Flight ID:</strong> {flight_id}</p>
            <p><strong>File:</strong> {os.path.basename(flight_file_path)}</p>
            <p><strong>Generated:</strong> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
        </div>
        
        <div class="card">
            <h2>📊 Summary Statistics</h2>
            <div class="stat">
                <div class="stat-value">{total_points:,}</div>
                <div class="stat-label">Total Data Points</div>
            </div>
            <div class="stat">
                <div class="stat-value" style="color: {'#e74c3c' if anomaly_count > 0 else '#27ae60'}">{anomaly_count}</div>
                <div class="stat-label">Anomalies Detected</div>
            </div>
            <div class="stat">
                <div class="stat-value">{anomaly_pct}%</div>
                <div class="stat-label">Anomaly Rate</div>
            </div>
        </div>
        
        <div class="card">
            <h2>📋 Phase Breakdown</h2>
            <table>
                <thead>
                    <tr>
                        <th>Phase</th>
                        <th>Total Points</th>
                        <th>Anomalies</th>
                        <th>Anomaly Rate</th>
                    </tr>
                </thead>
                <tbody>
    """
    
    for phase, stats in results['phases_summary'].items():
        if stats['total_points'] > 0:
            html += f"""
                    <tr>
                        <td><strong>{phase.capitalize()}</strong></td>
                        <td>{stats['total_points']:,}</td>
                        <td class="anomaly">{stats['anomaly_count']}</td>
                        <td>{stats['anomaly_percentage']:.1f}%</td>
                    </tr>
            """
    
    html += """
                </tbody>
            </table>
        </div>
    """
    
    if anomaly_count > 0:
        html += """
        <div class="card">
            <h2>🎯 Detected Anomalies</h2>
            <table>
                <thead>
                    <tr>
                        <th>Parameter</th>
                        <th>Phase</th>
                        <th>Time</th>
                        <th>Value</th>
                    </tr>
                </thead>
                <tbody>
        """
        
        for anomaly in results['anomalies']:
            html += f"""
                    <tr>
                        <td><strong>{anomaly['parameter']}</strong></td>
                        <td>{anomaly['phase']}</td>
                        <td>{anomaly['time']:.0f}</td>
                        <td class="anomaly">{anomaly['value']:.2f}</td>
                    </tr>
            """
        
        html += """
                </tbody>
            </table>
        </div>
        """
    else:
        html += """
        <div class="card">
            <h2>✅ All Clear!</h2>
            <p style="color: #27ae60; font-size: 1.2em;">No anomalies detected in this flight. All parameters within normal ranges.</p>
        </div>
        """
    
    html += """
        <div class="card">
            <p style="color: #666; text-align: center;">
                <em>Note: For detailed charts and visualizations, please use the full web application.</em>
            </p>
        </div>
    </body>
    </html>
    """
    
    # Save report
    report_filename = f"anomaly_report_flight_{flight_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    report_path = os.path.join('reports', report_filename)
    
    os.makedirs('reports', exist_ok=True)
    
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"📄 Basic report saved: {report_path}")
    return report_path

def main():
    """Main entry point for the analysis script."""
    try:
        results = analyze_flight_direct(
            file_path=FLIGHT_FILE_PATH,
            sheet_name=SHEET_NAME,
            add_to_training=ADD_TO_TRAINING,
            data_folder=MODEL_DATA_FOLDER,
            auto_open=AUTO_OPEN_BROWSER
        )
        
        if results:
            print("\n✅ Analysis successful!")
            print(f"   Results saved and report generated.")
            sys.exit(0)
        else:
            print("\n❌ Analysis failed.")
            print(f"   Please check the errors above and try again.")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\n\n❌ Analysis interrupted by user (Ctrl+C)")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
