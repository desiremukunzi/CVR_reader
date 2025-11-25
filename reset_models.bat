@echo off
echo ================================================
echo  RESETTING ANOMALY DETECTION MODELS
echo ================================================
echo.
echo This will delete the old models so they can be
echo retrained with the new contamination rate (0.5%)
echo.
pause

cd /d "%~dp0"

if exist "flight_data\trained_anomaly_models.joblib" (
    echo Deleting old model file...
    del "flight_data\trained_anomaly_models.joblib"
    echo ✓ Old models deleted successfully!
) else (
    echo No model file found - already clean!
)

echo.
echo ================================================
echo  NEXT STEPS:
echo ================================================
echo 1. Restart Flask: python app.py
echo 2. Run analysis with checkbox checked
echo 3. New models will be trained automatically
echo.
pause
