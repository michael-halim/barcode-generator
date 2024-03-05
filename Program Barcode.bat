@ECHO OFF

cd C:\Users\Michael Halim\Downloads\Placard\barcode-generator

call env\Scripts\activate

if %errorlevel% neq 0 (
    echo Error: Environment activation failed.
    exit /b 1
)

call python placard.py

PAUSE