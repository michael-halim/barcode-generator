@ECHO OFF

cd "D:\Program\Visual Studio Code\Python\Placard Project\barcode-generator"

call env\Scripts\activate

if %errorlevel% neq 0 (
    echo Error: Environment activation failed.
    exit /b 1
)

call python placard.py

PAUSE