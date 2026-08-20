@echo off
echo PCI Billing Tool - Build
echo.
pip install customtkinter pandas openpyxl pyinstaller --quiet
for /f "delims=" %%i in ('python -c "import customtkinter; print(customtkinter.__path__[0])"') do set CTK=%%i
pyinstaller --onefile --windowed --name "PCI_Billing_Tool" --add-data "%CTK%;customtkinter" pci_simple.py
echo.
if exist "dist\PCI_Billing_Tool.exe" (
    echo Build successful: dist\PCI_Billing_Tool.exe
) else (
    echo Build failed. Check errors above.
)
pause
