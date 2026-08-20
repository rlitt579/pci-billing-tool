#!/bin/bash
echo "PCI Billing Tool - Mac Build"
echo ""
pip3 install customtkinter pandas openpyxl pyinstaller --quiet
CTK=$(python3 -c "import customtkinter; print(customtkinter.__path__[0])")
pyinstaller --onefile --windowed --name "PCI_Billing_Tool" --add-data "$CTK:customtkinter" pci_simple.py
echo ""
if [ -d "dist/PCI_Billing_Tool.app" ] || [ -f "dist/PCI_Billing_Tool" ]; then
    echo "Build successful: dist/PCI_Billing_Tool"
else
    echo "Build failed. Check errors above."
fi
