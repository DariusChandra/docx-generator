## preparation for running the script
brew install tcl-tk
brew install python-tk
pip3 install openpyxl
pip3 install pandas
pip3 install python-docx

python3 docx_generator.py

## to make windows executable

pip3 install pyinstaller
pyinstaller --onefile --windowed docx_generator.py