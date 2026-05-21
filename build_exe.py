import PyInstaller.__main__
import os

# Define the paths
app_dir = os.path.dirname(os.path.abspath(__file__))

PyInstaller.__main__.run([
    'run.py',
    '--name=ContractCompareApp',
    '--onedir',          # Use onedir instead of onefile for faster startup and less antivirus false positives
    '--clean',
    '--noconfirm',
    '--add-data=app.py;.',
    '--add-data=utils.py;.',
    '--add-data=excel_generator.py;.',
    '--copy-metadata=streamlit',
    '--copy-metadata=google-genai',
    '--hidden-import=openpyxl',
    '--hidden-import=pdfplumber',
    '--hidden-import=google.genai',
    '--hidden-import=toml',
])
