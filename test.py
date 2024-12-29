import sys

# Print Python version and paths
print("Python version:", sys.version)
print("Python path:", sys.executable)
print("Module paths:", sys.path)

# Test import
try:
    from docx import Document
    print("python-docx is successfully imported!")
except ModuleNotFoundError as e:
    print("Error importing python-docx:", e)