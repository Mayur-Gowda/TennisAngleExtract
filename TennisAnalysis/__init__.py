import os
from .angle import Angle
from .excel import ExcelSave, sheetSetup
from .roi import ROI
from .orgData import CreateDF
from .compute import Compute


# Define a function to initialize directory structure
def initialize_directories():
    # Get the current directory where the package (__init__.py) resides
    current_dir = os.path.dirname(__file__)

    # Define the paths for the directories
    croppedImages_path = os.path.join(current_dir, '..', 'CroppedImages')
    analyzedAngles_path = os.path.join(current_dir, '..', 'AnalyzedAngles')

    # Check if the directories already exist
    if not os.path.exists(croppedImages_path):
        # If not, create Folder1
        os.makedirs(croppedImages_path)

    if not os.path.exists(analyzedAngles_path):
        # If not, create Folder2
        os.makedirs(analyzedAngles_path)

    excelAngles_path = os.path.join(analyzedAngles_path, 'ExcelSheets')
    csvAngles_path = os.path.join(analyzedAngles_path, 'CSVFiles')

    # Create subfolders inside Folder2 if they don't exist
    for folder in [excelAngles_path, csvAngles_path]:
        if not os.path.exists(folder):
            os.makedirs(folder)


# Call the function to initialize directories when the package is imported
initialize_directories()
