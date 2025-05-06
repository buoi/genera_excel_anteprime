from typing import Tuple
import os.path

def check_excel_file(file_path: str) -> Tuple[bool, str]:
    """
    Check if the given file path is a valid Excel file.
    
    Args:
        file_path: Path to the file to check
        
    Returns:
        Tuple containing:
            - Boolean indicating if the file is valid
            - Message with details about validation result
    """
    # Check if file exists
    if not os.path.isfile(file_path):
        return False, "Il file non esiste"
    
    # Check file extension
    valid_extensions = ['.xlsx', '.xls']
    file_extension = os.path.splitext(file_path)[1].lower()
    
    if file_extension not in valid_extensions:
        return False, f"Il file non è un file Excel valido. Estensione rilevata: {file_extension}"
    
    # If we get here, basic checks passed
    return True, "File Excel valido"

import pandas as pd
from typing import Tuple, Dict, Any

def parse_excel_file(file_path: str) -> Tuple[bool, Dict[str, Any], str]:
    """
    Parse an Excel file and extract key information.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        Tuple containing:
            - Boolean indicating success
            - Dictionary with extracted information
            - Message with details about parsing result
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Extract basic information
        info = {
            "rows": len(df),
            "columns": len(df.columns),
            "column_names": list(df.columns),
        }
        
        return True, info, "File Excel analizzato con successo"
    
    except Exception as e:
        return False, {}, f"Errore nell'analisi del file Excel: {str(e)}"

def check_image_folder(folder_path: str) -> Tuple[bool, str]:
    """
    Check if the given path is a valid folder containing images.
    
    Args:
        folder_path: Path to the folder to check
        
    Returns:
        Tuple containing:
            - Boolean indicating if the folder is valid
            - Message with details about validation result
    """
    # Check if folder exists
    if not os.path.isdir(folder_path):
        return False, "La cartella non esiste"
    
    # Check if folder contains any images
    image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff']
    has_images = False
    
    # Walk through the folder
    for file in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file)
        if os.path.isfile(file_path):
            # Check file extension
            ext = os.path.splitext(file_path)[1].lower()
            if ext in image_extensions:
                has_images = True
                break
    
    if not has_images:
        return False, "La cartella non contiene immagini supportate"
    
    # If we get here, all checks passed
    return True, "Cartella immagini valida"

def analyze_image_folder(folder_path: str) -> Tuple[bool, Dict[str, Any], str]:
    """
    Analyze an image folder and extract key information.
    
    Args:
        folder_path: Path to the folder to check
        
    Returns:
        Tuple containing:
            - Boolean indicating success
            - Dictionary with image information
            - Message with analysis details
    """
    # Check if folder exists
    if not os.path.isdir(folder_path):
        return False, {}, "La cartella non esiste"
    
    # Analyze images in the folder
    image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff']
    image_count = 0
    image_types = {}
    
    try:
        # Walk through the folder
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            if os.path.isfile(file_path):
                # Check file extension
                ext = os.path.splitext(file_path)[1].lower()
                if ext in image_extensions:
                    image_count += 1
                    # Count by image type
                    image_types[ext] = image_types.get(ext, 0) + 1
        
        if image_count == 0:
            return False, {}, "La cartella non contiene immagini supportate"
        
        # Prepare info dictionary
        info = {
            "total_images": image_count,
            "image_types": image_types
        }
        
        return True, info, "Cartella analizzata con successo"
    
    except Exception as e:
        return False, {}, f"Errore nell'analisi della cartella: {str(e)}"