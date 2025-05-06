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