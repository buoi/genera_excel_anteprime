from typing import Tuple, Dict, Any, List
import os, os.path, re
import numpy as np
import pandas as pd
import xlsxwriter
from PIL import Image

from numbers_parser import Document as NumbersDocument

# List of required columns for our specific Excel format
REQUIRED_COLUMNS = [
    "CODICE TAILOR", "POSIZIONE", "CATEGORIA", "FOTO", "FOTO DETTAGLIO", 
    "COMPOSIZIONE", "FORNITORE", "ART. FORNITORE", "UNITA' DI MISURA", 
    "ALTEZZA", "PESO", "ARMATURA", "LAVORAZIONE", "DESCRIZIONE", 
    "MOTIVO", "SOSTENIBILITA'", "CERTIFICAZIONE"
]

def normalize_image_filename(filename: str) -> str:
    """
    Verifica e normalizza il nome di un file immagine, correggendo piccoli errori.
    
    Args:
        filename: Il nome del file da normalizzare
        
    Returns:
        Nome del file normalizzato
    
    Esempi:
        >>> normalize_image_filename("L1170719.JPG")
        'L1170719.JPG'
        >>> normalize_image_filename("L1170719JPG")
        'L1170719.JPG'
        >>> normalize_image_filename("L1170719jpg")
        'L1170719.jpg'
        >>> normalize_image_filename("L1170719")
        'L1170719.JPG'
    """
    import re
    import os.path
    
    # Lista delle estensioni immagine supportate
    supported_extensions = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff']
    
    # Se il filename è vuoto o None, restituisci stringa vuota
    if not filename:
        return ""
    
    # Normalizza al minuscolo per semplificare l'analisi
    filename = filename.strip()
    
    # Controllo se ha già un'estensione corretta
    base, ext = os.path.splitext(filename)
    
    # Se l'estensione esiste e inizia con un punto, verifichiamo se è supportata
    if ext and ext.startswith('.'):
        # Rimuovi il punto e normalizza
        ext_clean = ext[1:].lower()
        if ext_clean in supported_extensions:
            # L'estensione è già corretta, la normalizziamo
            return f"{base}.{ext_clean.upper()}"
    
    # Se siamo qui, l'estensione è mancante o non ha il punto
    
    # Pattern per trovare estensioni incorporate nel nome senza punto
    pattern = r'(.+?)(jpe?g|png|gif|bmp|tiff)$'
    match = re.match(pattern, filename.lower())
    
    if match:
        # Abbiamo trovato un'estensione incorporata, estraiamola
        base = match.group(1)
        ext = match.group(2)
        
        # Normalizza jpg/jpeg
        if ext == 'jpeg':
            ext = 'jpg'
            
        return f"{base}.{ext.upper()}"
    
    # Se non troviamo un'estensione, assumiamo JPG come default
    return f"{filename}.JPG"

def normalize_column_name(column: str) -> str:
    """
    Normalize a column name for fuzzy matching.
    
    Args:
        column: Original column name
        
    Returns:
        Normalized column name (lowercase, no spaces, no special chars)
    """
    # Convert to lowercase
    normalized = column.lower()
    # Remove spaces and special characters
    normalized = re.sub(r'[^a-z0-9]', '', normalized)
    return normalized

def match_column_name(column: str, df_columns: List[str]) -> str:
    """
    Try to match a required column name with the actual columns in the dataframe.
    
    Args:
        column: Required column name
        df_columns: List of column names in the dataframe
        
    Returns:
        The matched column name from df_columns, or None if no match found
    """
    # First try exact match
    if column in df_columns:
        return column
    
    # Try case-insensitive match
    for df_col in df_columns:
        if column.lower() == df_col.lower():
            return df_col
    
    # Try fuzzy match
    normalized_req = normalize_column_name(column)
    for df_col in df_columns:
        normalized_df = normalize_column_name(df_col)
        
        # Check if normalized strings are very similar
        if normalized_req == normalized_df:
            return df_col
        
        # Check if one is contained in the other
        if (normalized_req in normalized_df) or (normalized_df in normalized_req):
            # Additional check: they should be at least 70% similar in length
            min_len = min(len(normalized_req), len(normalized_df))
            max_len = max(len(normalized_req), len(normalized_df))
            if min_len / max_len >= 0.7:
                return df_col
    
    # No match found
    return None

def numbers_to_dataframe(file_path: str) -> pd.DataFrame:
    """
    Convert a Numbers file to a pandas DataFrame using numbers-parser.
    
    Args:
        file_path: Path to the Numbers file
        
    Returns:
        DataFrame containing the data from the Numbers file
    """
    if NumbersDocument is None:
        raise ImportError("numbers-parser library is not installed. Please install with: pip install numbers-parser")
    
    # Open the Numbers document
    doc = NumbersDocument(file_path)
    
    # Use the first sheet and first table by default
    sheet = doc.sheets[0]
    table = sheet.tables[0]
    
    # Get rows data - this returns a list of lists with cell objects
    rows_data = table.rows()
    
    if len(rows_data) == 0:
        return pd.DataFrame()
    
    # Extract headers from the first row
    headers = [cell.value if hasattr(cell, 'value') else f"Column_{i}" 
               for i, cell in enumerate(rows_data[0])]
    
    # Extract data rows (skipping header row)
    data = []
    for row in rows_data[1:]:
        row_data = {}
        for i, cell in enumerate(row):
            if i < len(headers):  # Ensure we don't go out of bounds with headers
                column_name = headers[i]
                row_data[column_name] = cell.value if hasattr(cell, 'value') else None
        data.append(row_data)
    
    # Create DataFrame
    return pd.DataFrame(data)

def check_excel_file(file_path: str) -> Tuple[bool, str]:
    """
    Check if the given file path is a valid Excel or Numbers file.
    
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
    valid_extensions = ['.xlsx', '.xls', '.numbers']
    file_extension = os.path.splitext(file_path)[1].lower()
    
    if file_extension not in valid_extensions:
        return False, f"Il file non è un formato valido. Estensione rilevata: {file_extension}"
    
    # If we get here, basic checks passed
    return True, "File valido"

def parse_excel_file(file_path: str) -> Tuple[bool, Dict[str, Any], str]:
    """
    Parse an Excel or Numbers file and extract key information.
    Also validates that the file contains all required columns.
    Removes rows with missing FOTO values and rows with duplicate FOTO values.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        Tuple containing:
            - Boolean indicating success
            - Dictionary with extracted information
            - Message with details about parsing result
    """
    try:
        # Check if it's a Numbers file
        file_extension = os.path.splitext(file_path)[1].lower()
        
        # Read the Excel or Numbers file into a pandas DataFrame
        if file_extension == '.numbers':
            if NumbersDocument is None:
                return False, {}, "Per supportare i file Numbers, installa la libreria numbers-parser: pip install numbers-parser"
            
            try:
                df = numbers_to_dataframe(file_path)
            except Exception as e:
                return False, {}, f"Errore nell'apertura del file Numbers: {str(e)}"
        else:
            # Regular Excel file
            df = pd.read_excel(file_path)
        
        # Extract basic information
        original_rows = len(df)
        info = {
            "rows": original_rows,
            "columns": len(df.columns),
            "column_names": list(df.columns),
        }
        
        # Check for required columns with fuzzy matching
        missing_columns = []
        column_mapping = {}  # Maps required column names to actual column names
        
        for column in REQUIRED_COLUMNS:
            matched_column = match_column_name(column, list(df.columns))
            if matched_column:
                column_mapping[column] = matched_column
            else:
                missing_columns.append(column)
        
        if missing_columns:
            return False, {}, f"Colonne mancanti: {', '.join(missing_columns)}"
        
        # Store the column mapping for later use
        info["column_mapping"] = column_mapping
        
        # Get FOTO column name
        foto_column = column_mapping["FOTO"]
        
        # Step 1: Count and remove rows with missing FOTO values
        missing_foto_mask = df[foto_column].isna() | (df[foto_column] == "")
        missing_foto_rows = missing_foto_mask.sum()
        
        # Remove rows with missing FOTO values
        if missing_foto_rows > 0:
            df = df[~missing_foto_mask].reset_index(drop=True)
            print(f"Removed {missing_foto_rows} rows with missing FOTO values")
        
        # Step 2: Deal with duplicates in FOTO column
        # Find duplicates
        foto_values = df[foto_column].astype(str)
        duplicate_mask = foto_values.duplicated(keep=False)
        duplicates = foto_values[duplicate_mask].unique()
        
        duplicate_rows_removed = 0
        
        # Remove duplicates, keeping rows with most information
        if len(duplicates) > 0:
            print(f"Found {len(duplicates)} duplicate FOTO values")
            
            # Create a helper column to count non-null values
            df['_info_count'] = df.notna().sum(axis=1)
            
            # Process each duplicate value
            cleaned_df = df.copy()
            for dup_value in duplicates:
                # Get all rows with this duplicate value
                dup_rows = df[df[foto_column] == dup_value]
                
                if len(dup_rows) > 1:
                    # Find the row with the most information
                    best_row_idx = dup_rows['_info_count'].idxmax()
                    
                    # Remove all other duplicates from cleaned_df
                    dup_indices = dup_rows.index.tolist()
                    dup_indices.remove(best_row_idx)  # Keep the best row
                    cleaned_df = cleaned_df.drop(dup_indices)
                    
                    # Count removed rows
                    duplicate_rows_removed += len(dup_indices)
            
            # Remove helper column
            cleaned_df = cleaned_df.drop('_info_count', axis=1)
            
            # Update dataframe
            df = cleaned_df
            
            print(f"Removed {duplicate_rows_removed} duplicate rows")
        
        # Total rows removed
        total_rows_removed = missing_foto_rows + duplicate_rows_removed
        
        # Record statistics
        info["rows_after_cleaning"] = len(df)
        info["total_rows_removed"] = total_rows_removed
        info["missing_foto_rows"] = missing_foto_rows
        info["duplicate_rows_removed"] = duplicate_rows_removed
        
        # Save the cleaned dataframe
        info["cleaned_df"] = df
        
        # Generate appropriate message
        if total_rows_removed > 0:
            message = f"File analizzato con successo. Rimosse {total_rows_removed} righe ({missing_foto_rows} senza FOTO, {duplicate_rows_removed} duplicate)."
        else:
            message = "File analizzato con successo."
            
        return True, info, message
    
    except Exception as e:
        return False, {}, f"Errore nell'analisi del file: {str(e)}"
    
def check_image_folder(folder_path: str) -> Tuple[bool, Dict[str, Any], str]:
    """
    Checks a folder to validate and count images at the first level only.
    
    Args:
        folder_path: Path to the folder to check
        
    Returns:
        Tuple containing:
            - Boolean indicating success
            - Dictionary with image information
            - Message with analysis details
    """
    print(f"DEBUG: Checking folder: {folder_path}")
    
    # Check if folder exists
    if not os.path.isdir(folder_path):
        print(f"DEBUG: Not a directory: {folder_path}")
        return False, {}, "La cartella non esiste"
    
    # Analyze images in the folder
    image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff']
    image_count = 0
    image_types = {}
    
    try:
        # Walk through the folder (only first level)
        print(f"DEBUG: Listing directory contents")
        for file in os.listdir(folder_path):
            file_path = os.path.join(folder_path, file)
            print(f"DEBUG: Found file: {file_path}")
            
            if os.path.isfile(file_path):
                # Check file extension
                ext = os.path.splitext(file_path)[1].lower()
                print(f"DEBUG: File extension: {ext}")
                
                if ext in image_extensions:
                    image_count += 1
                    # Count by image type
                    image_types[ext] = image_types.get(ext, 0) + 1
                    print(f"DEBUG: Found image: {file}, count now: {image_count}")
        
        print(f"DEBUG: Total images found: {image_count}")
        if image_count == 0:
            print("DEBUG: No supported images found")
            return False, {}, "La cartella non contiene immagini supportate"
        
        # Prepare info dictionary
        info = {
            "total_images": image_count,
            "image_types": image_types
        }
        
        print(f"DEBUG: Success, returning info: {info}")
        return True, info, "Cartella analizzata con successo"
    
    except Exception as e:
        print(f"DEBUG: Error occurred: {str(e)}")
        return False, {}, f"Errore nell'analisi della cartella: {str(e)}"

def generate_excel_output(df, foto_column, images_folder, output_path, 
                         progress_callback=None, status_callback=None):
    """Generate Excel output with thumbnails."""
    # This will be filled in later
    pass

def generate_crop_images(df, foto_column, images_folder, output_path,
                       progress_callback=None, status_callback=None):
    """Generate cropped images for website use."""
    # This will be filled in later
    pass

def generate_csv_output(df, output_path, column_mapping):
    """
    Generate CSV file for website import with specific column order and names.
    
    Args:
        df: DataFrame with valid rows (already cleaned and filtered)
        output_path: Path to save the CSV file
        column_mapping: Mapping from original column names to their actual names in the dataframe
        
    Returns:
        Dictionary with processing information
    """
    try:
        
        # Define the CSV columns in the exact order required
        CSV_COLUMNS = [
            "tax:casse", "tax:categorie", "cf:foto", "cf:foto_dettaglio", "tax:composizioni",
            "cf:codice", "cf:fornitore", "cf:art_fornitore", "cf:unita_di_misura", "cf:altezza",
            "cf:peso", "tax:armature", "tax:lavorazioni", "tax:descrizioni", "tax:motivi",
            "tax:sostenibili", "cf:certificazione"
        ]
        
        # Mapping from CSV columns to Excel columns
        CSV_TO_EXCEL_MAPPING = {
            "tax:casse": "POSIZIONE",
            "tax:categorie": "CATEGORIA",
            "cf:foto": "FOTO",
            "cf:foto_dettaglio": "FOTO DETTAGLIO",
            "tax:composizioni": "COMPOSIZIONE",
            "cf:codice": "CODICE TAILOR",
            "cf:fornitore": "FORNITORE",
            "cf:art_fornitore": "ART. FORNITORE",
            "cf:unita_di_misura": "UNITA' DI MISURA",
            "cf:altezza": "ALTEZZA",
            "cf:peso": "PESO",
            "tax:armature": "ARMATURA",
            "tax:lavorazioni": "LAVORAZIONE",
            "tax:descrizioni": "DESCRIZIONE",
            "tax:motivi": "MOTIVO",
            "tax:sostenibili": "SOSTENIBILITA'",
            "cf:certificazione": "CERTIFICAZIONE"
        }
        
        # Create a new DataFrame with the CSV columns
        csv_df = pd.DataFrame(columns=CSV_COLUMNS)
        total_rows = len(df)
        chunk_size = 60
        chunk_number = (total_rows + chunk_size - 1) // chunk_size  # Ceiling division
        csv_paths =[]

        split_df = np.array_split(df, chunk_number)
        
        for j, chunk_df in enumerate(split_df):
            # For each row in the valid data, map to the CSV format
            for i, row in chunk_df.iterrows():
                csv_row = {}
                
                # For each CSV column, get the corresponding Excel column and value
                for csv_col in CSV_COLUMNS:
                    excel_col = CSV_TO_EXCEL_MAPPING.get(csv_col)
                    if excel_col:
                        # Get the actual column name from the mapping (handles case/spelling variations)
                        actual_col = column_mapping.get(excel_col)
                        if actual_col in row:
                            csv_row[csv_col] = row[actual_col]
                        else:
                            csv_row[csv_col] = ""  # Handle missing columns
                    else:
                        csv_row[csv_col] = ""  # Handle unmapped columns
                
                # Add the row to the CSV DataFrame
                csv_df = pd.concat([csv_df, pd.DataFrame([csv_row])], ignore_index=True)
            
            # Create CSV file path
            csv_fiemane = f"import_campioni_{j+1}di{chunk_number}.csv"
            csv_path = os.path.join(output_path, csv_fiemane)
            csv_paths.append(csv_path)
            
            # Export DataFrame to CSV
            print("csv df chunk", csv_df.head())
            csv_df.to_csv(csv_path, index=False)
            csv_df = pd.DataFrame(columns=CSV_COLUMNS)
        
        return {"success": True, "csv_path": csv_paths, "rows": len(csv_df)}
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"Error generating CSV: {str(e)}")
        return {"success": False, "error": str(e)}

def process_files(excel_path: str, images_folder: str, output_path: str, 
                  progress_callback=None, status_callback=None):
    """
    Process files to generate Excel output with thumbnails.
    
    Args:
        excel_path: Path to the Excel file
        images_folder: Path to the folder containing images
        output_path: Path to save outputs
        progress_callback: Function to call with progress updates
        status_callback: Function to call with status messages
        
    Returns:
        Tuple containing:
            - Boolean indicating success
            - Dictionary with processing information
    """
    try:
        # Create output directory
        os.makedirs(output_path, exist_ok=True)
        
        # Get the cleaned DataFrame from the existing function
        _, info, _ = parse_excel_file(excel_path)
        df = info["cleaned_df"]
        column_mapping = info["column_mapping"]
        foto_column = column_mapping["FOTO"]
        
        # Set up output paths
        excel_output_path = os.path.join(output_path, "anteprime_excel.xlsx")
        crops_dir = os.path.join(output_path, "crops")
        csv_output_path = os.path.join(output_path, "website_import.csv")
        
        # Create necessary directories
        os.makedirs(crops_dir, exist_ok=True)
        
        # Create a temporary directory for thumbnails
        thumbs_dir = os.path.join(output_path, ".thumbnails")
        os.makedirs(thumbs_dir, exist_ok=True)
        
        # Create Excel workbook
        workbook = xlsxwriter.Workbook(excel_output_path)
        worksheet = workbook.add_worksheet()
        
        # Configure Excel worksheet
        worksheet.set_column(0, 0, 30)  # ANTEPRIMA column width - increased for larger thumbnails
        worksheet.set_column(1, len(REQUIRED_COLUMNS), 16)  # Other columns width
        worksheet.set_default_row(120, True)  # Row height for thumbnails - increased
        
        # Prepare ordered columns exactly as specified
        ordered_columns = [column_mapping[col] for col in REQUIRED_COLUMNS]
        
        # Prepare header with ANTEPRIMA as first column
        header = ['ANTEPRIMA'] + REQUIRED_COLUMNS
        
        # Write header row
        for j, col_name in enumerate(header):
            worksheet.write(0, j, col_name)
        
        # Track missing images and valid rows
        missing_images = []
        valid_rows_data = []
        
        # Process each row
        total_rows = len(df)
        if status_callback:
            status_callback("Elaborazione in corso...")
        
        df = df.fillna('')
        # Main processing loop
        for i, (_, row) in enumerate(df.iterrows()):
            if progress_callback:
                progress_callback(i+1, total_rows)
            
            # Get image path
            image_path = str(row[foto_column]) if not pd.isna(row[foto_column]) else ""
            image_path = normalize_image_filename(image_path)
            if not image_path:
                missing_images.append("(Vuoto)")
                continue
                    
            full_image_path = os.path.join(images_folder, image_path)
            
            # Check if image exists
            if os.path.isfile(full_image_path):
                try:
                    # Open and process image
                    img = Image.open(full_image_path)
                    
                    # 1. Generate thumbnail for Excel
                    thumb_filename = f"thumb_{i}_{os.path.basename(image_path)}"
                    thumb_path = os.path.join(thumbs_dir, thumb_filename)
                    _ = create_thumbnail(img, thumb_path)
                    
                    # 2. Generate crop for website
                    base_name = os.path.splitext(os.path.basename(image_path))[0]
                    crop_filename = f"{base_name}_dettaglio.jpg"
                    crop_path = os.path.join(crops_dir, crop_filename)
                    _ = crop_image(img, crop_path)
                    
                    # Create a normalized copy of the row
                    modified_row = normalize_row(row, column_mapping)
                    
                    # Update the FOTO DETTAGLIO field with the crop filename
                    foto_dettaglio_col = column_mapping["FOTO DETTAGLIO"]
                    modified_row[foto_dettaglio_col] = crop_filename
                    
                    # Store modified row data for Excel and CSV
                    valid_rows_data.append((modified_row, thumb_path))
                    
                except Exception as e:
                    if status_callback:
                        status_callback(f"Errore con immagine {image_path}: {str(e)}")
                    missing_images.append(f"{image_path} (errore: {str(e)})")
            else:
                missing_images.append(image_path)
        
        # Write valid rows to Excel with exactly the specified columns
        for i, (row, thumb_path) in enumerate(valid_rows_data):
            excel_row = i + 1  # +1 for header
            
            # Write only the required columns in the specified order
            for j, orig_col in enumerate(REQUIRED_COLUMNS):
                mapped_col = column_mapping[orig_col]
                worksheet.write(excel_row, j+1, row[mapped_col])
            
            # Insert larger thumbnail
            worksheet.insert_image(excel_row, 0, thumb_path, {'x_scale': 0.5, 'y_scale': 0.5})
        
        # Close Excel workbook
        workbook.close()
        
        # Create DataFrame with only valid rows and only required columns
        valid_df = pd.DataFrame([row for row, _ in valid_rows_data], columns=df.columns)
        
        # Generate CSV file
        if status_callback:
            status_callback("Genero il file CSV...")
            
        csv_result = generate_csv_output(valid_df, output_path, column_mapping)
        
        # Return results
        return True, {
            "excel_success": True,
            "crops_success": True,
            "csv_success": True,
            "processed_rows": len(valid_rows_data),
            "missing_images": len(missing_images),
            "total_rows": total_rows,
            "excel_path": excel_output_path,
            "csv_path": csv_output_path,
            "crops_dir": crops_dir
        }
        
    except Exception as e:
        if status_callback:
            status_callback(f"Errore durante l'elaborazione: {str(e)}")
        return False, {"error": str(e)}
    
def create_thumbnail(img, output_path, max_size=(500, 500), quality=70):
    """
    Create a thumbnail from an image and save it.
    
    Args:
        img: PIL Image object to thumbnail
        output_path: Path to save the thumbnail
        max_size: Maximum dimensions (width, height)
        quality: JPEG quality (0-100)
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Make sure the output directory exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # Create a copy to avoid modifying the original
        thumb_img = img.copy()
        # Rotate image if needed
        thumb_img = thumb_img.rotate(90, expand=True)
        # Resize to fit within max_size while preserving aspect ratio
        thumb_img.thumbnail(max_size, Image.LANCZOS)
        
        # Save the thumbnail
        thumb_img.save(output_path, optimize=True, quality=quality)
        
        return True
    except Exception as e:
        print(f"Error creating thumbnail: {str(e)}")
        return False

def crop_image(img, output_path, max_size=(1000, 1000), quality=80):
    """
    Crop an image according to specific parameters and save it.
    
    Args:
        img: PIL Image object to crop
        output_path: Path to save the cropped image
        max_size: Maximum dimensions after cropping
        quality: JPEG quality (0-100)
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Make sure the output directory exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # Apply crop
        
        crop_width = 3648/2-300
        crop_height = 4864/2-300
        downshift = -200
        img = img.rotate(90, expand=True)
        
        # Calculate crop coordinates
        x = (img.width - crop_width)//2
        y = (img.height - crop_height)//2
        crop_box = (x, y+downshift, x+crop_width, y+crop_height+downshift)
        
        # Apply the crop
        cropped_img = img.crop(crop_box)
        
        # Resize if needed
        cropped_img.thumbnail(max_size, Image.LANCZOS)
        
        # Save the cropped image
        cropped_img.save(output_path, format="JPEG", quality=quality)
        
        return True
    except Exception as e:
        print(f"Error cropping image: {str(e)}")
        return False

def normalize_row(row, column_mapping):
    """
    Normalize row data before writing to output files.
    
    Args:
        row: DataFrame row to normalize
        column_mapping: Mapping from canonical column names to actual column names
        
    Returns:
        Normalized copy of row
    """
    # Create a copy to avoid modifying the original
    normalized_row = row.copy()
    
    # Handle altezza and peso (extract and keep the smallest number)
    for field in ["ALTEZZA", "PESO"]:
        if field in column_mapping:
            col = column_mapping[field]
            if col in row and row[col] is not None and not pd.isna(row[col]):
                # Get the value as string
                value = str(row[col])
                
                # Try to extract all numbers from the string
                import re
                numbers = re.findall(r'\d+(?:\.\d+)?', value)
                
                # If numbers found, convert to float and keep the smallest
                if numbers:
                    try:
                        float_numbers = [float(num) for num in numbers]
                        smallest = min(float_numbers)
                        # Set the normalized value
                        normalized_row[col] = smallest
                    except ValueError:
                        # If conversion fails, keep original
                        pass
    
    # Add more normalization rules as needed
    
    return normalized_row