import sys, time
from typing import Optional
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QScrollArea,
                             QVBoxLayout, QWidget, QHBoxLayout, QPushButton,
                             QLineEdit, QFileDialog, QTextEdit, QProgressBar)
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QPalette, QColor, QPixmap, QIcon, QFont

from PyQt5.QtCore import QSettings
settings = QSettings("buio", "GeneraExcelAnteprime")

# Import custom modules
from styles import *
from processor import *

class DropArea(QLabel):
    def __init__(self, placeholder: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.setText(placeholder)
        self.setStyleSheet(DROP_AREA_NORMAL)
        self.setAcceptDrops(True)
        self.setMinimumSize(300, 200)
        self.setWordWrap(True)
        self.file_path: Optional[str] = None

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()


class FileDropArea(DropArea):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__("Trascina qui il file excel", parent)


    def dropEvent(self, event: QDropEvent) -> None:
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            filename = urls[0].fileName()
            
            # Validate the Excel file
            is_valid, message = check_excel_file(file_path)
            if is_valid:
                # Parse the Excel file to get information
                success, info, parse_message = parse_excel_file(file_path)

                if success:
                    # Store the file path for later use
                    self.file_path = file_path

                    # Format information for display
                    elements_info = f"{filename}\n\nColonne: {len(info['column_names'])}"
                
                    # Check if we have information about removed rows
                    if "total_rows_removed" in info and info["total_rows_removed"] > 0:
                        # Show both before and after counts when rows were removed
                        elements_info += f"\nElementi validi: {info['rows_after_cleaning']}\n"
                        #elements_info += f"Elementi originali: {info['rows']}\n"
                        elements_info += f"Righe rimosse: {info['total_rows_removed']}"
                    else:
                        # Show only row count if no rows were removed
                        elements_info += f"\nElementi: {info['rows']}\n"
                        
                    # Update the display text with success and information
                    self.setText(elements_info)
                    self.setStyleSheet(DROP_AREA_SUCCESS)
                else:
                    # Update the display with parsing error
                    self.setText(f"Errore Parsing Excel:\n{parse_message}")
                    self.setStyleSheet(DROP_AREA_ERROR)
                    self.file_path = None
                    
            else:
                # Update the display with error
                self.setText(f"Errore:\n{message}")
                self.setStyleSheet(DROP_AREA_ERROR)
                self.file_path = None


class FolderDropArea(DropArea):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__("Trascina qui la cartella immagini", parent)

    def dropEvent(self, event: QDropEvent) -> None:
        urls = event.mimeData().urls()
        if urls:
            folder_path = urls[0].toLocalFile()
            print(f"DEBUG: Dropped path: {folder_path}")
            
            # Check if the path is actually a directory
            import os
            if not os.path.isdir(folder_path):
                print(f"DEBUG: Not a directory")
                self.setText("Errore:\nDevi trascinare una cartella, non un file")
                self.setStyleSheet(DROP_AREA_ERROR)
                self.file_path = None
                return
                
            foldername = os.path.basename(os.path.normpath(folder_path))
            # Extract folder name from path
            print(f"DEBUG: Folder name: {foldername}")
            
            # Check the folder for images
            print(f"DEBUG: Calling check_image_folder")
            success, info, message = check_image_folder(folder_path)
            print(f"DEBUG: check_image_folder results: success={success}, message={message}")
            
            if success:
                # Store the folder path for later use
                self.file_path = folder_path
                
                # Format the image type breakdown
                type_info = ""
                for ext, count in info["image_types"].items():
                    type_info += f"\n{ext}: {count}"
                print(f"DEBUG: Image types: {type_info}")

                # Update the display text with success and information
                display_text = f"Cartella:\n{foldername}\n\nImmagini: {info['total_images']}{type_info}"
                print(f"DEBUG: Setting display text: {display_text}")
                self.setText(display_text)
                self.setStyleSheet(DROP_AREA_SUCCESS)
            else:
                # Update the display with error
                print(f"DEBUG: Setting error text: {message}")
                # Update the display with error but include the folder name
                self.setText(f"Cartella:\n{foldername}\n\nErrore:\n{message}")
                self.setStyleSheet(DROP_AREA_ERROR)
                self.file_path = None

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.generate_button = QPushButton("Genera Output")
        self.setGeometry(100, 100, 650, 580)
        
        # Set app icon
        app_icon = QIcon("logo.png")
        self.setWindowIcon(app_icon)
        
        # Set window background color
        palette = self.palette()
        palette.setColor(QPalette.Window, QColor("#FFFFFF"))
        self.setPalette(palette)
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        #scroll_area = QScrollArea()
        #scroll_area.setWidgetResizable(True)
        #scroll_area.setWidget(central_widget)
        #self.setCentralWidget(scroll_area)
        
        # Create main layout
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(25, 25, 25, 25)
        main_layout.setSpacing(20)
        
        # Create header layout (logo + title)
        header_layout = QHBoxLayout()
        
        # Add logo to header
        self.logo_label = QLabel()
        logo_pixmap = QPixmap("logo.png")
        self.logo_label.setPixmap(logo_pixmap.scaled(50, 50, Qt.KeepAspectRatio))
        header_layout.addWidget(self.logo_label)
        
        # Add title to header
        self.title_label = QLabel("Genera Excel Anteprime")
        title_font = QFont()
        title_font.setPointSize(24)
        title_font.setBold(True)
        self.title_label.setFont(title_font)
        self.title_label.setStyleSheet(TITLE_TEXT)
        self.title_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        header_layout.addWidget(self.title_label)
        header_layout.setStretch(1, 1)  # Give more stretch to title
        
        # Add header to main layout
        main_layout.addLayout(header_layout)
        
        # Add description with bullet points
        self.description_label = QLabel("Trascina excel BOX e cartella immagini BOX per generare:")
        self.description_label.setWordWrap(True)
        self.description_label.setStyleSheet(DESCRIPTION_TEXT)
        main_layout.addWidget(self.description_label)
        
        self.formatted_info = QLabel("Excel anteprime • Cartella con crop per sito • CSV per caricamento sito")
        self.formatted_info.setAlignment(Qt.AlignLeft)
        self.formatted_info.setStyleSheet(FORMATTED_INFO)
        main_layout.addWidget(self.formatted_info)

        # Add the new information text
        self.info_label = QLabel("Supporta file .xls, .xlsx e .numbers. Cerca le immagini nella cartella principale, non controlla le sottocartelle")
        self.info_label.setWordWrap(True)
        self.info_label.setStyleSheet(INFO_TEXT)  # You'll need to add this style
        main_layout.addWidget(self.info_label)

        # Create horizontal layout for drop areas
        drop_layout = QHBoxLayout()
        drop_layout.setSpacing(20)
        
        # Create horizontal layout for drop areas
        drop_layout = QHBoxLayout()
        drop_layout.setSpacing(20)
        
        # Add file drop area
        self.file_drop_area = FileDropArea()
        drop_layout.addWidget(self.file_drop_area)
        
        # Add folder drop area
        self.folder_drop_area = FolderDropArea()
        drop_layout.addWidget(self.folder_drop_area)
        
        # Add drop layout to main layout
        main_layout.addLayout(drop_layout)

        # Create output path selector
        path_layout = QHBoxLayout()
        # Create output path selector
        path_layout = QHBoxLayout()
        
        # Add text field for path
        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText("Seleziona percorso di output...")
        self.output_path.setStyleSheet(PATH_INPUT)
        path_layout.addWidget(self.output_path)
        
        # Add browse button
        self.browse_button = QPushButton("Sfoglia...")
        self.browse_button.setStyleSheet(BROWSE_BUTTON)
        self.browse_button.setCursor(Qt.PointingHandCursor)
        path_layout.addWidget(self.browse_button)
        
        # Connect browse button
        self.browse_button.clicked.connect(self.select_output_path)
        
        # Add path layout to main layout
        main_layout.addLayout(path_layout)

        # Add generate button
        self.generate_button = QPushButton("Genera Excel Anteprime")
        self.generate_button.setStyleSheet(GENERATE_BUTTON)
        self.generate_button.setCursor(Qt.PointingHandCursor)
        self.generate_button.clicked.connect(self.generate_outputs)
        main_layout.addWidget(self.generate_button)
        
        # Add status text area for missing images
        self.status_text = QTextEdit()
        self.status_text.setReadOnly(True)
        self.status_text.setMaximumHeight(80)
        self.status_text.setPlaceholderText("Stato processamento...")
        self.status_text.setStyleSheet(STATUS_TEXT)
        main_layout.addWidget(self.status_text)

        # Add progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p% completato")
        self.progress_bar.setStyleSheet(PROGRESS_BAR)
        main_layout.addWidget(self.progress_bar)

        # Add output status indicators
        status_layout = QHBoxLayout()
        status_layout.setSpacing(10)

        # Excel indicator
        self.excel_status = QLabel("Excel: ⌛")
        self.excel_status.setStyleSheet(STATUS_INDICATOR_WAITING)
        status_layout.addWidget(self.excel_status)

        # Crops indicator
        self.crops_status = QLabel("Crops: ⌛")
        self.crops_status.setStyleSheet(STATUS_INDICATOR_WAITING)
        status_layout.addWidget(self.crops_status)

        # CSV indicator
        self.csv_status = QLabel("CSV: ⌛")
        self.csv_status.setStyleSheet(STATUS_INDICATOR_WAITING)
        status_layout.addWidget(self.csv_status)

        # Add status layout to main layout
        main_layout.addLayout(status_layout)
        
        
        # Add subtitle at the bottom
        self.subtitle_label = QLabel("Marco Buiani X Archivio tailor 2025")
        subtitle_font = QFont()
        subtitle_font.setPointSize(9)
        subtitle_font.setItalic(True)
        self.subtitle_label.setFont(subtitle_font)
        self.subtitle_label.setStyleSheet(FOOTER_TEXT)
        self.subtitle_label.setAlignment(Qt.AlignRight)
        main_layout.addWidget(self.subtitle_label)

    def select_output_path(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Seleziona Cartella di Output")
        if path:
            self.output_path.setText(path)

    def update_status_text(self, message: str) -> None:
        """Update the status text area with a message."""
        self.status_text.setText(message)

    def update_missing_images(self, missing_images: List[str], total_images: int) -> None:
        """Update the status text with missing image information."""
        if not missing_images:
            self.status_text.setText("Tutte le immagini trovate!")
            return
            
        missing_count = len(missing_images)
        message = f"Immagini mancanti: {missing_count}/{total_images}\n"
        
        # Add missing image filenames
        for img in missing_images:
            message += f"- {img}\n"
            
        self.status_text.setText(message)

    def update_output_status(self, output_type: str, success: bool) -> None:
        """Update the status of an output indicator."""
        if output_type == "excel":
            label = self.excel_status
            text = "Excel: "
        elif output_type == "crops":
            label = self.crops_status
            text = "Crops: "
        elif output_type == "csv":
            label = self.csv_status
            text = "CSV: "
        else:
            return
        
        if success:
            label.setText(text + "✓")
            label.setStyleSheet(STATUS_INDICATOR_SUCCESS)
        else:
            label.setText(text + "✗")
            label.setStyleSheet(STATUS_INDICATOR_ERROR)

    def reset_processing_ui(self) -> None:
        """Reset all processing UI elements to initial state."""
        self.progress_bar.setValue(0)
        self.status_text.clear()
        self.excel_status.setText("Excel: ⌛")
        self.excel_status.setStyleSheet(STATUS_INDICATOR_WAITING)
        self.crops_status.setText("Crops: ⌛")
        self.crops_status.setStyleSheet(STATUS_INDICATOR_WAITING)
        self.csv_status.setText("CSV: ⌛")
        self.csv_status.setStyleSheet(STATUS_INDICATOR_WAITING)

    def generate_outputs(self) -> None:
        """Generate all output files."""
        # Reset UI components first
        self.reset_processing_ui()

        # Get file paths
        excel_path = self.file_drop_area.file_path
        images_folder = self.folder_drop_area.file_path
        output_path = self.output_path.text()
        
        # Validate all inputs are provided
        if not excel_path:
            self.status_text.setText("Errore: File Excel non fornito.")
            self.status_text.setStyleSheet(STATUS_TEXT_ERROR)
            return
            
        if not images_folder:
            self.status_text.setText("Errore: Cartella immagini non fornita.")
            self.status_text.setStyleSheet(STATUS_TEXT_ERROR)
            return
            
        if not output_path:
            self.status_text.setText("Errore: Percorso di output non specificato.")
            self.status_text.setStyleSheet(STATUS_TEXT_ERROR)
            return
        
        # Define simple callbacks
        def update_progress(current, total):
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(current)
            QApplication.processEvents()
        
        def update_status(message):
            self.status_text.setText(message)
            QApplication.processEvents()
        
        # Process files
        success, results = process_files(excel_path, images_folder, output_path, 
                                    update_progress, update_status)
        
        # Update UI based on results
        if success:
            # Update status indicators
            self.update_output_status("excel", results.get("excel_success", False))
            self.update_output_status("crops", results.get("crops_success", False))
            self.update_output_status("csv", results.get("csv_success", False))
            
            # Update status text
            self.status_text.setText(f"Elaborazione completata!\n"
                                    f"Righe processate: {results.get('processed_rows', 0)}\n"
                                    f"Immagini mancanti: {results.get('missing_images', 0)}")
        else:
            # Show error
            self.status_text.setText(f"Errore: {results.get('error', 'Errore sconosciuto')}")
            self.status_text.setStyleSheet(STATUS_TEXT_ERROR)
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("logo.png"))  # Set the app icon for taskbar/dock
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())