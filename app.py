import sys
from typing import Optional
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, 
                             QVBoxLayout, QWidget, QHBoxLayout, QPushButton,
                             QLineEdit, QFileDialog)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QPalette, QColor, QPixmap, QIcon, QFont

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
                    elements_info = f"File:\n{filename}\n\nElementi: {info['rows']}\nColonne: {len(info['column_names'])}"
                
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
            foldername = urls[0].fileName()
            
            
            # Validate the folder
            is_valid, message = check_image_folder(folder_path)
            print("ciao")
            if is_valid:
                # Analyze the folder to count images
                success, info, analyze_message = analyze_image_folder(folder_path)
            
                if success:
                    # Store the folder path for later use
                    self.file_path = folder_path
                    
                    # Format the image type breakdown
                    type_info = ""
                    for ext, count in info["image_types"].items():
                        type_info += f"\n{ext}: {count}"

                    # Update the display text with success and information
                    self.setText(f"Cartella:\n{foldername}\n\nImmagini: {info['total_images']}{type_info}")
                    self.setStyleSheet(DROP_AREA_SUCCESS)
                else:
                    # Update the display with analysis error
                    self.setText(f"Errore:\n{analyze_message}")
                    self.setStyleSheet(DROP_AREA_ERROR)
                    self.file_path = None

            else:
                # Update the display with error
                self.setText(f"Errore:\n{message}")
                self.setStyleSheet(DROP_AREA_ERROR)
                self.file_path = None


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Genera Excel Anteprime")
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
        
        # Create main layout
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(25, 25, 25, 25)
        main_layout.setSpacing(20)
        
        # Create header layout (logo + title)
        header_layout = QHBoxLayout()
        
        # Add logo to header
        self.logo_label = QLabel()
        logo_pixmap = QPixmap("logo.png")
        self.logo_label.setPixmap(logo_pixmap.scaled(80, 80, Qt.KeepAspectRatio))
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
        self.description_label = QLabel("Trascina excel BOX e cartella immagini per generare:")
        self.description_label.setWordWrap(True)
        self.description_label.setStyleSheet(DESCRIPTION_TEXT)
        main_layout.addWidget(self.description_label)
        
        # Add bullet points
        bullet_point_layout = QVBoxLayout()
        bullet_point_layout.setSpacing(2)
        bullet_point_layout.setContentsMargins(20, 0, 0, 10)
        
        bullet_points = [
            "• Excel anteprime",
            "• Cartella con crop per sito",
            "• CSV per caricamento sito"
        ]
        
        for point in bullet_points:
            bullet_label = QLabel(point)
            bullet_label.setStyleSheet(BULLET_POINT)
            bullet_point_layout.addWidget(bullet_label)
        
        main_layout.addLayout(bullet_point_layout)
        
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
        
        # Add button
        self.generate_button = QPushButton("Genera Excel Anteprime")
        self.generate_button.setStyleSheet(GENERATE_BUTTON)
        self.generate_button.setCursor(Qt.PointingHandCursor)
        self.generate_button.clicked.connect(self.generate_excel)
        main_layout.addWidget(self.generate_button)
        
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
    
    def generate_excel(self) -> None:
        # Get file paths
        excel_path = self.file_drop_area.file_path
        images_folder = self.folder_drop_area.file_path
        output_path = self.output_path.text()
        
        # Validate all inputs are provided
        if not excel_path:
            print("Excel file not provided")
            return
            
        if not images_folder:
            print("Images folder not provided")
            return
            
        if not output_path:
            print("Output path not provided")
            return
        
        print(f"Generating Excel with:")
        print(f"Excel file: {excel_path}")
        print(f"Images folder: {images_folder}")
        print(f"Output path: {output_path}")
        
        # Call processing function here
        # process_files(excel_path, images_folder, output_path)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("logo.png"))  # Set the app icon for taskbar/dock
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())