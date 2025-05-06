import sys
from typing import Optional
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, 
                             QVBoxLayout, QWidget, QHBoxLayout, QPushButton,
                             QLineEdit, QFileDialog)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QPalette, QColor, QPixmap, QIcon, QFont

from processor import check_excel_file

class DropArea(QLabel):
    def __init__(self, placeholder: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.setText(placeholder)
        self.setStyleSheet("""
            background-color: #F7F7FF;
            border: 2px dashed #B8C4FF;
            border-radius: 8px;
            padding: 20px;
            font-size: 18px;
            color: #666677;
        """)
        self.setAcceptDrops(True)
        self.setMinimumSize(300, 200)
        self.setWordWrap(True)

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
                # Store the file path for later use
                self.file_path = file_path
                # Update the display text with success
                self.setText(f"File:\n{filename}")
                self.setStyleSheet("""
                    background-color: #F0FFF0;  /* Light green background */
                    border: 2px dashed #90EE90;
                    border-radius: 8px;
                    padding: 20px;
                    font-size: 18px;
                    color: #666677;
                """)
            else:
                # Update the display with error
                self.setText(f"Errore:\n{message}")
                self.setStyleSheet("""
                    background-color: #FFF0F0;  /* Light red background */
                    border: 2px dashed #FFB6C1;
                    border-radius: 8px;
                    padding: 20px;
                    font-size: 18px;
                    color: #666677;
                """)


class FolderDropArea(DropArea):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__("Trascina qui la cartella immagini", parent)

    def dropEvent(self, event: QDropEvent) -> None:
        urls = event.mimeData().urls()
        if urls:
            foldername = urls[0].fileName()
            self.setText(f"Cartella:\n{foldername}")


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
        self.title_label.setStyleSheet("color: #444455;")
        self.title_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        header_layout.addWidget(self.title_label)
        header_layout.setStretch(1, 1)  # Give more stretch to title
        
        # Add header to main layout
        main_layout.addLayout(header_layout)
        
        # Add description with bullet points
        self.description_label = QLabel("Trascina excel BOX e cartella immagini per generare:")
        self.description_label.setWordWrap(True)
        self.description_label.setStyleSheet("""
            color: #666677;
            font-size: 14px;
            margin-bottom: 5px;
        """)
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
            bullet_label.setStyleSheet("""
                color: #666677;
                font-size: 14px;
            """)
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
        self.output_path.setStyleSheet("""
            border: 1px solid #B8C4FF;
            border-radius: 4px;
            padding: 8px;
            font-size: 14px;
        """)
        path_layout.addWidget(self.output_path)
        
        # Add browse button
        self.browse_button = QPushButton("Sfoglia...")
        self.browse_button.setStyleSheet("""
            background-color: #E0E4FF;
            color: #444455;
            border: none;
            border-radius: 4px;
            padding: 8px 15px;
            font-size: 14px;
        """)
        self.browse_button.setCursor(Qt.PointingHandCursor)
        path_layout.addWidget(self.browse_button)
        
        # Connect browse button
        self.browse_button.clicked.connect(self.select_output_path)
        
        # Add path layout to main layout
        main_layout.addLayout(path_layout)
        
        # Add button
        self.generate_button = QPushButton("Genera Excel Anteprime")
        self.generate_button.setStyleSheet("""
            background-color: #B8C4FF;
            color: #444455;
            border: none;
            border-radius: 4px;
            padding: 10px 20px;
            font-size: 14px;
            font-weight: bold;
        """)
        self.generate_button.setCursor(Qt.PointingHandCursor)
        self.generate_button.clicked.connect(self.generate_excel)
        main_layout.addWidget(self.generate_button)
        
        # Add subtitle at the bottom
        self.subtitle_label = QLabel("Marco Buiani X Archivio tailor 2025")
        subtitle_font = QFont()
        subtitle_font.setPointSize(9)
        subtitle_font.setItalic(True)
        self.subtitle_label.setFont(subtitle_font)
        self.subtitle_label.setStyleSheet("color: #888899;")
        self.subtitle_label.setAlignment(Qt.AlignRight)
        main_layout.addWidget(self.subtitle_label)

    def select_output_path(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Seleziona Cartella di Output")
        if path:
            self.output_path.setText(path)
    
    def generate_excel(self) -> None:
        # This function will be called when the button is clicked
        print("Generate Excel button clicked")
        # Add your Excel generation logic here


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("logo.png"))  # Set the app icon for taskbar/dock
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())