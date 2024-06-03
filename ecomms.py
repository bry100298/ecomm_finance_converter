import sys
import os
import subprocess
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QHBoxLayout, QWidget, QPushButton, QSizePolicy, QComboBox, QProgressBar, QProgressDialog
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import Qt, QThread, pyqtSignal

class ScriptRunner(QThread):
    progress = pyqtSignal(int)

    def __init__(self, script, progress_per_script):
        super().__init__()
        self.script = script
        self.progress_per_script = progress_per_script

    def run(self):
        subprocess.run(['python', os.path.join('ecomm_automation', 'functions', 'Fritolay', self.script)])
        self.progress.emit(self.progress_per_script)

class MainWindow(QMainWindow):
    def run_scripts(self):
        scripts = ['Lazada.py', 'Shopee.py', 'Tiktok.py']
        total_scripts = len(scripts)
        progress_per_script = int(100 / total_scripts)

        self.progress_dialog = QProgressDialog("Running scripts...", "Cancel", 0, 100, self)
        self.progress_dialog.setWindowModality(Qt.ApplicationModal)
        self.progress_dialog.setAutoClose(True)
        self.progress_dialog.setValue(0)
        self.progress_dialog.setWindowTitle("ECOMM")
        self.progress_dialog.setWindowIcon(self.windowIcon())

        self.completed_scripts = 0
        self.threads = []

        for script in scripts:
            runner = ScriptRunner(script, progress_per_script)
            runner.progress.connect(self.update_progress)
            self.threads.append(runner)
            runner.start()

    def update_progress(self, value):
        self.completed_scripts += 1
        self.progress_dialog.setValue(self.completed_scripts * value)

        if self.completed_scripts == len(self.threads):
            self.progress_dialog.close()

    def __init__(self):
        super(MainWindow, self).__init__()

        self.completed_scripts = 0

        # Set window properties
        self.setWindowTitle('ECOMMS AUTOMATION SYSTEM')
        self.setGeometry(100, 100, 750, 550)
        self.setFixedSize(750, 550)  # This will make the window not resizable

        # Define parent directory
        parent_dir = 'ecomm_automation'
        # Load and set the window icon
        icon_path = os.path.join(parent_dir, 'assets', 'benbytree_icon.ico')
        self.setWindowIcon(QIcon(icon_path))

        # Define subdirectories
        frame0 = os.path.join(parent_dir, 'assets', 'frame0')
        fritolay = os.path.join(parent_dir, 'functions', 'Fritolay')
        glico = os.path.join(parent_dir, 'functions', 'Glico')

        # Load and display the background image
        background_label = QLabel(self)
        pixmap = QPixmap(os.path.join(frame0, "image_1.png"))
        background_label.setPixmap(pixmap)
        background_label.setScaledContents(True)
        background_label.setGeometry(0, 0, 750, 550)

        # Create a header
        header = QLabel("ECOMMS AUTOMATION SYSTEM", self)
        header.setAlignment(Qt.AlignCenter)
        header.setStyleSheet("background-color: #6EC8C3; color: #1D4916; font-size: 24px; padding: 10px;")
        header.setFixedHeight(71)
        header.setGeometry(0, 0, 750, 71)

        # Create a sidebar with sections and buttons
        sidebar_layout = QVBoxLayout()
        sidebar_layout.setContentsMargins(0, 0, 0, 0)  # Set layout margins to zero
        sidebar_layout.setSpacing(0)  # Set spacing to zero

        buttons = [("Home", "home.png"), ("Raw Data", "raw_data_icon.png"), ("SKU", "sku_icon.png"), ("Orders Report", "orders_report_icon.png")]
        self.sidebar_buttons = []
        for button_text, icon_file in buttons:
            btn = QPushButton(button_text)
            btn.setIcon(QIcon(os.path.join(frame0, icon_file)))
            btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            btn.setFixedHeight(50)
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #FFFFFF; 
                    padding: 10px; 
                    border: none;
                    text-align: left;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #e0e0e0;
                }
                QPushButton:pressed {
                    background-color: #d0d0d0;
                    border-left: 5px solid #6EC8C3;
                }
            """)
            btn.clicked.connect(lambda checked, b=btn: self.highlight_button(b))
            sidebar_layout.addWidget(btn)
            self.sidebar_buttons.append(btn)

        sidebar_widget = QWidget()
        sidebar_widget.setLayout(sidebar_layout)
        sidebar_widget.setFixedWidth(195)
        sidebar_widget.setStyleSheet("background-color: white; border-right: 2px solid #ccc;")
        # sidebar_widget.setGeometry(0, 71, 195, 409)  # Positioned below the header and above the footer

        # Create the main content area
        main_content_layout = QVBoxLayout()
        main_content_layout.setContentsMargins(0, 0, 0, 0)  # Set layout margins to zero

        main_content_widget = QWidget()
        main_content_widget.setLayout(main_content_layout)
        main_content_widget.setStyleSheet("background-color: transparent;")

        # Create a footer with dropdowns and buttons
        footer_layout = QHBoxLayout()
        footer_layout.setContentsMargins(10, 5, 10, 5)
        footer_layout.setSpacing(10)

        combo_box1 = QComboBox()
        combo_box1.addItem("Placeholder 1")
        combo_box1.setFixedWidth(150)

        combo_box2 = QComboBox()
        combo_box2.addItem("Placeholder 2")
        combo_box2.setFixedWidth(150)

        extract_button = QPushButton("Extract")
        extract_button.setFixedWidth(100)

        run_button = QPushButton("Run")
        run_button.setFixedWidth(100)
        run_button.clicked.connect(self.run_scripts)

        footer_layout.addWidget(combo_box1)
        footer_layout.addWidget(combo_box2)
        footer_layout.addStretch()
        footer_layout.addWidget(extract_button)
        footer_layout.addWidget(run_button)

        footer_widget = QWidget()
        footer_widget.setLayout(footer_layout)
        footer_widget.setStyleSheet("background-color: #6EC8C3;")
        footer_widget.setFixedHeight(40)
        footer_widget.setGeometry(0, 510, 750, 40)

        # Combine sidebar and main content
        content_layout = QHBoxLayout()
        content_layout.setContentsMargins(0, 0, 0, 0)  # Set layout margins to zero
        content_layout.setSpacing(0)  # Set spacing to zero
        content_layout.addWidget(sidebar_widget)
        content_layout.addWidget(main_content_widget)

        content_widget = QWidget()
        content_widget.setLayout(content_layout)

        # Create the main layout
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0, 0, 0, 0)  # Set layout margins to zero
        main_layout.setSpacing(0)  # Set spacing to zero
        main_layout.addWidget(header)
        main_layout.addWidget(content_widget)
        main_layout.addWidget(footer_widget)

        # Create a central widget to hold the main layout
        central_widget = QWidget()
        central_widget.setLayout(main_layout)

        # Set the central widget of the main window
        self.setCentralWidget(central_widget)

        # Adjust background image stacking order
        background_label.lower()
        central_widget.raise_()

    def highlight_button(self, button):
        for btn in self.sidebar_buttons:
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #FFFFFF; 
                    padding: 10px; 
                    border: none;
                    text-align: left;
                    font-size: 14px;
                }
                QPushButton:hover {
                    background-color: #e0e0e0;
                }
            """)
        button.setStyleSheet("""
            QPushButton {
                background-color: #d0d0d0;
                border-left: 5px solid #6EC8C3;
                padding: 10px;
                text-align: left;
                font-size: 14px;
            }
        """)

def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
