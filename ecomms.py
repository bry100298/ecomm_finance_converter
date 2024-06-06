import sys
import os
import subprocess
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QHBoxLayout, QWidget, QPushButton, QSizePolicy, QComboBox, QProgressBar, QProgressDialog, QFileDialog, QMessageBox
from PyQt5.QtGui import QPixmap, QIcon
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QMetaObject, Q_ARG
import time

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

store = ['Fritolay', 'Glico']
platform = ['Lazada', 'Shopee', 'Tiktok']
ROW_LIMIT = 999

class ScriptRunner(QThread):
    progress = pyqtSignal(int)

    def __init__(self, store_name, script, progress_per_script):
        super().__init__()
        self.store_name = store_name
        self.script = script
        self.progress_per_script = progress_per_script

    # def run(self):
    #     script_path = resource_path(os.path.join('ecomm_automation', 'functions', self.store_name, self.script))
    #     subprocess.run(['python', script_path])
    #     # subprocess.run(['python', os.path.join('ecomm_automation', 'functions', self.store_name, self.script)])
    #     self.progress.emit(self.progress_per_script)

    def run(self):
        script_path = resource_path(os.path.join('ecomm_automation', 'functions', self.store_name, self.script))
        if os.name == 'nt':  # Check if the OS is Windows
            # Use CREATE_NO_WINDOW on Windows to avoid showing the console window
            subprocess.run(['python', script_path], creationflags=subprocess.CREATE_NO_WINDOW)
        else:
            # On other platforms, redirect the output to avoid showing the console window
            subprocess.run(['python', script_path], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        self.progress.emit(self.progress_per_script)

class MainWindow(QMainWindow):
    def run_scripts(self):
        scripts = ['Lazada.py', 'Shopee.py', 'Tiktok.py']
        # scripts = ['Lazada.pyw', 'Shopee.pyw', 'Tiktok.pyw']
        total_scripts = len(scripts) * len(store)
        progress_per_script = int(100 / total_scripts)

        self.progress_dialog = QProgressDialog("Running scripts...", "Cancel", 0, 100, self)
        self.progress_dialog.setWindowModality(Qt.ApplicationModal)
        self.progress_dialog.setAutoClose(True)
        self.progress_dialog.setValue(0)
        self.progress_dialog.setWindowTitle("ECOMM")
        self.progress_dialog.setWindowIcon(self.windowIcon())

        self.completed_scripts = 0
        self.threads = []

        for store_name in store:
            for script in scripts:
                runner = ScriptRunner(store_name, script, progress_per_script)
                runner.progress.connect(self.update_run_progress)
                self.threads.append(runner)
                runner.start()

    def update_run_progress(self, value):
        self.completed_scripts += 1
        self.progress_dialog.setValue(self.completed_scripts * value)

        if self.completed_scripts == len(self.threads):
            self.progress_dialog.close()

    def handle_extract(self):
        selected_option = self.combo_box1.currentText()
        if selected_option == "QuickBooks":
            self.extract_quickbooks_files()
        elif selected_option == "Consolidation":
            self.extract_consolidation_files()

    def extract_quickbooks_files(self):
        target_dir = QFileDialog.getExistingDirectory(self, "Select Directory")
        if not target_dir:
            return

        self.progress_dialog = QProgressDialog("Extracting files...", "Cancel", 0, 100, self)
        self.progress_dialog.setWindowModality(Qt.ApplicationModal)
        self.progress_dialog.setAutoClose(True)
        self.progress_dialog.setValue(0)
        self.progress_dialog.setWindowTitle("ECOMM")
        self.progress_dialog.setWindowIcon(self.windowIcon())
        self.progress_dialog.show()

        total_tasks = 0
        for store_name in store:
            for platform_name in platform:
                quickbooks_dir = os.path.join(store_name, platform_name, 'Outbound', 'QuickBooks')
                if os.path.exists(quickbooks_dir):
                    for file_name in os.listdir(quickbooks_dir):
                        if file_name.endswith('.xlsx'):
                            df = pd.read_excel(os.path.join(quickbooks_dir, file_name))
                            total_tasks += (len(df) + ROW_LIMIT - 1) // ROW_LIMIT

        self.completed_tasks = 0

        for store_name in store:
            store_dir = os.path.join(target_dir, store_name)
            os.makedirs(store_dir, exist_ok=True)

            for platform_name in platform:
                quickbooks_dir = os.path.join(store_name, platform_name, 'Outbound', 'QuickBooks')
                if os.path.exists(quickbooks_dir):
                    for file_name in os.listdir(quickbooks_dir):
                        if file_name.endswith('.xlsx'):
                            file_path = os.path.join(quickbooks_dir, file_name)
                            df = pd.read_excel(file_path)
                            self.split_excel(df, store_dir, total_tasks)

        self.progress_dialog.close()

    def update_quickbooks_progress(self, total_tasks, num_rows):
        progress_per_task = 100 / total_tasks
        self.completed_tasks += (num_rows + ROW_LIMIT - 1) // ROW_LIMIT
        progress = self.completed_tasks * progress_per_task
        self.progress_dialog.setValue(int(progress))
        QApplication.processEvents()

    def split_excel(self, df, output_dir, total_tasks):
        num_splits = (len(df) + ROW_LIMIT - 1) // ROW_LIMIT  # Calculate the number of splits needed

        for i in range(num_splits):
            start_row = i * ROW_LIMIT
            end_row = min(start_row + ROW_LIMIT, len(df))
            split_df = df.iloc[start_row:end_row]

            timestamp = int(time.time())
            output_file = os.path.join(output_dir, f'{timestamp}_{i+1}.xlsx')
            split_df.to_excel(output_file, index=False)

            self.update_quickbooks_progress(total_tasks, end_row - start_row)

    # def process_store_files(self, store_name, target_folder):
    #     # input_folder = os.path.join('ecomm_automation', 'functions', store_name, 'Shopee', 'Outbound', 'QuickBooks')
    #     input_folder = os.path.join(store_name, 'Outbound', 'QuickBooks')
    #     output_folder = os.path.join(target_folder, store_name, 'QuickBooks')
    #     os.makedirs(output_folder, exist_ok=True)

    #     for filename in os.listdir(input_folder):
    #         if filename.endswith('.xlsx'):
    #             file_path = os.path.join(input_folder, filename)
    #             self.split_excel_file(file_path, output_folder)

    # def split_excel_file(self, file_path, output_folder):
    #     df = pd.read_excel(file_path)
    #     total_rows = len(df)
    #     chunks = (total_rows // 1000) + 1
    #     base_filename = os.path.splitext(os.path.basename(file_path))[0]

    #     for i in range(chunks):
    #         start_row = i * 1000
    #         end_row = (i + 1) * 1000
    #         chunk_df = df[start_row:end_row]
    #         output_file = os.path.join(output_folder, f"{base_filename}_part{i+1}.xlsx")
    #         chunk_df.to_excel(output_file, index=False)

    #     QMessageBox.information(self, "Success", f"Files extracted and saved to {output_folder}")

    def extract_consolidation_files(self):
        target_dir = QFileDialog.getExistingDirectory(self, "Select Directory")
        if not target_dir:
            return

        self.progress_dialog = QProgressDialog("Extracting files...", "Cancel", 0, 100, self)
        self.progress_dialog.setWindowModality(Qt.ApplicationModal)
        self.progress_dialog.setAutoClose(True)
        self.progress_dialog.setValue(0)
        self.progress_dialog.setWindowTitle("ECOMM")
        self.progress_dialog.setWindowIcon(self.windowIcon())
        self.progress_dialog.show()

        total_files = 0
        for store_name in store:
            for platform_name in platform:
                consolidation_dir = os.path.join(store_name, platform_name, 'Outbound', 'Consolidation')
                if os.path.exists(consolidation_dir):
                    total_files += len([file for file in os.listdir(consolidation_dir) if file.endswith('.xlsx')])

        self.completed_tasks = 0

        for store_name in store:
            store_dir = os.path.join(target_dir, store_name)
            os.makedirs(store_dir, exist_ok=True)

            writer = pd.ExcelWriter(os.path.join(store_dir, 'consolidation.xlsx'), engine='xlsxwriter')

            for platform_name in platform:
                platform_data = pd.DataFrame()
                consolidation_dir = os.path.join(store_name, platform_name, 'Outbound', 'Consolidation')
                if os.path.exists(consolidation_dir):
                    for file_name in os.listdir(consolidation_dir):
                        if file_name.endswith('.xlsx'):
                            file_path = os.path.join(consolidation_dir, file_name)
                            df = pd.read_excel(file_path)
                            platform_data = pd.concat([platform_data, df])
                            self.update_consolidation_progress(total_files)

                platform_data.to_excel(writer, sheet_name=platform_name, index=False)

            writer.close()

        self.progress_dialog.close()

    def update_consolidation_progress(self, total_files):
        progress_per_task = 100 / total_files
        self.completed_tasks += 1
        progress = self.completed_tasks * progress_per_task
        self.progress_dialog.setValue(int(progress))
        QApplication.processEvents()

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
        # icon_path = os.path.join(parent_dir, 'assets', 'benbytree_icon.ico')
        icon_path = resource_path(os.path.join(parent_dir, 'assets', 'benbytree_icon.ico'))
        self.setWindowIcon(QIcon(icon_path))

        # Define subdirectories
        frame0 = os.path.join(parent_dir, 'assets', 'frame0')
        fritolay = os.path.join(parent_dir, 'functions', 'Fritolay')
        glico = os.path.join(parent_dir, 'functions', 'Glico')

        # Load and display the background image
        background_label = QLabel(self)
        # pixmap = QPixmap(os.path.join(frame0, "image_1.png"))
        pixmap = QPixmap(resource_path(os.path.join(frame0, "image_1.png")))
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
            btn.setIcon(QIcon(resource_path(os.path.join(frame0, icon_file))))
            # btn.setIcon(QIcon(os.path.join(frame0, icon_file)))
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

        #Comment here
        # combo_box1 = QComboBox()
        # combo_box1.addItem("Placeholder 1")
        # combo_box1.setFixedWidth(150)

        self.combo_box1 = QComboBox()
        self.combo_box1.addItem("Please select extraction")
        self.combo_box1.addItem("Consolidation")
        self.combo_box1.addItem("QuickBooks")
        self.combo_box1.setFixedWidth(150)
        self.combo_box1.setStyleSheet("""
            QComboBox {
                border: 1px solid #6EC8C3;
                border-radius: 5px;
                padding: 5px;
                background-color: #FFFFFF;
            }
        """)

        # combo_box2 = QComboBox()
        # combo_box2.addItem("Placeholder 2")
        # combo_box2.setFixedWidth(150)

        #Comment here
        # extract_button = QPushButton("Extract")
        # extract_button.setFixedWidth(100)

        # extract_button = QPushButton("Extract")
        # extract_button.setFixedWidth(100)
        # extract_button.clicked.connect(self.extract_quickbooks_files)

        extract_button = QPushButton("Extract")
        extract_button.setFixedWidth(100)
        extract_button.setStyleSheet("""
            QPushButton {
                background-color: #FFFFFF;
                border: none;
                color: black;
                padding: 5px;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #d0bfff;
            }
            QPushButton:pressed {
                background-color: #b5b69c;
            }
        """)
        extract_button.clicked.connect(self.handle_extract)

        run_button = QPushButton("Run")
        run_button.setFixedWidth(100)
        run_button.setStyleSheet("""
            QPushButton {
                background-color: #FFFFFF;
                border: none;
                color: black;
                padding: 5px;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #d0bfff;
            }
            QPushButton:pressed {
                background-color: #b5b69c;
            }
        """)
        run_button.clicked.connect(self.run_scripts)

        #Comment here
        # footer_layout.addWidget(combo_box1)
        footer_layout.addWidget(self.combo_box1)
        # footer_layout.addWidget(combo_box2)
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
