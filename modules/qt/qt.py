import sys
import os
import shutil
from typing import Dict, List, Optional

from PyQt5.QtWidgets import (QMainWindow, QLabel, QVBoxLayout, QWidget,
                             QHBoxLayout, QTabWidget, QFrame, QPushButton,
                             QScrollArea, QDesktopWidget, QFileDialog,
                             QInputDialog, QApplication, QProgressBar,
                             QStackedWidget)

from PyQt5.QtCore import (QThread, pyqtSignal, Qt, QFileSystemWatcher,
                          QStandardPaths)

from PyQt5.QtGui import QPainter, QColor, QBrush

from modules.condition.schemas import CellsConditionReport, CellsConditionState
from modules.params.schemas import AppParams
from modules.performances.time_counter import time_execution
from modules.report import report_instance
from modules.checker.checker import Checker
from modules.condition import CELLS_CONDITIONS
from modules.report.schemas import UIReportCell
from modules.excel import excel_handler
from modules.utils.utils import get_current_date_hour

from config.logger_config import logger

checker = Checker(cells_conditions=CELLS_CONDITIONS)


class Worker(QThread):
    """Worker thread to perform background updates."""
    update_completed = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)

    @time_execution
    def run(self):
        """Run the update process in a separate thread."""
        try:
            report_ui = self.update_state()
            self.update_completed.emit(report_ui)
        except Exception as e:
            self.error_occurred.emit(str(e))

    def update_state(self) -> Dict[str, List[UIReportCell]]:
        """Update the state by checking cell conditions.

        Returns:
            Dict[str, List[UIReportCell]]: The report data for the UI.
        """

        excel_handler.load_openpyxl(
            excel_abs_path=excel_handler.excel_abs_path)

        cells_condition_report: List[
            CellsConditionReport] = checker.check_cells_conditions()

        report_instance.cells_condition_report = cells_condition_report

        report_ui: Dict[str, List[UIReportCell]] = report_instance.get_report()

        return report_ui


class ColoredCircle(QWidget):
    """Widget to display a colored circle."""

    def __init__(self, color: QColor, parent=None):
        """Initialize the colored circle widget.

        Args:
            color (QColor): The color of the circle.
            parent (QWidget, optional): The parent widget. Defaults to None.
        """
        super().__init__(parent)
        self.color = color
        self.setFixedSize(20, 20)

    def paintEvent(self, event):
        """Handle the paint event to draw the circle.

        Args:
            event (QPaintEvent): The paint event.
        """
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        brush = QBrush(self.color)
        painter.setBrush(brush)
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(0, 0, 20, 20)


class MainWindow(QMainWindow):
    """Main window of the PyQt application."""

    def __init__(self):
        """Initialize the main window."""
        super().__init__()

        # Set up window properties
        self.setWindowTitle("Pronéo")
        # Use the standard window frame with title bar

        self.resize(900, 700)
        self.setMinimumSize(800, 600)

        self.file_watcher = None
        self.scroll_positions = {}
        self.worker = Worker()
        self.loading_indicator = None
        self.original_scroll_content = None  # To store the original content
        self.stack = QStackedWidget()
        self.placeholder_label = None

        self.setup_ui()
        self.position_window()

        self.connect_signals()

        self.error_stack_status = ""

    def setup_ui(self):
        """Set up the UI components."""
        self.main_widget = QWidget()
        self.main_layout = QVBoxLayout(self.main_widget)
        self.main_layout.setContentsMargins(10, 10, 10, 10)
        self.main_layout.setSpacing(10)

        self.setup_status_labels()
        self.setup_buttons()
        self.setup_placeholder()
        self.setup_loading_indicator()  # Set up the loading indicator widget
        self.setup_tab_widget()

        # Add placeholder, loading indicator, and tab widget to the stack
        self.stack.addWidget(self.placeholder_label)
        self.stack.addWidget(
            self.loading_widget)  # Add loading widget to the stack
        self.stack.addWidget(self.tab_widget)

        # Initially show the placeholder
        self.stack.setCurrentWidget(self.placeholder_label)

        # Add the stack to the main layout
        self.main_layout.addWidget(self.stack)

        self.setCentralWidget(self.main_widget)

        # Apply styles
        self.apply_styles()

    def setup_loading_indicator(self):
        """Set up the loading indicator widget."""
        self.loading_indicator = QProgressBar()
        self.loading_indicator.setRange(0, 0)  # Indeterminate mode
        self.loading_indicator.setTextVisible(False)
        self.loading_indicator.setFixedHeight(20)
        self.loading_indicator.setStyleSheet("""
            QProgressBar {
                background-color: #F0F0F0;
                border: none;
                border-radius: 10px;
            }
            QProgressBar::chunk {
                background-color: #007ACC;
                border-radius: 10px;
            }
        """)

        self.loading_widget = QWidget()
        layout = QVBoxLayout(self.loading_widget)
        layout.addStretch()
        layout.addWidget(self.loading_indicator, alignment=Qt.AlignCenter)
        layout.addStretch()

    def setup_placeholder(self):
        """Set up the placeholder label displayed before loading a file."""
        self.placeholder_label = QLabel("Charger ou créer un fichier.")
        self.placeholder_label.setAlignment(Qt.AlignCenter)
        self.placeholder_label.setStyleSheet(
            "font-size: 40px; color: #2b292a;")

    def setup_status_labels(self):
        """Set up the status and reminder labels."""
        self.status = QLabel("Status: OK")
        self.status.setAlignment(Qt.AlignCenter)
        self.status.setStyleSheet("color: black; font-size: 14px;")

        self.reminder_label = QLabel(
            "⚠️ Sauvegardez le fichier pour que toutes les modifications soient prises en compte. ⚠️"
        )
        self.reminder_label.setAlignment(Qt.AlignCenter)
        self.reminder_label.setStyleSheet(
            "color: red; font-size: 16px; font-weight: bold;")

        self.file_path_label = QLabel()
        self.file_path_label.setAlignment(Qt.AlignCenter)
        self.file_path_label.setStyleSheet("color: black; font-size: 12px;")

        self.main_layout.addWidget(self.status)
        self.main_layout.addWidget(self.reminder_label)
        self.main_layout.addWidget(self.file_path_label)

    def setup_buttons(self):
        """Set up the load, new file, and check buttons."""
        # Create horizontal layout for the buttons
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(20, 10, 20, 10)
        button_layout.setSpacing(20)

        # Load Excel button
        self.load_button = QPushButton("Charger Excel")
        self.load_button.setFixedSize(140, 40)
        self.load_button.clicked.connect(self.load_excel_file)

        # New Excel button
        self.new_file_button = QPushButton("Nouveau Excel")
        self.new_file_button.setFixedSize(140, 40)
        self.new_file_button.clicked.connect(self.create_new_excel)

        # Check button
        self.check_button = QPushButton("Vérifier")
        self.check_button.setFixedSize(160, 50)
        self.check_button.setStyleSheet("font-size: 18px; font-weight: bold;")
        self.check_button.clicked.connect(self.manual_check)

        # Add widgets to the button layout
        button_layout.addStretch()
        button_layout.addWidget(self.load_button)
        button_layout.addWidget(self.new_file_button)
        button_layout.addWidget(self.check_button)
        button_layout.addStretch()

        # Add the button layout to the main layout
        self.main_layout.addLayout(button_layout)

    def setup_tab_widget(self):
        """Set up the tab widget."""
        self.tab_widget = QTabWidget(self)
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: none;
                background-color: #FFFFFF;
            }
            QTabBar::tab {
                background: #E0E0E0;
                color: #000000;
                padding: 10px;
                min-width: 100px;
            }
            QTabBar::tab:selected {
                background: #FFFFFF;
                color: #000000;
            }
        """)

    def connect_signals(self):
        """Connect signals and slots."""
        self.worker.update_completed.connect(self.on_update_completed)
        self.worker.error_occurred.connect(self.on_worker_error)

    def on_worker_started(self):
        """Handle the event when the worker starts."""
        logger.info("Worker started.")
        if self.stack.currentWidget() == self.tab_widget:
            self.show_wait_indicator_in_current_tab()
        # If we're on the loading widget, no need to show the wait indicator in the tab

    def on_worker_finished(self):
        """Handle the event when the worker finishes."""
        logger.info("Worker finished.")
        if self.stack.currentWidget() == self.loading_widget:
            # Switch to the tab widget
            self.stack.setCurrentWidget(self.tab_widget)
        else:
            self.hide_wait_indicator_in_current_tab()

    def on_worker_error(self, error_message):
        """Handle errors from the worker thread.

        Args:
            error_message (str): The error message from the worker.
        """

        self.error_stack_status += "\n" + error_message

        logger.error("QT => Worker error: %s", error_message)

        self.status.setText(f"Status: {self.error_stack_status}")

    def start_worker(self):
        """Start the worker thread to perform the update in the background."""
        if not self.worker.isRunning():
            self.worker.start()

    def on_update_completed(self, report_ui: dict):
        """Handle the completion of the update and refresh the tabs.

        Args:
            report_ui (dict): The report data to update the UI.
        """
        try:
            self.on_worker_finished()
            self.refresh_tabs(report_ui)
            # self.status.setText("Status: OK")
        except Exception as e:
            self.error_stack_status += "\n" + str(e)

            self.status.setText(f"Status: {str(self.error_stack_status)}")

            logger.error("QT => Failed to refresh tabs: %s", str(e))

    def create_new_excel(self):
        """Create a new Excel file by copying an existing one."""
        try:
            new_file_name, ok = QInputDialog.getText(
                self, "New Excel (xlsm)",
                "Enter the file name (without extension):")

            if ok and new_file_name:
                new_file_name = f"{new_file_name}.xlsm"

                source_file_path = AppParams().excel_abs_path

                desktop_path = QStandardPaths.writableLocation(
                    QStandardPaths.DesktopLocation)

                destination_file_path = os.path.join(desktop_path,
                                                     new_file_name)

                shutil.copy(source_file_path, destination_file_path)

                self.file_path_label.setText(
                    f"Configuring file {destination_file_path} ...")

                excel_handler.load_excel(excel_abs_path=destination_file_path)

                self.setup_file_watcher()

                self.stack.setCurrentWidget(self.loading_widget)

                self.start_worker()

        except Exception as e:
            self.error_stack_status += "\n" + "Error creating new file => " + str(
                e)

            self.status.setText(f"Status: {str(self.error_stack_status)}")

            logger.error("QT => Failed to create new Excel: %s", str(e))

    def load_excel_file(self):
        """Open a file dialog to load an Excel file."""
        try:
            options = QFileDialog.Options()

            options |= QFileDialog.ReadOnly

            file_name, _ = QFileDialog.getOpenFileName(
                self,
                "Load Excel File",
                "",
                "Excel Files (*.xlsx *.xlsm)",
                options=options,
            )

            if file_name:
                excel_handler.load_excel(excel_abs_path=file_name)

                self.file_path_label.setText(
                    f"Configuring file {file_name} ...")

                self.setup_file_watcher()

                self.stack.setCurrentWidget(self.loading_widget)

                self.start_worker()

        except Exception as e:

            self.error_stack_status += "\n" + "Error loading file => " + str(e)

            self.status.setText(f"{str(self.error_stack_status)}")

            logger.error("QT => Failed to load new Excel: %s", str(e))

    def setup_file_watcher(self):
        """Set up a file watcher to monitor changes to the Excel file."""
        if self.file_watcher:
            # Remove existing watcher
            self.file_watcher.fileChanged.disconnect(
                self.on_excel_file_changed)
            self.file_watcher.removePath(excel_handler.excel_abs_path)
            self.file_watcher.deleteLater()
            self.file_watcher = None

        self.file_watcher = QFileSystemWatcher()
        self.file_watcher.addPath(excel_handler.excel_abs_path)
        self.file_watcher.fileChanged.connect(self.on_excel_file_changed)

    def on_excel_file_changed(self, path):
        """Handle the event when the Excel file is modified.

        Args:
            path (str): The path to the modified Excel file.
        """
        logger.info("Excel file %s has been modified.", path)

        self.on_worker_started()

        self.start_worker()

    def manual_check(self):
        """Perform a manual check by running the worker function."""

        try:
            if not self.worker.isRunning():

                self.on_worker_started()

                logger.info("Manual check initiated by user.")

                self.start_worker()

            else:
                logger.info("Worker is already running. Manual check ignored.")

        except Exception as e:
            self.error_stack_status += "\n" + str(e)

            self.status.setText(f"Status: {str(self.error_stack_status)}")

            logger.error("QT => Failed to manual check : %s", str(e))

    def position_window(self):
        """Position the window at the center of the screen."""
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def refresh_tabs(self, report_ui: Dict[str, List[UIReportCell]]) -> None:
        """Update the tabs with the current reports while preserving the selected tab and scroll position."""

        current_index = self.tab_widget.currentIndex()
        current_tab_name = self.tab_widget.tabText(
            current_index) if current_index >= 0 else None

        # Save scroll positions
        for i in range(self.tab_widget.count()):
            tab_name = self.tab_widget.tabText(i)
            scroll_area = self.tab_widget.widget(i).findChild(QScrollArea)
            if scroll_area:
                self.scroll_positions[
                    tab_name] = scroll_area.verticalScrollBar().value()

        self.tab_widget.clear()

        for sheet_name, report_cells in report_ui.items():

            tab = self.create_tab(sheet_name, report_cells)
            self.tab_widget.addTab(tab, sheet_name)

        # Restore selected tab and scroll positions
        if current_tab_name:
            for i in range(self.tab_widget.count()):
                if self.tab_widget.tabText(i) == current_tab_name:
                    self.tab_widget.setCurrentIndex(i)
                    break

        for i in range(self.tab_widget.count()):
            tab_name = self.tab_widget.tabText(i)
            scroll_area = self.tab_widget.widget(i).findChild(QScrollArea)
            if scroll_area and tab_name in self.scroll_positions:
                scroll_area.verticalScrollBar().setValue(
                    self.scroll_positions[tab_name])

        if excel_handler.excel_abs_path:
            self.file_path_label.setText(
                f"{excel_handler.excel_abs_path} vérifié le {get_current_date_hour().lower()}"
            )

    def create_tab(self, sheet_name: str,
                   report_cells: List[UIReportCell]) -> QWidget:
        """Create a tab with report cells for a given sheet."""

        tab = QWidget()
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("border: none;")
        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        for cell in report_cells:
            frame = self.create_report_cell_frame(cell)
            layout.addWidget(frame)

        layout.addStretch()
        scroll_content.setLayout(layout)
        scroll_area.setWidget(scroll_content)
        tab_layout = QVBoxLayout()
        tab_layout.setContentsMargins(0, 0, 0, 0)
        tab_layout.addWidget(scroll_area)
        tab.setLayout(tab_layout)
        return tab

    def create_report_cell_frame(self, cell: UIReportCell) -> QFrame:
        """Create a frame for a single report cell.

        Args:
            cell (UIReportCell): The report cell data.

        Returns:
            QFrame: The frame widget containing the report cell.
        """
        frame = QFrame()
        frame.setStyleSheet("""
            QFrame {
                background-color: #F0F0F0;
                border-radius: 5px;
            }
        """)
        h_layout = QHBoxLayout()
        h_layout.setContentsMargins(10, 10, 10, 10)
        h_layout.setSpacing(10)

        color = QColor(
            '#00C853') if cell.state == CellsConditionState.OK else QColor(
                '#D50000')
        circle = ColoredCircle(color)
        label = QLabel(cell.instruction)
        label.setWordWrap(True)
        label.setStyleSheet("color: #000000; font-size: 14px;")

        focus_button = QPushButton("Go to")
        focus_button.setFixedSize(80, 30)
        focus_button.clicked.connect(lambda _, sn=cell.sheet_names[0], ca=cell.
                                     cell_adress: self.focus_on_cell(sn, ca))
        focus_button.setStyleSheet("""
            QPushButton {
                background-color: #007ACC;
                color: white;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #005F9E;
            }
        """)

        h_layout.addWidget(circle)
        h_layout.addWidget(label, 1)
        h_layout.addWidget(focus_button)
        frame.setLayout(h_layout)
        return frame

    def focus_on_cell(self, sheet_name: str,
                      cell_address: Optional[str]) -> None:
        """Navigate to a specific sheet and cell in Excel.

        Args:
            sheet_name (str): The name of the sheet.
            cell_address (Optional[str]): The address of the cell. Defaults to "A1" if None.
        """
        if cell_address is None:
            cell_address = "A1"
        excel_handler.go_to_sheet_and_cell(sheet_name=sheet_name,
                                           cell_address=cell_address)

    def show_wait_indicator_in_current_tab(self):
        """Show a wait indicator in the current tab using a QProgressBar."""
        current_index = self.tab_widget.currentIndex()
        if current_index < 0:
            return

        current_tab = self.tab_widget.widget(current_index)
        if current_tab is None:
            return

        # Access the scroll area in the current tab
        scroll_area = current_tab.findChild(QScrollArea)
        if scroll_area is None:
            return

        # Save the original widget to restore later
        self.original_scroll_content = scroll_area.takeWidget()

        # Create indeterminate progress bar
        self.loading_indicator_tab = QProgressBar()
        self.loading_indicator_tab.setRange(0, 0)  # Indeterminate mode
        self.loading_indicator_tab.setTextVisible(False)
        self.loading_indicator_tab.setFixedHeight(20)
        self.loading_indicator_tab.setStyleSheet("""
            QProgressBar {
                background-color: #F0F0F0;
                border: none;
                border-radius: 10px;
            }
            QProgressBar::chunk {
                background-color: #007ACC;
                border-radius: 10px;
            }
        """)

        # Create a widget to hold the progress bar
        loading_widget = QWidget()
        layout = QVBoxLayout(loading_widget)
        layout.addStretch()
        layout.addWidget(self.loading_indicator_tab, alignment=Qt.AlignCenter)
        layout.addStretch()

        # Set the loading widget as the new widget in the scroll area
        scroll_area.setWidget(loading_widget)

    def hide_wait_indicator_in_current_tab(self):
        """Hide the wait indicator in the current tab and restore content."""
        current_index = self.tab_widget.currentIndex()
        if current_index < 0:
            return

        current_tab = self.tab_widget.widget(current_index)
        if current_tab is None:
            return

        scroll_area = current_tab.findChild(QScrollArea)
        if scroll_area is None:
            return

        # Remove loading indicator
        if self.loading_indicator_tab:
            self.loading_indicator_tab = None

        # Restore the original content
        if self.original_scroll_content:
            scroll_area.setWidget(self.original_scroll_content)
            self.original_scroll_content = None

    def apply_styles(self):
        """Apply styles to the main window."""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #a5cae6;
                color: #000000;
            }
            QLabel {
                color: #000000;
            }
            QPushButton {
                background-color: #007ACC;
                color: white;
                border: none;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #005F9E;
            }
            QPushButton:pressed {
                background-color: #003F6B;
            }
        """)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
