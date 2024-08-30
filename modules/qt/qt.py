import os
import shutil
from typing import Dict, List, Optional
from PyQt5.QtWidgets import (QMainWindow, QLabel, QVBoxLayout, QWidget,
                             QHBoxLayout, QTabWidget, QFrame, QPushButton,
                             QScrollArea, QDesktopWidget, QFileDialog,
                             QInputDialog)
from PyQt5.QtCore import QTimer, QThread, pyqtSignal, Qt
from PyQt5.QtGui import QPainter, QColor, QBrush

from modules.condition.schemas import CellsConditionReport, CellsConditionState
from modules.params.schemas import AppParams
from modules.report import report_instance
from modules.checker.checker import Checker
from modules.condition import CELLS_CONDITIONS
from modules.report.schemas import UIReportCell
from modules.excel import excel_handler

from config.logger_config import logger

checker = Checker(cells_conditions=CELLS_CONDITIONS)


def update_state() -> Dict[str, List[UIReportCell]]:
    """Appelée pour mettre à jour l'état."""

    if excel_handler.excel_abs_path:
        excel_handler.load_excel(excel_abs_path=excel_handler.excel_abs_path)

    cells_condition_report: List[
        CellsConditionReport] = checker.check_cells_conditions()

    report_instance.cells_condition_report = cells_condition_report

    report_ui: Dict[str, List[UIReportCell]] = report_instance.get_report()

    return report_ui


class Worker(QThread):
    update_completed = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)

    def run(self):
        try:
            # Perform the time-consuming task in a separate thread
            report_ui = update_state()
            self.update_completed.emit(report_ui)
        except Exception as e:
            self.error_occurred.emit(str(e))


class ColoredCircle(QWidget):
    """Widget pour afficher un cercle coloré."""

    def __init__(self, color: QColor, parent=None):
        super().__init__(parent)
        self.color = color
        self.setFixedSize(20, 20)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        brush = QBrush(self.color)
        painter.setBrush(brush)
        painter.drawEllipse(0, 0, 20, 20)


class MainWindow(QMainWindow):
    """Fenêtre principale de l'application PyQt."""

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Pronéo")
        self.setGeometry(100, 200, 800, 800)

        # Set window flags to stay on top and be frameless
        # self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)

        self.main_layout = QVBoxLayout()

        # Create a custom title bar
        self.title_bar = QWidget()
        self.title_bar_layout = QHBoxLayout(self.title_bar)
        self.title_bar_layout.setContentsMargins(0, 0, 0, 0)
        self.title_label = QLabel("Pronéo")
        self.title_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.hide_button = QPushButton("_")
        self.hide_button.setFixedSize(30, 30)
        self.hide_button.clicked.connect(self.showMinimized)
        self.close_button = QPushButton("X")
        self.close_button.setFixedSize(30, 30)
        self.close_button.clicked.connect(self.close)

        self.title_bar_layout.addWidget(self.title_label)
        self.title_bar_layout.addStretch()
        self.title_bar_layout.addWidget(self.hide_button)
        self.title_bar_layout.addWidget(self.close_button)

        self.main_layout.addWidget(self.title_bar)

        self.file_path_intro = "Fichier vérifié : "

        self.file_path_label = QLabel(self)
        self.file_path_label.setText(
            f"{self.file_path_intro} {excel_handler.excel_abs_path}"
            if excel_handler.excel_abs_path else
            f"{self.file_path_intro} veuillez séléctionnez un fichier")

        self.main_layout.addWidget(self.file_path_label)

        self.load_button = QPushButton("Charger Excel")
        self.load_button.setFixedSize(120, 30)
        self.load_button.clicked.connect(self.load_excel_file)
        self.main_layout.addWidget(self.load_button)

        self.new_file_button = QPushButton("Nouveau Excel")
        self.new_file_button.setFixedSize(120, 30)
        self.new_file_button.clicked.connect(self.create_new_excel)
        self.main_layout.addWidget(self.new_file_button)

        self.tab_widget = QTabWidget(self)
        self.main_layout.addWidget(self.tab_widget)

        container = QWidget()
        container.setLayout(self.main_layout)
        self.setCentralWidget(container)

        self.scroll_positions = {}  # To store scroll positions for each tab

        # Position the window in the top right corner
        self.position_window()

        # Initialize and start the timer
        self.init_timer()

        self.worker = Worker()
        self.worker.update_completed.connect(self.on_update_completed)
        self.worker.error_occurred.connect(self.on_worker_error)

    def on_worker_error(self, error_message):
        logger.error(f"Worker error: {error_message}")

    def init_timer(self):
        """Initialize the timer to call check_and_refresh every X seconds."""
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.start_worker)
        self.timer.start(1000)

    def start_worker(self):
        """Starts the worker thread to perform the update in the background."""
        if not self.worker.isRunning():
            self.worker.start()

    def on_update_completed(self, report_ui: dict):
        """Handle the completion of the update and refresh the tabs."""
        self.refresh_tabs(report_ui)

    def create_new_excel(self):
        """Create a new Excel file by copying an existing one."""

        new_file_name, ok = QInputDialog.getText(
            self, "Nouveau Excel (xlsm)",
            "Entrez le nom du fichier (sans extension):")

        if ok and new_file_name:
            new_file_name = f"{new_file_name}.xlsm"

            # Get the source file path from AppParams
            source_file_path = AppParams().excel_abs_path

            # Get the user's Desktop directory
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

            destination_file_path = os.path.join(desktop_path, new_file_name)

            try:
                # Copy the file to the new location on the Desktop
                shutil.copy(source_file_path, destination_file_path)

                # Update the label with the new file path
                self.file_path_label.setText(
                    f"{self.file_path_intro} {destination_file_path}")

                # Load the new Excel file and update the UI
                excel_handler.load_excel(excel_abs_path=destination_file_path)

                self.start_worker()

            except Exception as e:
                logger(f"Failed to create new Excel file: {e}")

    def load_excel_file(self):
        """Open a file dialog to load an Excel file."""
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "Charger le fichier Excel",
            "",
            "Excel Files (*.xlsx *.xlsm)",
            options=options,
        )
        if file_name:
            # Load the selected Excel file
            excel_handler.load_excel(excel_abs_path=file_name)

            # Update the label with the new file path
            self.file_path_label.setText(f"{self.file_path_intro} {file_name}")

            # Refresh the state
            self.start_worker()

    def position_window(self):
        screen = QDesktopWidget().screenNumber(QDesktopWidget().cursor().pos())
        screen_geometry = QDesktopWidget().availableGeometry(screen)
        window_geometry = self.frameGeometry()
        window_geometry.moveTopRight(screen_geometry.topRight())
        self.move(window_geometry.topLeft())

    def refresh_tabs(self, report_ui: Dict[str, List[UIReportCell]]) -> None:
        """Met à jour les onglets avec les rapports actuels tout en conservant l'onglet sélectionné et la position de défilement."""

        current_index = self.tab_widget.currentIndex()
        current_tab_name = self.tab_widget.tabText(
            current_index) if current_index >= 0 else None

        # Store current scroll positions
        for i in range(self.tab_widget.count()):
            tab_name = self.tab_widget.tabText(i)
            scroll_area = self.tab_widget.widget(i).findChild(QScrollArea)
            if scroll_area:
                self.scroll_positions[
                    tab_name] = scroll_area.verticalScrollBar().value()

        self.tab_widget.clear()

        for sheet_name, report_cells in report_ui.items():
            tab = QWidget()
            scroll_area = QScrollArea()
            scroll_area.setWidgetResizable(True)
            scroll_content = QWidget()
            layout = QVBoxLayout(scroll_content)

            for cell in report_cells:
                frame = QFrame()
                frame.setFrameShape(QFrame.Shape.Box)
                frame.setFrameShadow(QFrame.Shadow.Raised)
                frame.setLineWidth(2)

                h_layout = QHBoxLayout()

                color = QColor(
                    'green'
                ) if cell.state == CellsConditionState.OK else QColor('red')
                circle = ColoredCircle(color)

                label = QLabel(cell.instruction)
                label.setWordWrap(True)

                focus_button = QPushButton("Accéder")
                focus_button.clicked.connect(lambda _, sn=cell.sheet_names[
                    0], ca=cell.cell_adress: self.focus_on_cell(sn, ca))

                h_layout.addWidget(circle)
                h_layout.addWidget(label, 1)
                h_layout.addWidget(focus_button)

                frame.setLayout(h_layout)
                layout.addWidget(frame)

            layout.addStretch()
            scroll_content.setLayout(layout)
            scroll_area.setWidget(scroll_content)

            tab_layout = QVBoxLayout()
            tab_layout.addWidget(scroll_area)
            tab.setLayout(tab_layout)

            self.tab_widget.addTab(tab, sheet_name)

            # Restore scroll position
            if sheet_name in self.scroll_positions:
                QTimer.singleShot(
                    0,
                    lambda pos=self.scroll_positions[sheet_name], sa=
                    scroll_area: sa.verticalScrollBar().setValue(pos))

        # Set the previously selected tab
        if current_tab_name:
            for i in range(self.tab_widget.count()):
                if self.tab_widget.tabText(i) == current_tab_name:
                    self.tab_widget.setCurrentIndex(i)
                    break

    def focus_on_cell(self, sheet_name: str,
                      cell_address: Optional[str]) -> None:
        """Fonction pour naviguer vers une feuille et une cellule spécifique."""

        if cell_address is None:
            cell_address = "A1"

        excel_handler.go_to_sheet_and_cell(sheet_name=sheet_name,
                                           cell_address=cell_address)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragPosition = event.globalPos() - self.frameGeometry(
            ).topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.dragPosition)
            event.accept()
