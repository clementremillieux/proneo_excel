"""_summary_

Returns:
    _type_: _description_
"""

from typing import Dict, List, Optional
from PyQt5.QtWidgets import QMainWindow, QLabel, QVBoxLayout, QWidget, QHBoxLayout, QTabWidget, QFrame, QPushButton, QApplication
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QPainter, QColor, QBrush

from modules.condition.schemas import CellsConditionReport, CellsConditionState
from modules.report import report_instance
from modules.checker.checker import Checker
from modules.condition import CELLS_CONDITIONS
from modules.report.schemas import UIReportCell
from modules.excel import excel_handler

checker = Checker(cells_conditions=CELLS_CONDITIONS)


def update_state() -> Dict[str, List[UIReportCell]]:
    """Appelée pour mettre à jour l'état."""
    cells_condition_report: List[
        CellsConditionReport] = checker.check_cells_conditions()

    report_instance.cells_condition_report = cells_condition_report

    report_ui: Dict[str, List[UIReportCell]] = report_instance.get_report()

    return report_ui


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
        self.setGeometry(100, 100, 600, 400)

        self.main_layout = QVBoxLayout()
        self.tab_widget = QTabWidget(self)
        self.main_layout.addWidget(self.tab_widget)

        container = QWidget()
        container.setLayout(self.main_layout)
        self.setCentralWidget(container)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.refresh_tabs)
        self.timer.start(2000)

    def refresh_tabs(self) -> None:
        """Met à jour les onglets avec les rapports actuels tout en conservant l'onglet sélectionné."""

        current_index = self.tab_widget.currentIndex()

        report_ui: Dict[str, List[UIReportCell]] = update_state()

        self.tab_widget.clear()

        for sheet_name, report_cells in report_ui.items():
            tab = QWidget()
            layout = QVBoxLayout()

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

                focus_button = QPushButton("Vérifier")

                focus_button.clicked.connect(lambda _, sn=cell.sheet_names[
                    0], ca=cell.cell_adress: self.focus_on_cell(sn, ca))

                h_layout.addWidget(circle)

                h_layout.addWidget(label)

                h_layout.addWidget(focus_button)

                h_layout.addStretch()

                frame.setLayout(h_layout)

                layout.addWidget(frame)

            tab.setLayout(layout)

            self.tab_widget.addTab(tab, sheet_name)

        if current_index >= 0 and current_index < self.tab_widget.count():
            self.tab_widget.setCurrentIndex(current_index)

    def focus_on_cell(self, sheet_name: str,
                      cell_address: Optional[str]) -> None:
        """Fonction pour naviguer vers une feuille et une cellule spécifique."""

        if cell_address is None:
            cell_address = "A1"

        excel_handler.go_to_sheet_and_cell(sheet_name=sheet_name,
                                           cell_address=cell_address)
