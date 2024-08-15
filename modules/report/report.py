"""_summary_"""

from typing import Any, Dict, List, Tuple

from modules.condition.schemas import CellsConditionState, ConditionType, CellsConditionReport

from modules.cells.schemas import BoxToCheck, CheckBoxToCheck, DateToCheck

from config.logger_config import logger
from modules.report.schemas import UIReportCell


class Report:
    """_summary_"""

    def __init__(self) -> None:
        self.cells_condition_report: List[CellsConditionReport] = []

    def log_condition_result(self, result: Any,
                             cells: List[BoxToCheck | CheckBoxToCheck
                                         | DateToCheck],
                             condition_type: ConditionType,
                             report_str: str) -> None:
        """_summary_
        """

        cells_name = " ".join(
            f"{cell.sheet_name} {cell.cell_address}" if isinstance(
                cell, BoxToCheck
                | DateToCheck) else f"{cell.sheet_name} {cell.checkbox_name}"
            for cell in cells)

        logger.info("Condition %s [%s] : %s", cells_name, condition_type.value,
                    result)

        logger.info("\t-> %s", report_str)

    def get_report(self) -> Dict[str, List[UIReportCell]]:
        """_summary_"""

        report: Dict[str, List[UIReportCell]] = {}

        for cell_condition_report in self.cells_condition_report:
            self.log_condition_result(
                result=cell_condition_report.state.value,
                cells=cell_condition_report.condition.cells_list,
                condition_type=cell_condition_report.condition.condition_type,
                report_str=cell_condition_report.report_str)

            report_key: str = cell_condition_report.condition.cells_list[
                0].sheet_name

            if not report.get(report_key, None):
                report[report_key] = [
                    UIReportCell(
                        state=cell_condition_report.state,
                        instruction=cell_condition_report.report_str,
                        sheet_names=[
                            cell.sheet_name for cell in
                            cell_condition_report.condition.cells_list
                        ],
                        cell_adress=cell_condition_report.condition.
                        cells_list[0].cell_address if not isinstance(
                            cell_condition_report.condition.cells_list[0],
                            CheckBoxToCheck) else None)
                ]

        return report
