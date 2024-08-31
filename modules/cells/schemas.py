"""_summary_"""

from datetime import datetime

from typing import Optional

from pydantic import BaseModel

from modules.excel import excel_handler

from modules.excel.schemas import CheckboxParams


class CellToCheck(BaseModel):
    """_summary_

    Args:
        BaseModel (_type_): _description_
    """

    sheet_name: str


class BoxToCheck(CellToCheck):
    """_summary_

    Args:
        CellToCheck (_type_): _description_
    """

    cell_address: str

    def get_value(self) -> Optional[str]:
        """_summary_

        Returns:
            str: _description_
        """

        return excel_handler.read_cell_value(sheet_name=self.sheet_name,
                                             cell_address=self.cell_address)


class DateToCheck(CellToCheck):
    """_summary_

    Args:
        CellToCheck (_type_): _description_
    """

    cell_address: str

    def get_value(self) -> Optional[datetime]:
        """_summary_

        Returns:
            str: _description_
        """

        return excel_handler.read_cell_date_value(
            sheet_name=self.sheet_name, cell_address=self.cell_address)


class CheckBoxToCheck(CellToCheck):
    """_summary_

    Args:
        CellToCheck (_type_): _description_
    """

    checkbox_name: str

    cell_address: str

    checkbox_params: CheckboxParams

    alias_name: Optional[str] = None

    def get_value(self) -> bool:
        """_summary_

        Returns:
            str: _description_
        """

        return excel_handler.get_checkbox_state(cell_address=self.cell_address,
                                                sheet_name=self.sheet_name)

        # return excel_handler.get_checkbox_state(
        #     checkbox_params=self.checkbox_params,
        #     checkbox_name=self.checkbox_name)
