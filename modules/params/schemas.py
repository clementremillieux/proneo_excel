"""_summary_"""
import os

import sys

import platform

from pydantic import BaseModel


def get_base_dir():
    """Determine the base directory depending on whether the app is bundled or not."""

    if getattr(sys, 'frozen', False):

        if platform.system() == "Darwin":

            base_dir = os.path.dirname(sys.executable)

            resources_dir = os.path.join(base_dir, '..', 'Resources')

            return os.path.abspath(resources_dir)

        base_dir = sys._MEIPASS

        return os.path.abspath(base_dir)

    return os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..')


def test_open_excel_file():
    """Test function to open the file at AppParams.excel_abs_path."""
    try:
        # Create an instance of AppParams to get the excel_abs_path
        app_params = AppParams()

        # Check if the file exists
        if not os.path.exists(app_params.excel_abs_path):
            print(
                f"Test Failed: File not found at path: {app_params.excel_abs_path}"
            )
            return

        # Attempt to open the file in read-binary mode
        with open(app_params.excel_abs_path, 'rb') as file:
            print(
                f"Test Passed: Successfully opened the file at {app_params.excel_abs_path}"
            )

    except Exception as e:
        print(
            f"Test Failed: An error occurred while trying to open the file: {str(e)}"
        )


class AppParams(BaseModel):
    """_summary_

    Args:
        BaseModel (_type_): _description_
    """

    base_dir: str = get_base_dir()

    excel_abs_path: str = os.path.join(
        base_dir, 'data', "Plan et Rapport d'audit certification V32.xlsm")

    vba_checkbox_module: str = "StoreSpecificCheckboxValue"

    vba_checkbox_result_sheet_name: str = "OPAC"

    vba_checkbox_result_cells: str = "A1"


test_open_excel_file()
