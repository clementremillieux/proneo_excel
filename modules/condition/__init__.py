"""_summary_"""

from modules.cells.schemas import BoxToCheck, DateToCheck

from modules.condition.condition import CellsConditions, ConditionDateSup, ConditionHasToBeFilled

from modules.sheet.schemas import SheetName

# OPAC DATE SUP AUDIT ################

cell_date_start = DateToCheck(sheet_name=SheetName.SHEET_2.value,
                              cell_address="B6")

cell_date_stop = DateToCheck(sheet_name=SheetName.SHEET_2.value,
                             cell_address="B7")

condition_opac_date_audit = ConditionDateSup(cell_date_start=cell_date_start,
                                             cell_date_stop=cell_date_stop,
                                             is_parent_condition=False)

CELLS_CONDITION_OPAC_DATE_AUDIT = CellsConditions(
    conditions=[condition_opac_date_audit])

######################################

# TEST HAS TO BE FILLED ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="C16")

test_has_to_be_filled = ConditionHasToBeFilled(cell=cell,
                                               is_parent_condition=False)

CELLS_CONDITION_TEST_HAS_TO_BE_FILLED = CellsConditions(
    conditions=[test_has_to_be_filled])

######################################

CELLS_CONDITIONS = [
    CELLS_CONDITION_OPAC_DATE_AUDIT, CELLS_CONDITION_TEST_HAS_TO_BE_FILLED
]
