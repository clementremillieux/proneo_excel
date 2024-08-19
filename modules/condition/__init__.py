"""_summary_"""

from modules.cells.schemas import BoxToCheck, CheckBoxToCheck, DateToCheck

from modules.condition.condition import (
    CellsConditions, ConditionAtLeastOneCellAmongList,
    ConditionAtLeastOneCheckBoxAmongList, ConditionCheckAllSheetDescription,
    ConditionCheckAllSheetDropDown, ConditionCheckAllSheetReference,
    ConditionDateSup, ConditionDateDurationBetween, ConditionHasNc,
    ConditionHasToBeChecked, ConditionHasToBeValue,
    ConditionIsNCFromCellNumber, ConditionIsNCFromCellText,
    ConditionIsNcMajFromCellNumber, ConditionNcAllJChoosed,
    ConditionOneCheckBoxAmongList, ConditionHasToBeFilled)

from modules.excel.schemas import CheckboxParams

from modules.sheet.schemas import SheetName

checkbox_params = CheckboxParams(
    apple_script_path="modules/excel/apple_script/checkbox.scpt")

###########################################
############### OPAC ######################
###########################################

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

# OPAC DURATION CORRESPONDING TO DATE ################

cell_date_start = DateToCheck(sheet_name=SheetName.SHEET_2.value,
                              cell_address="B6")

cell_date_stop = DateToCheck(sheet_name=SheetName.SHEET_2.value,
                             cell_address="B7")

cell_duration = BoxToCheck(sheet_name=SheetName.SHEET_2.value,
                           cell_address="B8")

condition_opac_duration = ConditionDateDurationBetween(
    cell_date_start=cell_date_start,
    cell_date_stop=cell_date_stop,
    cell_duration=cell_duration,
    is_parent_condition=False)

CELLS_CONDITION_OPAC_DURATION = CellsConditions(
    conditions=[condition_opac_duration])

######################################

# OPAC AUDIT TYPE ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 66",
                         cell_address="B9",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 67",
                         cell_address="B10",
                         checkbox_params=checkbox_params)

condition_opac_audit_type = ConditionOneCheckBoxAmongList(
    cells=[cell_1, cell_2], is_parent_condition=False)

CELLS_CONDITION_OPAC_AUDIT_TYPE = CellsConditions(
    conditions=[condition_opac_audit_type])

######################################

# OPAC COMPANY NAME ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="B12")

company_name_has_to_be_filled = ConditionHasToBeFilled(
    cell=cell, is_parent_condition=False)

CELLS_CONDITION_OPAC_COMPANY_NAME = CellsConditions(
    conditions=[company_name_has_to_be_filled])

######################################

# OPAC COMPANY ADRESSE ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="B13")

company_adresse_has_to_be_filled = ConditionHasToBeFilled(
    cell=cell, is_parent_condition=False)

CELLS_CONDITION_OPAC_COMPANY_ADRESS = CellsConditions(
    conditions=[company_adresse_has_to_be_filled])

######################################

# OPAC COMPANY CP ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="B14")

company_cp_has_to_be_filled = ConditionHasToBeFilled(cell=cell,
                                                     is_parent_condition=False)

CELLS_CONDITION_OPAC_COMPANY_CP = CellsConditions(
    conditions=[company_cp_has_to_be_filled])

######################################

# OPAC COMPANY CITY ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="B15")

company_city_has_to_be_filled = ConditionHasToBeFilled(
    cell=cell, is_parent_condition=False)

CELLS_CONDITION_OPAC_COMPANY_CITY = CellsConditions(
    conditions=[company_city_has_to_be_filled])

######################################

# OPAC COMPANY SIREN ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="B16")

company_siren_has_to_be_filled = ConditionHasToBeFilled(
    cell=cell, is_parent_condition=False)

CELLS_CONDITION_OPAC_COMPANY_SIREN = CellsConditions(
    conditions=[company_siren_has_to_be_filled])

######################################

# OPAC COMPANY NDA ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="B17")

company_nda_has_to_be_filled = ConditionHasToBeFilled(
    cell=cell, is_parent_condition=False)

CELLS_CONDITION_OPAC_COMPANY_NDA = CellsConditions(
    conditions=[company_nda_has_to_be_filled])

######################################

# OPAC RP NAME ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="B21")

rp_name_has_to_be_filled = ConditionHasToBeFilled(cell=cell,
                                                  is_parent_condition=False)

CELLS_CONDITION_OPAC_RP_NAME = CellsConditions(
    conditions=[rp_name_has_to_be_filled])

######################################

# OPAC RP ROLE ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="B22")

rp_role_has_to_be_filled = ConditionHasToBeFilled(cell=cell,
                                                  is_parent_condition=False)

CELLS_CONDITION_OPAC_RP_ROLE = CellsConditions(
    conditions=[rp_role_has_to_be_filled])

######################################

# OPAC RP PHONE ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="B23")

rp_phone_has_to_be_filled = ConditionHasToBeFilled(cell=cell,
                                                   is_parent_condition=False)

CELLS_CONDITION_OPAC_RP_PHONE = CellsConditions(
    conditions=[rp_phone_has_to_be_filled])

######################################

# OPAC RP MAIL ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="B24")

rp_mail_has_to_be_filled = ConditionHasToBeFilled(cell=cell,
                                                  is_parent_condition=False)

CELLS_CONDITION_OPAC_RP_MAIL = CellsConditions(
    conditions=[rp_mail_has_to_be_filled])

######################################

# OPAC DISTANCE ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 103",
                         cell_address="B33",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 104",
                         cell_address="C33",
                         checkbox_params=checkbox_params)

condition_opac_distance = ConditionOneCheckBoxAmongList(
    cells=[cell_1, cell_2], is_parent_condition=False)

CELLS_CONDITION_OPAC_DISTANCE = CellsConditions(
    conditions=[condition_opac_distance])

######################################

# OPAC COMPANY PERIOD ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 105",
                         cell_address="B34",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 106",
                         cell_address="C34",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_PERIOD = CellsConditions(conditions=[
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2],
                                  is_parent_condition=False)
])

######################################

# OPAC MORE 2 DAY ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 107",
                         cell_address="B36",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 108",
                         cell_address="C36",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_MORE_TWO_DAY = CellsConditions(conditions=[
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2],
                                  is_parent_condition=False)
])

######################################

# OPAC RNCP ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 109",
                         cell_address="B37",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 110",
                         cell_address="C37",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_RNCP = CellsConditions(conditions=[
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2],
                                  is_parent_condition=False)
])

######################################

# OPAC RS ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 113",
                         cell_address="B38",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 114",
                         cell_address="C38",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_RS = CellsConditions(conditions=[
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2],
                                  is_parent_condition=False)
])

######################################

# OPAC SUB CONTRACTOR ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 115",
                         cell_address="B39",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 116",
                         cell_address="C39",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_SUB_CONTRACTOR = CellsConditions(conditions=[
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2],
                                  is_parent_condition=False)
])

######################################

# OPAC SUB CONTRACTOR ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 127",
                         cell_address="B40",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 128",
                         cell_address="C40",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_SUB_CONTRACTOR_2 = CellsConditions(conditions=[
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2],
                                  is_parent_condition=False)
])

######################################

# OPAC PSH ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 117",
                         cell_address="B41",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 118",
                         cell_address="C41",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_PSH = CellsConditions(conditions=[
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2],
                                  is_parent_condition=False)
])

######################################

# OPAC CONTRACTOR 100% ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 121",
                         cell_address="D40",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 122",
                         cell_address="D41",
                         checkbox_params=checkbox_params)

cell_3 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 129",
                         cell_address="D42",
                         checkbox_params=checkbox_params)

cell_4 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 130",
                         cell_address="D43",
                         checkbox_params=checkbox_params)

cell_5 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 131",
                         cell_address="D44",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_CONTRACTOR_100 = CellsConditions(conditions=[
    ConditionAtLeastOneCheckBoxAmongList(
        cells=[cell_1, cell_2, cell_3, cell_4, cell_5],
        is_parent_condition=False)
])

######################################

######################################

# OPAC DESCRIPTION ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="A47")

CELLS_CONDITION_OPAC_DESCRIPTION = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# OPAC AF FOLDER ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="A61")

cell_to_check = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                checkbox_name="Check Box 41",
                                cell_address="F22",
                                checkbox_params=checkbox_params)

cell_to_check_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                  checkbox_name="Check Box 45",
                                  cell_address="G22",
                                  checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_AF_FOLDER = CellsConditions(conditions=[
    ConditionAtLeastOneCheckBoxAmongList(
        cells=[cell_to_check, cell_to_check_2], is_parent_condition=True),
    ConditionHasToBeFilled(cell=cell, is_parent_condition=False),
])

######################################

# OPAC AF JUSTIFY ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="A64")

cell_to_check = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                checkbox_name="Check Box 41",
                                cell_address="F22",
                                checkbox_params=checkbox_params)

cell_to_check_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                  checkbox_name="Check Box 45",
                                  cell_address="G22",
                                  checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_AF_JUSTIFY = CellsConditions(conditions=[
    ConditionAtLeastOneCheckBoxAmongList(
        cells=[cell_to_check, cell_to_check_2], is_parent_condition=True),
    ConditionHasToBeFilled(cell=cell, is_parent_condition=False),
])
######################################

# OPAC BC FOLDER ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="A67")

cell_to_check = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                checkbox_name="Check Box 42",
                                cell_address="F23",
                                checkbox_params=checkbox_params)

cell_to_check_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                  checkbox_name="Check Box 46",
                                  cell_address="G23",
                                  checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_BC_FOLDER = CellsConditions(conditions=[
    ConditionAtLeastOneCheckBoxAmongList(
        cells=[cell_to_check, cell_to_check_2], is_parent_condition=True),
    ConditionHasToBeFilled(cell=cell, is_parent_condition=False),
])

######################################

# OPAC BC JUSTIFY ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="A70")

cell_to_check = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                checkbox_name="Check Box 42",
                                cell_address="F23",
                                checkbox_params=checkbox_params)

cell_to_check_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                  checkbox_name="Check Box 46",
                                  cell_address="G23",
                                  checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_BC_JUSTIFY = CellsConditions(conditions=[
    ConditionAtLeastOneCheckBoxAmongList(
        cells=[cell_to_check, cell_to_check_2], is_parent_condition=True),
    ConditionHasToBeFilled(cell=cell, is_parent_condition=False),
])

######################################

# OPAC VAE FOLDER ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="A73")

cell_to_check = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                checkbox_name="Check Box 43",
                                cell_address="F24",
                                checkbox_params=checkbox_params)

cell_to_check_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                  checkbox_name="Check Box 47",
                                  cell_address="G24",
                                  checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_VAE_FOLDER = CellsConditions(conditions=[
    ConditionAtLeastOneCheckBoxAmongList(
        cells=[cell_to_check, cell_to_check_2], is_parent_condition=True),
    ConditionHasToBeFilled(cell=cell, is_parent_condition=False),
])

######################################

# OPAC VAE JUSTIFY ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="A76")

cell_to_check = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                checkbox_name="Check Box 43",
                                cell_address="F24",
                                checkbox_params=checkbox_params)

cell_to_check_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                  checkbox_name="Check Box 47",
                                  cell_address="G24",
                                  checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_VAE_JUSTIFY = CellsConditions(conditions=[
    ConditionAtLeastOneCheckBoxAmongList(
        cells=[cell_to_check, cell_to_check_2], is_parent_condition=True),
    ConditionHasToBeFilled(cell=cell, is_parent_condition=False),
])

######################################

# OPAC FA FOLDER ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="A79")

cell_to_check = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                checkbox_name="Check Box 44",
                                cell_address="F25",
                                checkbox_params=checkbox_params)

cell_to_check_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                  checkbox_name="Check Box 48",
                                  cell_address="G25",
                                  checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_FA_FOLDER = CellsConditions(conditions=[
    ConditionAtLeastOneCheckBoxAmongList(
        cells=[cell_to_check, cell_to_check_2], is_parent_condition=True),
    ConditionHasToBeFilled(cell=cell, is_parent_condition=False),
])

######################################

# OPAC FA JUSTIFY ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="A82")

cell_to_check = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                checkbox_name="Check Box 44",
                                cell_address="F25",
                                checkbox_params=checkbox_params)

cell_to_check_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                  checkbox_name="Check Box 48",
                                  cell_address="G25",
                                  checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_FA_JUSTIFY = CellsConditions(conditions=[
    ConditionAtLeastOneCheckBoxAmongList(
        cells=[cell_to_check, cell_to_check_2], is_parent_condition=True),
    ConditionHasToBeFilled(cell=cell, is_parent_condition=False),
])

######################################

# OPAC AUDITOR NAME ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="F8")

CELLS_CONDITION_OPAC_AUDITOR_NAME = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# OPAC AUDITOR PHONE ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="F9")

CELLS_CONDITION_OPAC_AUDITOR_PHONE = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# OPAC AUDITOR MAIL ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="F10")

CELLS_CONDITION_OPAC_AUDITOR_MAIL = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# OPAC REMOTE CONDITION ################

cell_value = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="E14")

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 57",
                         cell_address="F17",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 59",
                         cell_address="G17",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_REMOTE_CONDITION = CellsConditions(conditions=[
    ConditionHasToBeValue(
        cell=cell_value, value="Audit à distance", is_parent_condition=True),
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2],
                                  is_parent_condition=False)
])

######################################

# OPAC REMOTE GOAL ################

cell_value = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="E14")

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 58",
                         cell_address="F18",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 60",
                         cell_address="G18",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_REMOTE_GOAL = CellsConditions(conditions=[
    ConditionHasToBeValue(
        cell=cell_value, value="Audit à distance", is_parent_condition=True),
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2],
                                  is_parent_condition=False)
])

######################################

# OPAC REMOTE TOOLS ################

cell_value = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="E14")

cell = BoxToCheck(
    sheet_name=SheetName.SHEET_2.value,
    cell_address="F19",
)

CELLS_CONDITION_OPAC_REMOTE_TOOLS = CellsConditions(conditions=[
    ConditionHasToBeValue(
        cell=cell_value, value="Audit à distance", is_parent_condition=True),
    ConditionHasToBeFilled(cell=cell, is_parent_condition=False)
])

######################################

# OPAC ACTION ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 41",
                         cell_address="F22",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 42",
                         cell_address="F23",
                         checkbox_params=checkbox_params)

cell_3 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 43",
                         cell_address="F24",
                         checkbox_params=checkbox_params)

cell_4 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 44",
                         cell_address="F25",
                         checkbox_params=checkbox_params)

cell_5 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 45",
                         cell_address="G22",
                         checkbox_params=checkbox_params)

cell_6 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 46",
                         cell_address="G23",
                         checkbox_params=checkbox_params)

cell_7 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 47",
                         cell_address="G24",
                         checkbox_params=checkbox_params)

cell_8 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 48",
                         cell_address="G25",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_ACTION = CellsConditions(conditions=[
    ConditionAtLeastOneCheckBoxAmongList(
        cells=[cell_1, cell_2, cell_3, cell_4, cell_5, cell_6, cell_7, cell_8],
        is_parent_condition=False)
])

######################################

# OPAC AUDIT TYPE BOX ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 51",
                         cell_address="F28",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 52",
                         cell_address="F29",
                         checkbox_params=checkbox_params)

cell_3 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 53",
                         cell_address="F30",
                         checkbox_params=checkbox_params)

cell_4 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 54",
                         cell_address="F31",
                         checkbox_params=checkbox_params)

cell_5 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 55",
                         cell_address="F32",
                         checkbox_params=checkbox_params)

cell_6 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 120",
                         cell_address="F33",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_AUDIT_TYPE_2 = CellsConditions(conditions=[
    ConditionOneCheckBoxAmongList(
        cells=[cell_1, cell_2, cell_3, cell_4, cell_5, cell_6],
        is_parent_condition=False)
])

######################################

# OPAC CNEFOP ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 50",
                         cell_address="F37",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 49",
                         cell_address="G37",
                         checkbox_params=checkbox_params)

cell_3 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 95",
                         cell_address="H37",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_CNEFOP = CellsConditions(conditions=[
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2, cell_3],
                                  is_parent_condition=False)
])

######################################

# OPAC WATCHING ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 78",
                         cell_address="F40",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 79",
                         cell_address="G40",
                         checkbox_params=checkbox_params)

cell_3 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 94",
                         cell_address="H40",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_WATCHING = CellsConditions(conditions=[
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2, cell_3],
                                  is_parent_condition=False)
])

######################################

# OPAC NC ################

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 81",
                         cell_address="F41",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 82",
                         cell_address="G41",
                         checkbox_params=checkbox_params)

cell_3 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 96",
                         cell_address="H41",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_NC = CellsConditions(conditions=[
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2, cell_3],
                                  is_parent_condition=False)
])

######################################

# OPAC NC ID ################

cell_value = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="F43")

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 81",
                         cell_address="F41",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_NC_ID = CellsConditions(conditions=[
    ConditionHasToBeChecked(cell=cell_1, is_parent_condition=True),
    ConditionHasToBeFilled(cell=cell_value, is_parent_condition=False)
])

######################################

# OPAC NC PREVIOUS ################

cell_to_check = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                checkbox_name="Check Box 81",
                                cell_address="F41",
                                checkbox_params=checkbox_params)

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 88",
                         cell_address="F44",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 90",
                         cell_address="G44",
                         checkbox_params=checkbox_params)

cell_3 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 97",
                         cell_address="H44",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_NC_PREVIOUS = CellsConditions(conditions=[
    ConditionHasToBeChecked(cell=cell_to_check, is_parent_condition=True),
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2, cell_3],
                                  is_parent_condition=False)
])

######################################

# OPAC NC CORRECTION ################

cell_to_check = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                checkbox_name="Check Box 81",
                                cell_address="F41",
                                checkbox_params=checkbox_params)

cell_1 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 92",
                         cell_address="F46",
                         checkbox_params=checkbox_params)

cell_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 93",
                         cell_address="G46",
                         checkbox_params=checkbox_params)

cell_3 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                         checkbox_name="Check Box 98",
                         cell_address="H46",
                         checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_NC_CORRECTION = CellsConditions(conditions=[
    ConditionHasToBeChecked(cell=cell_to_check, is_parent_condition=True),
    ConditionOneCheckBoxAmongList(cells=[cell_1, cell_2, cell_3],
                                  is_parent_condition=False)
])

######################################

# OPAC NC SHEET ################

cell = BoxToCheck(
    sheet_name=SheetName.SHEET_2.value,
    cell_address="F48",
)

CELLS_CONDITION_OPAC_NC_SHEET = CellsConditions(conditions=[
    ConditionIsNCFromCellText(cell=cell, is_parent_condition=False),
])

# OPAC NC MAJ ################

cell = BoxToCheck(
    sheet_name=SheetName.SHEET_2.value,
    cell_address="G52",
)

CELLS_CONDITION_OPAC_NC_MAJ = CellsConditions(conditions=[
    ConditionIsNcMajFromCellNumber(cell=cell, is_parent_condition=False),
])

######################################

######################################

# OPAC NC SHEET ################

cell = BoxToCheck(
    sheet_name=SheetName.SHEET_2.value,
    cell_address="G53",
)

CELLS_CONDITION_OPAC_NC_SHEET_NUMBER = CellsConditions(conditions=[
    ConditionIsNCFromCellNumber(cell=cell, is_parent_condition=False),
])

######################################

# OPAC DATE SEND ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="G58")

CELLS_CONDITION_OPAC_DATE_SEND = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# OPAC DATE SEND NC ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="G59")

CELLS_CONDITION_OPAC_DATE_SEND_NC = CellsConditions(conditions=[
    ConditionHasNc(is_parent_condition=True),
    ConditionHasToBeFilled(cell=cell, is_parent_condition=False)
])

######################################

# OPAC DATE SEND RAPPORT ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="G60")

CELLS_CONDITION_OPAC_DATE_SEND_RAPPORT = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# OPAC PASS ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="G60")

cell_to_check = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                checkbox_name="Check Box 62",
                                cell_address="F63",
                                checkbox_params=checkbox_params)

cell_to_check_2 = CheckBoxToCheck(sheet_name=SheetName.SHEET_2.value,
                                  checkbox_name="Check Box 63",
                                  cell_address="F64",
                                  checkbox_params=checkbox_params)

CELLS_CONDITION_OPAC_PASS = CellsConditions(conditions=[
    ConditionHasToBeFilled(cell=cell, is_parent_condition=True),
    ConditionOneCheckBoxAmongList(cells=[cell_to_check, cell_to_check_2],
                                  is_parent_condition=False)
])

######################################

# OPAC DATE SIGNATURE ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_2.value, cell_address="E71")

CELLS_CONDITION_OPAC_SIGNATURE = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# PLAN DATE 1 ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A17")

CELLS_CONDITION_PLAN_DATE_1 = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# PLAN DATE 1 ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A17")

CELLS_CONDITION_PLAN_DATE_1 = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# # PLAN DATE 2 ################

# cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A18")

# CELLS_CONDITION_PLAN_DATE_2 = CellsConditions(
#     conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

# ######################################

# PLAN DATE 3 ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A19")

CELLS_CONDITION_PLAN_DATE_3 = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# PLAN DATE 4 ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A20")

CELLS_CONDITION_PLAN_DATE_4 = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# PLAN DATE 5 ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A21")

CELLS_CONDITION_PLAN_DATE_5 = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# PLAN DATE 6 ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A22")

CELLS_CONDITION_PLAN_DATE_6 = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# PLAN DATE 7 ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A23")

CELLS_CONDITION_PLAN_DATE_7 = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# PLAN DATE 8 ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A24")

CELLS_CONDITION_PLAN_DATE_8 = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# PLAN DATE 9 ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A25")

CELLS_CONDITION_PLAN_DATE_9 = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# PLAN DATE 10 ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A26")

CELLS_CONDITION_PLAN_DATE_10 = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# PLAN DATE 11 ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A27")

CELLS_CONDITION_PLAN_DATE_11 = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# # PLAN DATE 12 ################

# cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A28")

# CELLS_CONDITION_PLAN_DATE_12 = CellsConditions(
#     conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

# ######################################

# PLAN DATE 13 ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A29")

CELLS_CONDITION_PLAN_DATE_13 = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# PLAN DATE 14 ################

cell = BoxToCheck(sheet_name=SheetName.SHEET_4.value, cell_address="A30")

CELLS_CONDITION_PLAN_DATE_14 = CellsConditions(
    conditions=[ConditionHasToBeFilled(cell=cell, is_parent_condition=False)])

######################################

# SIGN IN OPEN NAME ################

CELLS_CONDITION_SIGN_IN_OPEN_NAME = CellsConditions(conditions=[
    ConditionAtLeastOneCellAmongList(cells=[
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="C16"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="C17"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="C18"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="C19"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="C20"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="C21")
    ],
                                     is_parent_condition=False)
])

######################################

# SIGN IN OPEN LASTNAME ################

CELLS_CONDITION_SIGN_IN_OPEN_LASTNAME = CellsConditions(conditions=[
    ConditionAtLeastOneCellAmongList(cells=[
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="D16"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="D17"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="D18"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="D19"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="D20"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="D21")
    ],
                                     is_parent_condition=False)
])

######################################

# SIGN IN OPEN FUNCTION ################

CELLS_CONDITION_SIGN_IN_OPEN_FUNCTION = CellsConditions(conditions=[
    ConditionAtLeastOneCellAmongList(cells=[
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="E16"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="E17"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="E18"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="E19"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="E20"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="E21")
    ],
                                     is_parent_condition=False)
])

######################################

# SIGN IN CLOSE NAME ################

CELLS_CONDITION_SIGN_IN_CLOSE_NAME = CellsConditions(conditions=[
    ConditionAtLeastOneCellAmongList(cells=[
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="C26"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="C27"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="C28"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="C29"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="C30"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="C31")
    ],
                                     is_parent_condition=False)
])

######################################

# SIGN IN CLOSE LASTNAME ################

CELLS_CONDITION_SIGN_IN_CLOSE_LASTNAME = CellsConditions(conditions=[
    ConditionAtLeastOneCellAmongList(cells=[
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="D26"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="D27"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="D28"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="D29"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="D30"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="D31")
    ],
                                     is_parent_condition=False)
])

######################################

# SIGN IN CLOSE FUNCTION ################

CELLS_CONDITION_SIGN_IN_CLOSE_FUNCTION = CellsConditions(conditions=[
    ConditionAtLeastOneCellAmongList(cells=[
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="E26"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="E27"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="E28"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="E29"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="E30"),
        BoxToCheck(sheet_name=SheetName.SHEET_6.value, cell_address="E31")
    ],
                                     is_parent_condition=False)
])

######################################

# RAPPORT RAPPORT J ################

CELLS_CONDITION_RAPPORT_J = CellsConditions(conditions=[
    ConditionNcAllJChoosed(cell=BoxToCheck(sheet_name=SheetName.SHEET_5.value,
                                           cell_address="J4"),
                           is_parent_condition=False)
])

######################################

# RAPPORT DESCRIPTION REPORT ################

CELLS_CONDITION_DROPDOWN_REPORT = CellsConditions(conditions=[
    ConditionCheckAllSheetDropDown(sheet_name=SheetName.SHEET_5.value,
                                   is_parent_condition=False)
])

######################################

# RAPPORT DROPDOWN REPORT ################

CELLS_CONDITION_DESCRIPTION_REPORT = CellsConditions(conditions=[
    ConditionCheckAllSheetDescription(sheet_name=SheetName.SHEET_5.value,
                                      is_parent_condition=False)
])

######################################

# RAPPORT REF REPORT ################

CELLS_CONDITION_REF_REPORT = CellsConditions(conditions=[
    ConditionCheckAllSheetReference(sheet_name=SheetName.SHEET_5.value,
                                    is_parent_condition=False)
])

######################################

CELLS_CONDITIONS = [
    CELLS_CONDITION_OPAC_DATE_AUDIT, CELLS_CONDITION_OPAC_DURATION,
    CELLS_CONDITION_OPAC_AUDIT_TYPE, CELLS_CONDITION_OPAC_COMPANY_NAME,
    CELLS_CONDITION_OPAC_COMPANY_ADRESS, CELLS_CONDITION_OPAC_COMPANY_CP,
    CELLS_CONDITION_OPAC_COMPANY_CITY, CELLS_CONDITION_OPAC_COMPANY_SIREN,
    CELLS_CONDITION_OPAC_COMPANY_NDA, CELLS_CONDITION_OPAC_RP_NAME,
    CELLS_CONDITION_OPAC_RP_ROLE, CELLS_CONDITION_OPAC_RP_PHONE,
    CELLS_CONDITION_OPAC_RP_MAIL, CELLS_CONDITION_OPAC_DISTANCE,
    CELLS_CONDITION_OPAC_DISTANCE, CELLS_CONDITION_OPAC_PERIOD,
    CELLS_CONDITION_OPAC_MORE_TWO_DAY, CELLS_CONDITION_OPAC_RNCP,
    CELLS_CONDITION_OPAC_RS, CELLS_CONDITION_SUB_CONTRACTOR,
    CELLS_CONDITION_SUB_CONTRACTOR_2, CELLS_CONDITION_PSH,
    CELLS_CONDITION_CONTRACTOR_100, CELLS_CONDITION_OPAC_DESCRIPTION,
    CELLS_CONDITION_OPAC_AF_FOLDER, CELLS_CONDITION_OPAC_AF_JUSTIFY,
    CELLS_CONDITION_OPAC_VAE_FOLDER, CELLS_CONDITION_OPAC_VAE_JUSTIFY,
    CELLS_CONDITION_OPAC_BC_FOLDER, CELLS_CONDITION_OPAC_BC_JUSTIFY,
    CELLS_CONDITION_OPAC_FA_FOLDER, CELLS_CONDITION_OPAC_FA_JUSTIFY,
    CELLS_CONDITION_OPAC_AUDITOR_NAME, CELLS_CONDITION_OPAC_AUDITOR_PHONE,
    CELLS_CONDITION_OPAC_AUDITOR_MAIL, CELLS_CONDITION_OPAC_REMOTE_CONDITION,
    CELLS_CONDITION_OPAC_REMOTE_GOAL, CELLS_CONDITION_OPAC_REMOTE_TOOLS,
    CELLS_CONDITION_OPAC_ACTION, CELLS_CONDITION_OPAC_AUDIT_TYPE_2,
    CELLS_CONDITION_CNEFOP, CELLS_CONDITION_WATCHING, CELLS_CONDITION_NC,
    CELLS_CONDITION_OPAC_NC_ID, CELLS_CONDITION_OPAC_NC_PREVIOUS,
    CELLS_CONDITION_OPAC_NC_CORRECTION, CELLS_CONDITION_OPAC_NC_SHEET,
    CELLS_CONDITION_OPAC_NC_MAJ, CELLS_CONDITION_OPAC_NC_SHEET_NUMBER,
    CELLS_CONDITION_OPAC_DATE_SEND, CELLS_CONDITION_OPAC_DATE_SEND_NC,
    CELLS_CONDITION_OPAC_DATE_SEND_RAPPORT, CELLS_CONDITION_OPAC_PASS,
    CELLS_CONDITION_OPAC_SIGNATURE, CELLS_CONDITION_PLAN_DATE_1,
    CELLS_CONDITION_PLAN_DATE_3, CELLS_CONDITION_PLAN_DATE_4,
    CELLS_CONDITION_PLAN_DATE_5, CELLS_CONDITION_PLAN_DATE_6,
    CELLS_CONDITION_PLAN_DATE_7, CELLS_CONDITION_PLAN_DATE_8,
    CELLS_CONDITION_PLAN_DATE_9, CELLS_CONDITION_PLAN_DATE_10,
    CELLS_CONDITION_PLAN_DATE_11, CELLS_CONDITION_PLAN_DATE_13,
    CELLS_CONDITION_PLAN_DATE_14, CELLS_CONDITION_SIGN_IN_OPEN_NAME,
    CELLS_CONDITION_SIGN_IN_OPEN_LASTNAME,
    CELLS_CONDITION_SIGN_IN_OPEN_FUNCTION, CELLS_CONDITION_SIGN_IN_CLOSE_NAME,
    CELLS_CONDITION_SIGN_IN_CLOSE_LASTNAME,
    CELLS_CONDITION_SIGN_IN_CLOSE_FUNCTION, CELLS_CONDITION_RAPPORT_J,
    CELLS_CONDITION_DROPDOWN_REPORT, CELLS_CONDITION_DESCRIPTION_REPORT,
    CELLS_CONDITION_REF_REPORT
]
