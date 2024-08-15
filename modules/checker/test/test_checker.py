"""_summary_"""

from modules.checker.checker import Checker

from modules.condition import CELLS_CONDITIONS

checker = Checker(cells_conditions=CELLS_CONDITIONS)

checker.check_cells_conditions()
