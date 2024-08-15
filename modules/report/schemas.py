"""_summary_"""

from typing import List, Optional

from pydantic import BaseModel

from modules.condition.schemas import CellsConditionState


class UIReportCell(BaseModel):
    """_summary_

    Args:
        BaseModel (_type_): _description_
    """

    state: CellsConditionState

    instruction: str

    sheet_names: List[str]

    cell_adress: Optional[str] = None
