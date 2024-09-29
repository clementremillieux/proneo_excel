"""_summary_"""

import time

from datetime import datetime as dt

from typing import Dict, Optional

from pydantic import BaseModel


class ExecutionPerformances(BaseModel):
    """_summary_

    Args:
        BaseModel (_type_): _description_
    """

    start_time: Optional[str] = None

    end_time: Optional[str] = None

    duration: Optional[float] = None

    def start(self):
        """_summary_"""

        self.start_time = dt.now().strftime("%Y-%m-%d %H:%M:%S")

        self.duration = time.time()

    def stop(self):
        """_summary_"""

        self.end_time = dt.now().strftime("%Y-%m-%d %H:%M:%S")

        if self.duration:
            self.duration = round(time.time() - self.duration, 1)

        else:
            self.duration = 0


class TimeExecutionCounter(BaseModel):
    """_summary_

    Args:
        BaseModel (_type_): _description_
    """

    executions_performances: Dict[str, ExecutionPerformances]
