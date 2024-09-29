"""logger_config.py"""

import os

import logging


class CustomFormatter(logging.Formatter):
    """Custom logger with colored output, aligned messages, and relative paths"""

    RED = "\033[31m"
    GREEN = "\033[92m"
    RESET = "\033[0m"
    ORANGE = "\033[41m"

    def format(self, record):
        """Format log records with color, alignment, and additional context"""

        original = super(CustomFormatter, self).format(record)

        if record.levelno == logging.INFO:
            level_name = self.GREEN + record.levelname + self.RESET

        elif record.levelno == logging.ERROR:
            level_name = self.RED + record.levelname + self.RESET

        elif record.levelno == logging.WARNING:
            level_name = self.ORANGE + record.levelname + self.RESET

        else:
            level_name = record.levelname

        project_root = os.path.abspath(
            os.path.join(os.path.dirname(__file__), ".."))

        relative_path = os.path.relpath(record.pathname, project_root)

        module = record.module

        function_name = record.funcName

        line_no = record.lineno

        dynamic_part = f"{level_name} -> [{relative_path}.{module}.{function_name}.{line_no}] :"

        max_dynamic_length = 100

        padding = max(0, max_dynamic_length - len(dynamic_part))

        padded_original = original.rjust(padding + len(original))

        timestamp = self.formatTime(record, "%Y-%m-%d %H:%M:%S")

        return f"{timestamp} {dynamic_part}\t{padded_original}"


def setup_logging():
    """setup the logging logger"""

    logger = logging.getLogger("Proneo app")

    if logger.hasHandlers():
        logger.handlers.clear()

    handler = logging.StreamHandler()

    handler.setFormatter(CustomFormatter("%(message)s"))

    logger.setLevel(logging.INFO)

    logger.addHandler(handler)

    logger.propagate = False


setup_logging()

logger = logging.getLogger("Proneo app")
