# logger_config.py
import logging
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path

from utils.logging_format import CustomFormatter

# Logs live in a project-relative logs/ directory (created if missing), not an absolute /logs path.
LOG_DIR = Path(__file__).resolve().parents[1] / "logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "logs.log"

logger = logging.getLogger("log")

# Guard against duplicate handlers if this module is imported more than once.
if not logger.handlers:
    # Weekly rotation (Mondays), keep 3 backups.
    fh = TimedRotatingFileHandler(LOG_FILE, when="W0", interval=1, backupCount=3, encoding="utf-8")
    fh.setFormatter(
        logging.Formatter(
            "%(asctime)s - %(name)s - %(filename)s - %(levelname)s - Line: %(lineno)d - %(message)s"
        )
    )
    logger.addHandler(fh)

    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(CustomFormatter())
    logger.addHandler(ch)

    logger.setLevel(logging.DEBUG)
