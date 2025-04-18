# logger_config.py
import logging
from logging.handlers import TimedRotatingFileHandler
from utils.logging_format import CustomFormatter
import os

# Set up the main logger
logger = logging.getLogger('log')

# File handler with monthly rotation
fh = TimedRotatingFileHandler('/logs/logs.log', when='W0', interval=1, backupCount=3)
plain_formatter = logging.Formatter("%(asctime)s - %(name)s - %(filename)s - %(levelname)s - Line: %(lineno)d - %(message)s")
fh.setFormatter(plain_formatter)
logger.addHandler(fh)

ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
ch.setFormatter(CustomFormatter())
logger.addHandler(ch)

logger.setLevel(logging.DEBUG)

logger.debug('Enabled')
logger.info('Enabled')
logger.warning('Enabled')
logger.error('Enabled')
logger.critical('Enabled')
