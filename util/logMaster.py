import logging
import os
from datetime import datetime
from dotenv import load_dotenv

class Logger:
    def __init__(self, name, level=logging.INFO):
        load_dotenv()

        log_path = os.getenv('LOG_PATH')
        log_file = log_path + f"DCM_{datetime.now().strftime('%Y%m%d')}.log"

        self.logger = logging.getLogger(name)
        if not self.logger.handlers:
            self.logger.setLevel(level)
            formatter = logging.Formatter('[%(asctime)s][%(name)s][%(levelname)s]:::   %(message)s', '%Y-%m-%d::%H:%M:%S')

            file_handler = logging.FileHandler(log_file)
            file_handler.setFormatter(formatter)
            self.logger.addHandler(file_handler)

    def debug(self, message):
        self.logger.debug(message)

    def info(self, message):
        self.logger.info(message)

    def warning(self, message):
        self.logger.warning(message)

    def error(self, message):
        self.logger.error(message)

    def line(self):
        self.logger.info("================================================================")

