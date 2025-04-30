import logging
from PyQt6.QtCore import QObject, pyqtSignal


class CustomHandler(logging.Handler, QObject):
    log_message = pyqtSignal(str)

    def __init__(self, max_lines=1000):
        logging.Handler.__init__(self)
        QObject.__init__(self)
        self.max_lines = max_lines

    def emit(self, record):
        msg = self.format(record)
        self.log_message.emit(msg)
