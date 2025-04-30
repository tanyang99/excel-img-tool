import logging

from PyQt6.QtCore import QObject, pyqtSignal

# 日志配置
LOG_LEVEL = logging.INFO
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'


class CustomHandler(logging.Handler, QObject):
    """
    自定义日志处理器，结合 logging 和 PyQt6 的信号槽机制。
    将日志消息通过信号发射到 GUI。
    """
    log_message = pyqtSignal(str)
    MAX_LOG_LINES = 1000

    def __init__(self, max_lines: int = MAX_LOG_LINES, level: int = LOG_LEVEL, format_str: str = LOG_FORMAT):
        """
        初始化自定义日志处理器。

        :param max_lines: 最大日志行数，默认为 1000。
        :param level: 日志级别，默认为 INFO。
        :param format_str: 日志格式字符串，默认为标准格式。
        """
        logging.Handler.__init__(self)
        QObject.__init__(self)
        self.max_lines = max(1, max_lines)  # 确保 max_lines 至少为 1
        self.level = self._validate_log_level(level)
        self.format_str = format_str if format_str else self.LOG_FORMAT

    def _validate_log_level(self, level: int) -> int:
        """
        验证并返回有效的日志级别。

        :param level: 要验证的日志级别。
        :return: 有效的日志级别。
        """
        valid_levels = {logging.DEBUG, logging.INFO, logging.WARNING, logging.ERROR, logging.CRITICAL}
        return level if level in valid_levels else self.LOG_LEVEL

    def configure(self) -> None:
        """
        配置处理器并将其添加到根日志器。
        """
        try:
            formatter = logging.Formatter(self.format_str)
            self.setFormatter(formatter)
            self.setLevel(self.level)
            root_logger = logging.getLogger()
            if not any(isinstance(h, CustomHandler) for h in root_logger.handlers):
                root_logger.addHandler(self)
                root_logger.setLevel(self.level)
                logging.debug("自定义日志处理器已添加到根日志器。")
        except Exception as e:
            logging.error(f"配置日志失败: {e}")

    def emit(self, record):
        """
        通过 log_message 信号发射日志记录。

        :param record: 日志记录对象。
        """
        try:
            msg = self.format(record)
            self.log_message.emit(msg)
        except Exception as e:
            logging.error(f"发射日志消息失败: {e}")

    def close(self):
        """
        在清理时从日志器中移除处理器。
        """
        try:
            logging.getLogger().removeHandler(self)
            logging.debug("自定义日志处理器已移除。")
        except Exception as e:
            logging.error(f"关闭日志处理器失败: {e}")
        super().close()
