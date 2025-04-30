# embed_worker.py
from PyQt6.QtCore import QObject, pyqtSignal
from excel_image_embedder import ExcelImageEmbedder
import logging


class EmbedWorker(QObject):
    # 定义信号，用于通知主线程任务完成或发生错误
    finished = pyqtSignal()
    error = pyqtSignal(str)

    # 如果需要更详细的进度，可以在这里添加其他信号，例如 progress_step = pyqtSignal()

    def __init__(self, file_sheet_map, parent=None):
        super().__init__(parent)
        self._file_sheet_map = file_sheet_map
        # 在工作者线程中创建 Embedder 实例
        self._embedder = ExcelImageEmbedder()

    def run(self):
        """
        这是在 QThread 中执行的实际任务方法。
        这个方法不应直接与 GUI 交互。
        """
        logging.info("工作者线程开始执行嵌入任务...")
        try:
            # 遍历文件和 sheet 地图，执行嵌入操作
            # self._file_sheet_map 包含了所有要处理的文件和 sheet 信息
            # embedder.embed_images 会处理整个 map
            self._embedder.embed_images(
                list(self._file_sheet_map.keys()),  # 传递文件路径列表
                self._file_sheet_map  # 传递文件到 sheet 索引的映射
            )

            logging.info("工作者线程完成所有嵌入任务。")
            self.finished.emit()  # 任务成功完成，发出 finished 信号

        except Exception as e:
            logging.error(f"工作者线程中发生错误: {e}")
            self.error.emit(str(e))  # 发生错误，发出 error 信号并传递错误信息
