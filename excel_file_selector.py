import os
import platform
from typing import List, Dict, Optional
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QPushButton, QTreeWidget, QTreeWidgetItem, QFileDialog, QMessageBox, \
    QTextEdit
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from custom_log_config import CustomHandler
from excel_image_embedder import ExcelImageEmbedder
import logging
from openpyxl.utils.exceptions import InvalidFileException

# 常量
WINDOW_TITLE = "Excel 文件选择器"
WINDOW_GEOMETRY = (300, 200, 1000, 800)
BROWSE_BUTTON_SIZE = (200, 50)
PROCESS_BUTTON_SIZE = (200, 50)
FILE_FILTER = "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)"


class Worker(QThread):
    """处理 Excel 文件的工作线程。"""
    progress = pyqtSignal(str)
    error = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, file_sheet_map: Dict, embedder_class):
        super().__init__()
        self.file_sheet_map = file_sheet_map
        self.embedder_class = embedder_class

    def run(self):
        """在工作线程中运行图片嵌入过程。"""
        try:
            logging.debug(f"Worker 线程启动，file_sheet_map: {self.file_sheet_map}")
            if not self.file_sheet_map:
                logging.error("file_sheet_map 为空，无法处理文件。")
                self.error.emit("错误：未选择任何文件或 sheet。")
                self.finished.emit()
                return

            for file_path, info in self.file_sheet_map.items():
                file_name = info.get("file_name", "")
                sheet_indices = info.get("sheet_indices", [])
                logging.debug(f"处理文件: {file_name}, 路径: {file_path}, sheet 索引: {sheet_indices}")

                # 验证文件路径
                if not os.path.exists(file_path):
                    logging.error(f"文件 {file_path} 不存在。")
                    self.error.emit(f"错误：文件 {file_name} 不存在。")
                    continue

                # 验证 sheet 索引
                if not sheet_indices:
                    logging.warning(f"文件 {file_name} 未选择任何 sheet，跳过。")
                    self.progress.emit(f"文件 {file_name} 未选择 sheet，跳过。")
                    continue

                self.progress.emit(f"正在处理文件: {file_name}, sheet 索引: {sheet_indices}")
                embedder = self.embedder_class()
                try:
                    logging.debug(f"调用 embed_images for {file_name}")
                    embedder.embed_images(
                        [file_path],
                        {file_name: sheet_indices},
                        progress_callback=self.progress.emit
                    )
                    self.progress.emit(f"文件 {file_name} 处理完成。")
                except Exception as e:
                    logging.error(f"embed_images 处理 {file_name} 失败: {str(e)}", exc_info=True)
                    self.error.emit(f"处理文件 {file_name} 时发生错误: {str(e)}")
                    continue

            self.progress.emit("所有选中的文件处理完成。")
            self.finished.emit()
        except (OSError, InvalidFileException) as e:
            logging.error(f"处理图片时发生系统错误: {str(e)}", exc_info=True)
            self.error.emit(f"处理图片时发生错误: {str(e)}")
            self.finished.emit()
        except Exception as e:
            logging.error(f"处理图片时发生意外错误: {str(e)}", exc_info=True)
            self.error.emit(f"处理图片时发生意外错误: {str(e)}")
            self.finished.emit()


class ExcelFileSelector(QWidget):
    def __init__(self, embedder_class=None):
        super().__init__()
        self.selected_file_paths: List[str] = []
        self.browse_button: Optional[QPushButton] = None
        self.file_tree: Optional[QTreeWidget] = None
        self.process_images_button: Optional[QPushButton] = None
        self.log_text_edit: Optional[QTextEdit] = None
        self.custom_log_handler: Optional[CustomHandler] = None
        self.worker: Optional[Worker] = None
        self.embedder_class = embedder_class or ExcelImageEmbedder
        self.init_ui()
        self.setup_logging()

    def init_ui(self) -> None:
        """初始化用户界面。"""
        self.setWindowTitle(WINDOW_TITLE)
        self.setGeometry(*WINDOW_GEOMETRY)
        layout = QVBoxLayout()
        self.browse_button = QPushButton("选择 Excel 文件", self)
        self.browse_button.setFixedSize(*BROWSE_BUTTON_SIZE)
        self.browse_button.clicked.connect(self.browse_files)
        layout.addWidget(self.browse_button)
        self.file_tree = QTreeWidget(self)
        self.file_tree.setHeaderLabels(["文件信息"])
        self.file_tree.setSelectionMode(QTreeWidget.SelectionMode.ExtendedSelection)
        layout.addWidget(self.file_tree, stretch=0.3)
        self.process_images_button = QPushButton("处理图片", self)
        self.process_images_button.setFixedSize(*PROCESS_BUTTON_SIZE)
        self.process_images_button.clicked.connect(self.process_selected_sheets)
        layout.addWidget(self.process_images_button)
        self.log_text_edit = QTextEdit(self)
        self.log_text_edit.setReadOnly(True)
        layout.addWidget(self.log_text_edit)
        self.setLayout(layout)

    def setup_logging(self) -> None:
        """使用自定义处理器配置日志记录。"""
        self.custom_log_handler = CustomHandler()
        self.custom_log_handler.configure()
        self.custom_log_handler.log_message.connect(self.append_log_message)

    def append_log_message(self, message: str) -> None:
        """将日志消息追加到文本编辑框。"""
        try:
            self.log_text_edit.append(message)
            while self.log_text_edit.document().lineCount() > self.custom_log_handler.max_lines:
                self.log_text_edit.setPlainText('\n'.join(self.log_text_edit.toPlainText().split('\n')[1:]))
            self.log_text_edit.ensureCursorVisible()
        except Exception as e:
            logging.error(f"追加日志消息失败: {e}")

    def browse_files(self) -> None:
        """打开文件对话框选择 Excel 文件并填充文件树。"""
        try:
            logging.info("正在打开文件选择对话框...")
            initial_dir = os.getcwd()
            options = QFileDialog.Option.DontUseNativeDialog if platform.system() == "Darwin" else QFileDialog.Option(0)
            file_paths, _ = QFileDialog.getOpenFileNames(
                self,
                "选择 Excel 文件",
                initial_dir,
                FILE_FILTER,
                options=options
            )
            if not file_paths:
                logging.info("未选择任何文件。")
                return

            logging.info(f"选择了 {len(file_paths)} 个文件。")
            if not self.embedder_class.check_file_count_and_size(file_paths):
                logging.warning("文件数量或大小检查失败，请检查日志。")
                return

            self.selected_file_paths = file_paths
            file_sheet_info = self.embedder_class.get_file_and_sheet_info(file_paths)
            self.file_tree.clear()
            for file_path in file_paths:
                file_name = os.path.basename(file_path)
                root_item = QTreeWidgetItem(self.file_tree, [file_name])
                root_item.setData(0, Qt.ItemDataRole.UserRole, file_path)
                for sheet_index, sheet_name in file_sheet_info.get(file_name, []):
                    child_item = QTreeWidgetItem(root_item, [f"sheet：{sheet_name} (Index: {sheet_index})"])
                    root_item.addChild(child_item)
            self.file_tree.expandAll()
            logging.info("文件和 sheet 信息已加载到文件树。")

        except Exception as e:
            logging.error(f"浏览文件时发生错误: {e}")
            QMessageBox.critical(self, "错误", f"浏览文件时发生错误: {e}")

    def process_selected_sheets(self) -> None:
        """在单独线程中处理选中的 sheet 以嵌入图片。"""
        try:
            selected_items = self.file_tree.selectedItems()
            if not selected_items:
                logging.warning("用户未选择任何要处理的 sheet。")
                QMessageBox.warning(self, "警告", "请选择要处理的 sheet。")
                return

            self.process_images_button.setEnabled(False)
            file_sheet_map = self.get_file_sheet_map(selected_items)
            if not file_sheet_map:
                logging.warning("没有有效的文件或 sheet 被选中。")
                QMessageBox.warning(self, "警告", "没有有效的文件或 sheet 被选中。")
                self.process_images_button.setEnabled(True)
                return

            logging.info(f"将处理以下文件和 sheet: {file_sheet_map}")
            self.worker = Worker(file_sheet_map, self.embedder_class)
            self.worker.progress.connect(self.append_log_message)
            self.worker.error.connect(self.handle_worker_error)
            self.worker.finished.connect(self.handle_worker_finished)
            self.worker.start()

        except Exception as e:
            logging.error(f"启动处理线程时发生错误: {e}")
            QMessageBox.critical(self, "错误", f"启动处理线程时发生错误: {e}")
            self.process_images_button.setEnabled(True)

    def handle_worker_error(self, error_msg: str) -> None:
        """处理工作线程中的错误。"""
        logging.error(error_msg)
        QMessageBox.critical(self, "错误", error_msg)

    def handle_worker_finished(self) -> None:
        """处理工作线程完成。"""
        self.process_images_button.setEnabled(True)
        logging.info("处理按钮已启用。")
        QMessageBox.information(self, "成功", "文件处理完成")
        self.worker = None

    def _parse_sheet_index(self, text: str) -> Optional[int]:
        """从 QTreeWidgetItem 文本中解析 sheet 索引。"""
        try:
            start_index = text.find("(Index: ") + len("(Index: ")
            end_index = text.find(")")
            if start_index == -1 or end_index == -1:
                logging.warning(f"无法解析 sheet 索引: {text}")
                return None
            return int(text[start_index:end_index])
        except (ValueError, IndexError) as e:
            logging.error(f"解析 sheet 索引失败: {text}. 错误: {e}")
            return None

    def add_sheet_indices(self, item: QTreeWidgetItem, file_path: str, file_name: str,
                          file_sheet_map: Dict, check_selection: bool = True) -> None:
        """
        将 sheet 索引添加到 file_sheet_map。
        :param item: QTreeWidgetItem (文件或 sheet)
        :param file_path: 文件完整路径
        :param file_name: 文件名
        :param file_sheet_map: 存储文件到 sheet 映射的字典
        :param check_selection: 是否检查项目是否被选中
        """
        if item.childCount() == 0:  # Sheet 项目
            if not check_selection or item.isSelected():
                sheet_index = self._parse_sheet_index(item.text(0))
                if sheet_index is not None:
                    if file_path not in file_sheet_map:
                        file_sheet_map[file_path] = {"file_name": file_name, "sheet_indices": []}
                    if sheet_index not in file_sheet_map[file_path]["sheet_indices"]:
                        file_sheet_map[file_path]["sheet_indices"].append(sheet_index)
        else:  # 文件项目
            for i in range(item.childCount()):
                self.add_sheet_indices(item.child(i), file_path, file_name, file_sheet_map, check_selection)

    def get_file_sheet_map(self, selected_items: List[QTreeWidgetItem]) -> Dict:
        """
        构建文件路径到其选中 sheet 索引的字典。
        :param selected_items: 选中的 QTreeWidgetItem 列表
        :return: 文件路径到文件名和 sheet 索引的字典
        """
        file_sheet_map = {}
        for item in selected_items:
            file_item = item
            while file_item.parent() is not None:
                file_item = file_item.parent()
            file_path = file_item.data(0, Qt.ItemDataRole.UserRole)
            file_name = os.path.basename(file_path)
            if not file_path:
                logging.warning(f"找不到与文件树项 '{file_name}' 对应的完整文件路径。")
                continue

            if file_path not in file_sheet_map:
                file_sheet_map[file_path] = {"file_name": file_name, "sheet_indices": []}

            if item.parent() is None:  # 选择了文件节点
                logging.debug(f"文件 {file_name} (父节点) 被选中，添加所有子 sheet。")
                for i in range(file_item.childCount()):
                    self.add_sheet_indices(file_item.child(i), file_path, file_name, file_sheet_map,
                                           check_selection=False)
            else:  # 选择了 sheet 节点
                logging.debug(f"Sheet '{item.text(0)}' (子节点) 被选中。")
                self.add_sheet_indices(item, file_path, file_name, file_sheet_map, check_selection=True)

        return file_sheet_map

    def get_selected_file_paths(self) -> List[str]:
        """返回选中的文件路径列表。"""
        return self.selected_file_paths

    def closeEvent(self, event) -> None:
        """处理窗口关闭事件。"""
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
        if self.custom_log_handler:
            self.custom_log_handler.close()
        logging.info("应用正在关闭。")
        event.accept()
