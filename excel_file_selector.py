import os
from PyQt6.QtWidgets import QWidget, QVBoxLayout, QPushButton, QTreeWidget, QTreeWidgetItem, QFileDialog, \
    QMessageBox, QTextEdit
from custom_log_config import CustomHandler
from excel_image_embedder import ExcelImageEmbedder
import logging


class ExcelFileSelector(QWidget):
    def __init__(self):
        super().__init__()
        self.browse_button = None
        self.file_tree = None
        self.process_images_button = None
        self.selected_file_paths = []
        self.log_text_edit = None
        self.custom_log_handler = None
        self.initUI()
        self.setup_logging()

    def initUI(self):
        # 设置主窗体标题和大小
        self.setWindowTitle("Excel 文件选择器")
        self.setGeometry(300, 200, 1000, 800)

        layout = QVBoxLayout()

        # 创建浏览按钮
        self.browse_button = QPushButton("选择 Excel 文件", self)
        self.browse_button.setFixedSize(200, 50)
        self.browse_button.clicked.connect(self.browse_files)
        layout.addWidget(self.browse_button)

        # 创建文件树
        self.file_tree = QTreeWidget(self)
        self.file_tree.setHeaderLabels(["文件信息"])
        self.file_tree.setSelectionMode(QTreeWidget.SelectionMode.ExtendedSelection)

        layout.addWidget(self.file_tree, stretch=0.3)

        # 创建图片处理按钮
        self.process_images_button = QPushButton("处理图片", self)
        self.process_images_button.clicked.connect(self.process_selected_sheets)
        layout.addWidget(self.process_images_button)

        # 创建日志文本框
        self.log_text_edit = QTextEdit(self)
        self.log_text_edit.setReadOnly(True)

        layout.addWidget(self.log_text_edit)

        self.setLayout(layout)

    def setup_logging(self):
        # Create the custom handler (no need for the text_edit in the constructor anymore)
        self.custom_log_handler = CustomHandler()
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        self.custom_log_handler.setFormatter(formatter)

        # Get the root logger and add the custom handler
        # Ensure you don't add multiple handlers if setup_logging is called more than once
        root_logger = logging.getLogger()
        if not any(isinstance(h, CustomHandler) for h in root_logger.handlers):
            root_logger.addHandler(self.custom_log_handler)
            # Set a minimum logging level, e.g., INFO
            root_logger.setLevel(logging.INFO)

        # Connect the handler's signal to a slot in this class
        self.custom_log_handler.log_message.connect(self.append_log_message)

    def append_log_message(self, message):
        """Slot to receive log messages and append to the text edit."""
        # Append the message
        self.log_text_edit.append(message)
        # Control maximum lines (optional, can also be done in handler's signal method if preferred)
        max_lines = self.custom_log_handler.max_lines  # Use the limit from the handler
        while self.log_text_edit.document().lineCount() > max_lines:
            cursor = self.log_text_edit.textCursor()
            cursor.movePosition(cursor.MoveOperation.Start)
            cursor.select(cursor.SelectionType.LineUnderCursor)
            cursor.removeSelectedText()
            # Need to explicitly remove the newline character left behind
            cursor.deleteChar()

    def browse_files(self):
        try:
            logging.info("正在打开文件选择对话框...")
            initial_dir = os.getcwd()
            file_paths, _ = QFileDialog.getOpenFileNames(
                self,
                "选择 Excel 文件",
                initial_dir,
                "Excel 文件 (*.xlsx *.xls);;所有文件 (*.*)"
            )

            if file_paths:
                logging.info(f"选择了 {len(file_paths)} 个文件.")
                if not ExcelImageEmbedder.check_file_count_and_size(file_paths):
                    logging.warning("文件数量或大小检查失败，请检查日志。")
                    return
                logging.info("文件数量和大小检查通过。")
                self.selected_file_paths = file_paths
                file_sheet_info = ExcelImageEmbedder.get_file_and_sheet_info(file_paths)

                self.file_tree.clear()
                for file_name, sheet_info in file_sheet_info.items():
                    root_item = QTreeWidgetItem(self.file_tree, [file_name])
                    for sheet_index, sheet_name in sheet_info:
                        child_item = QTreeWidgetItem(root_item, [f"sheet：{sheet_name} (Index: {sheet_index})"])
                        root_item.addChild(child_item)
                self.file_tree.expandAll()
                logging.info("文件和 sheet 信息已加载到文件树。")

                # 默认选择第一个文件的第一个 sheet
                top_level_item = self.file_tree.topLevelItem(0)
                if top_level_item and top_level_item.childCount() > 0:
                    first_sheet_item = top_level_item.child(0)
                    first_sheet_item.setSelected(True)
                    logging.info("默认选中第一个文件的第一个 sheet。")
            else:
                logging.info("未选择任何文件。")
        except Exception as e:
            logging.error(f"浏览文件时发生错误: {e}")
            # Show an error message box as well for critical errors
            QMessageBox.critical(self, "错误", f"浏览文件时发生错误: {e}")

    def process_selected_sheets(self):
        try:
            logging.info("开始处理选中的 sheet...")
            selected_items = self.file_tree.selectedItems()
            if not selected_items:
                logging.warning("用户未选择任何要处理的 sheet。")
                QMessageBox.warning(self, "警告", "请选择要处理的 sheet。")
                return

            # Disable button to prevent re-entry
            self.process_images_button.setEnabled(False)

            file_sheet_map = self.get_file_sheet_map(selected_items)
            logging.info(f"将处理以下文件和 sheet: {file_sheet_map}")

            # Assuming ExcelImageEmbedder().embed_images can handle logging internally
            # and that these logs will go through the root logger to our handler.
            # If embed_images runs in a *separate thread*, the Signal/Slot mechanism is crucial.
            # If it blocks the UI thread, logs will appear as they are generated within embed_images calls.
            for file_path, info in file_sheet_map.items():
                file_name = info["file_name"]
                sheet_indices = info["sheet_indices"]
                logging.info(f"正在处理文件: {file_name}, sheet 索引: {sheet_indices}")
                embedder = ExcelImageEmbedder()

                embedder.embed_images([file_path], {file_name: sheet_indices})
                logging.info(f"文件 {file_name} 处理完成。")

            logging.info("所有选中的文件处理完成。")
            QMessageBox.information(self, "成功", "文件处理完成")

        except Exception as e:
            logging.error(f"处理图片时发生错误: {e}")
            QMessageBox.critical(self, "错误", f"处理图片时发生错误: {e}")
        finally:
            # Always re-enable the button
            self.process_images_button.setEnabled(True)
            logging.info("处理按钮已启用。")

    def add_sheet_indices(self, item, file_path, file_name, file_sheet_map):
        # Helper to recursively add selected sheet indices
        if item.childCount() == 0:  # It's a sheet item
            # Check if the item is actually selected before processing
            if item.isSelected():
                try:
                    # Extract sheet index from the item text
                    text = item.text(0)
                    start_index = text.find("(Index: ") + len("(Index: ")
                    end_index = text.find(")")
                    if start_index != -1 and end_index != -1:
                        sheet_index_str = text[start_index:end_index]
                        sheet_index = int(sheet_index_str)

                        # Ensure the file_path entry exists
                        if file_path not in file_sheet_map:
                            file_sheet_map[file_path] = {
                                "file_name": file_name,
                                "sheet_indices": []
                            }
                        if sheet_index not in file_sheet_map[file_path]["sheet_indices"]:
                            file_sheet_map[file_path]["sheet_indices"].append(sheet_index)
                    else:
                        logging.warning(f"无法解析 sheet index: {text}")
                except ValueError as ve:
                    logging.error(f"无法将 sheet 索引转换为整数: {item.text(0)}. 错误: {ve}")
                except Exception as e:
                    logging.error(f"处理 sheet 索引时发生未知错误: {item.text(0)}. 错误: {e}")

        else:  # It's a file item or a non-leaf selected item
            # Recursively check children *only if the parent is selected* or if children are individually selected
            # The get_file_sheet_map loop handles starting with selected items.
            # This function should only be called on selected items or their descendants.
            for i in range(item.childCount()):
                child = item.child(i)
                # Recursively call for children
                self.add_sheet_indices(child, file_path, file_name, file_sheet_map)

    def get_file_sheet_map(self, selected_items):
        """
        Builds a dictionary mapping file paths to their selected sheet indices.
        Correctly handles selecting parent file nodes (includes all sheets)
        and selecting individual sheet nodes.
        """
        file_sheet_map = {}

        # Iterate through the initially selected items from the tree
        for item in selected_items:
            # Find the top-level file item for the current selected item
            file_item = item
            while file_item.parent() is not None:
                file_item = file_item.parent()

            file_name = file_item.text(0)
            file_path = next(
                (path for path in self.selected_file_paths if os.path.basename(path) == file_name), None)

            if not file_path:
                logging.warning(f"找不到与文件树项 '{file_name}' 对应的完整文件路径。")
                continue

            # Ensure the dictionary entry for this file path exists
            if file_path not in file_sheet_map:
                file_sheet_map[file_path] = {
                    "file_name": file_name,
                    "sheet_indices": []
                }

            # --- 关键判断和处理父节点在这里 ---
            if item.parent() is None:  # **如果最初选中的 item 就是一个顶层文件节点**
                logging.debug(f"文件 {file_name} (父节点) 被选中，添加所有子 sheet。")
                # 如果是文件节点被选中，则 **遍历该文件节点下的所有子节点 (sheet)**
                # 注意这里是 file_item.childCount()，遍历的是所有子节点
                for i in range(file_item.childCount()):
                    sheet_item = file_item.child(i)
                    # 并调用辅助函数将每个子 sheet 的索引添加到 map 中
                    # _extract_and_add_sheet_index 不再检查 item.isSelected()
                    self._extract_and_add_sheet_index(sheet_item, file_path, file_name, file_sheet_map)

            else:  # 如果最初选中的 item 是一个子 sheet 节点
                # ... (这部分处理选中子节点，逻辑不变)
                sheet_item = item
                logging.debug(f"Sheet '{sheet_item.text(0)}' (子节点) 被选中。")
                # 如果是 sheet 节点被选中，则只将这个单独的 sheet 索引添加到 map 中
                self._extract_and_add_sheet_index(sheet_item, file_path, file_name, file_sheet_map)

        return file_sheet_map

    def get_selected_file_paths(self):
        # This method seems unused in the provided code, but keeping it.
        return self.selected_file_paths

    def closeEvent(self, event):
        logging.info("应用正在关闭。")
        event.accept()
