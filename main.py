import sys
from PyQt6.QtWidgets import QApplication  # 导入 QApplication

from excel_file_selector import ExcelFileSelector  # 从您的文件导入主窗口类


# 确保 ExcelFileSelector.py, excel_image_embedder.py, custom_log_config.py
# embed_worker.py (如果分开写了) 都存在于同一个目录下，或者在 Python 的搜索路径中。

def main():
    """
    应用程序的主入口点。
    """
    # 1. 创建 QApplication 实例
    # sys.argv 是命令行参数列表，通常需要传递给 QApplication
    app = QApplication(sys.argv)

    # 2. 创建主窗口实例
    # ExcelFileSelector 类的 __init__ 会自动初始化界面和日志
    main_window = ExcelFileSelector()

    # 3. 显示主窗口
    main_window.show()

    # 4. 启动应用程序的事件循环
    # app.exec() 会阻塞，直到应用程序退出
    # sys.exit() 确保干净地退出
    sys.exit(app.exec())


# 当脚本直接运行时，执行 main 函数
if __name__ == "__main__":
    main()
