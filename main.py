import sys
import platform
import logging
import argparse
from PyQt6.QtWidgets import QApplication
from PyQt6.QtCore import Qt

import custom_log_config
from excel_file_selector import ExcelFileSelector

# 应用程序元数据
APP_NAME = "Excel Image Embedder"
APP_VERSION = "1.0.0"


def setup_basic_logging(verbose: bool = False) -> None:
    """在 GUI 初始化前配置基本的控制台日志记录。"""
    level = custom_log_config.LOG_LEVEL
    logging.basicConfig(
        level=level,
        format=custom_log_config.LOG_FORMAT,
        handlers=[logging.StreamHandler(sys.stderr)]
    )
    logging.debug("基本控制台日志记录已初始化。")


def parse_arguments() -> argparse.Namespace:
    """解析命令行参数。"""
    try:
        parser = argparse.ArgumentParser(description="Excel Image Embedder 应用程序")
        parser.add_argument(
            '--verbose',
            action='store_true',
            help='启用调试日志'
        )
        return parser.parse_args()
    except Exception as e:
        logging.error(f"解析命令行参数失败: {e}")
        raise


def main() -> int:
    """
    Excel Image Embedder 应用程序的主入口点。

    初始化 QApplication，设置基本日志记录，创建并显示主窗口，并启动事件循环。

    返回:
        int: 应用程序的退出代码（成功为 0，失败为非零）。
    """
    # 解析命令行参数
    args = parse_arguments()

    # 设置基本控制台日志记录以捕获早期错误
    setup_basic_logging(verbose=args.verbose)

    try:
        # 使用系统参数初始化 QApplication
        app = QApplication(sys.argv)
        logging.debug("QApplication 已初始化。")

        # 配置 QApplication 属性
        app.setApplicationName(APP_NAME)
        app.setApplicationVersion(APP_VERSION)
        # 仅在 macOS 上禁用原生对话框以缓解 IMK 错误
        if platform.system() == "Darwin":  # macOS
            app.setAttribute(Qt.ApplicationAttribute.AA_DontUseNativeDialogs)
        app.setQuitOnLastWindowClosed(True)

        # 创建并显示主窗口
        main_window = ExcelFileSelector()
        main_window.show()
        logging.info(f"{APP_NAME} v{APP_VERSION} 已启动。")

        # 启动事件循环并返回退出代码
        return app.exec()

    except ImportError as e:
        logging.error(f"无法导入所需模块: {e}")
        return 1
    except RuntimeError as e:
        logging.error(f"初始化 QApplication 失败: {e}")
        return 2
    except Exception as e:
        logging.error(f"初始化期间发生意外错误: {e}")
        return 3
    finally:
        # 确保干净关闭
        logging.debug("正在关闭应用程序。")
        if 'app' in locals():
            app.quit()


if __name__ == "__main__":
    sys.exit(main())
