import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import os
import time
import tkinter as tk
from tkinter import filedialog, messagebox


class ExcelImageEmbedder:
    def __init__(self, file_path):
        self.file_path = file_path
        self.downloaded_urls = set()

    @staticmethod
    def is_image_url(url):
        """
        判断URL是否为图片地址
        """
        image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp']
        return any(url.lower().endswith(ext) for ext in image_extensions)

    @staticmethod
    def download_image(url, save_path):
        """
        下载图片
        """
        try:
            response = requests.get(url, stream=True)
            response.raise_for_status()
            with open(save_path, 'wb') as file:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        file.write(chunk)
            return True
        except Exception as e:
            print(f"下载图片 {url} 时出错: {e}")
            return False

    def _embed_image_to_cell(self, ws, img_path, row_index, col_index):
        """
        封装嵌入图片到单元格的操作
        """
        try:
            img = Image(img_path)
            img.width = 100
            img.height = 100
            ws.add_image(img, f'{chr(65 + col_index)}{row_index + 1}')
        except Exception as e:
            print(f"在单元格 {chr(65 + col_index)}{row_index + 1} 嵌入图片时出错: {e}")

    def embed_images(self):
        start_time = time.time()
        info = f"开始处理时间: {time.ctime(start_time)}\n"

        # 读取 Excel 文件
        df = pd.read_excel(self.file_path)
        wb = load_workbook(self.file_path)
        ws = wb.active

        # 遍历 DataFrame 的每一行
        for row_index, row in df.iterrows():
            for col_index, value in enumerate(row):
                if isinstance(value, str) and self.is_image_url(value):
                    image_name = os.path.basename(value)
                    save_path = f"downloaded_images/{image_name}"
                    if value in self.downloaded_urls:
                        info += f"图片 {value} 已下载，直接嵌入单元格。\n"
                        self._embed_image_to_cell(ws, save_path, row_index, col_index)
                    else:
                        # 下载图片
                        os.makedirs(os.path.dirname(save_path), exist_ok=True)
                        if self.download_image(value, save_path):
                            self.downloaded_urls.add(value)
                            self._embed_image_to_cell(ws, save_path, row_index, col_index)

        # 保存修改后的 Excel 文件
        new_file_path = self.file_path.replace('.xlsx', '_with_images.xlsx')
        wb.save(new_file_path)

        end_time = time.time()
        info += f"结束处理时间: {time.ctime(end_time)}\n"
        info += f"处理耗时: {end_time - start_time:.2f} 秒\n"
        info += f"已保存修改后的文件到 {new_file_path}"

        return info


def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        embedder = ExcelImageEmbedder(file_path)
        result_info = embedder.embed_images()
        messagebox.showinfo("处理结果", result_info)


if __name__ == "__main__":
    select_file()