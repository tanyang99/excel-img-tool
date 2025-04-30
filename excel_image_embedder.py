import concurrent.futures  # Import ThreadPoolExecutor
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import requests
import time
import logging
import queue

# Maximum file count
MAX_FILE_COUNT = 10
MAX_TOTAL_SIZE = 500 * 1024 * 1024
MAX_CONCURRENT_DOWNLOADS = 3


class ExcelImageEmbedder:
    def __init__(self):
        self._successfully_downloaded_urls = set()
        self._download_queue = queue.Queue()  # Queue for URLs to download

    @staticmethod
    def is_image_url(value):
        """
        判断单元格值是否为可能是图片地址的字符串
        :param value: 单元格的值
        :return: 如果是字符串且符合图片URL格式返回True，否则返回False
        """
        if not isinstance(value, str):
            return False
        # Simple check for common image extensions in the URL path (case-insensitive)
        # This is not foolproof but covers most cases.
        image_extensions = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp')  # Add more common formats
        # Check both the value itself and potentially if it's part of a larger string
        lower_value = value.lower()
        return lower_value.startswith(('http://', 'https://')) and any(ext in lower_value for ext in image_extensions)

    def _download_image(self, url, save_path):
        """
        下载图片并保存到指定路径
        :param url: 图片URL
        :param save_path: 保存路径
        :return: 保存路径如果下载成功，否则返回 None
        """
        try:
            # Avoid re-downloading within the same instance if already successful
            if url in self._successfully_downloaded_urls:
                logging.debug(f"图片 {url} 已在本会话中成功下载，跳过下载。")
                return save_path  # Return expected path as if downloaded again

            response = requests.get(url, stream=True, timeout=10)  # Add a timeout
            response.raise_for_status()  # Raise HTTPError for bad responses (4xx or 5xx)

            # Ensure directory exists
            os.makedirs(os.path.dirname(save_path), exist_ok=True)

            with open(save_path, 'wb') as file:
                for chunk in response.iter_content(chunk_size=8192):
                    file.write(chunk)

            self._successfully_downloaded_urls.add(url)  # Mark as successfully downloaded
            logging.debug(f"图片 {url} 下载成功，保存到 {save_path}")
            return save_path

        except requests.exceptions.Timeout:
            logging.error(f"下载图片 {url} 时发生超时错误。")
            return None
        except requests.exceptions.HTTPError as http_err:
            logging.error(f"下载图片 {url} 时发生HTTP错误: {http_err}")
            return None
        except requests.exceptions.RequestException as req_err:
            logging.error(f"下载图片 {url} 时发生请求错误: {req_err}")
            return None
        except Exception as e:
            logging.error(f"下载图片 {url} 时发生未知错误: {e}")
            return None

    def _embed_image_to_cell(self, ws, img_path, row_index, col_index):
        """
        封装嵌入图片到单元格的操作
        :param ws: 工作表对象 (openpyxl worksheet)
        :param img_path: 图片的路径
        :param row_index: 单元格的行索引 (0-based)
        :param col_index: 单元格的列索引 (0-based)
        :return: 嵌入成功返回True，失败返回False
        """
        # openpyxl uses 1-based indexing for cell coordinates
        cell_coordinate = f'{chr(65 + col_index)}{row_index + 1}'
        try:
            # Check if the image file actually exists before attempting to embed
            if not os.path.exists(img_path):
                logging.error(f"在单元格 {cell_coordinate} 嵌入图片时出错: 图片文件不存在于 {img_path}")
                return False

            img = Image(img_path)
            # Set a default size, you might want to make this configurable
            img.width = 100
            img.height = 100

            # Add the image to the worksheet, anchored to the top-left corner of the cell
            ws.add_image(img, cell_coordinate)
            # logging.info(f"图片成功嵌入到单元格 {cell_coordinate}") # Removed as per request 1
            return True
        except Exception as e:
            # Log error message only for failed embeds as per request 1
            logging.error(f"在单元格 {cell_coordinate} 嵌入图片时出错: {e}")
            return False

    @staticmethod
    def check_file_count_and_size(file_paths):
        """
        检查文件数量和总大小
        :param file_paths: 文件路径列表
        :return: 如果文件数量和总大小符合要求返回True，否则返回False
        """
        # Check file count
        if len(file_paths) > MAX_FILE_COUNT:
            logging.error(f"错误: 最多只能选择 {MAX_FILE_COUNT} 个文件。")
            return False
        # Calculate total size
        total_size = 0
        for path in file_paths:
            if not os.path.exists(path):
                logging.error(f"错误: 文件 {path} 未找到或无法访问。")
                return False
            try:
                total_size += os.path.getsize(path)
            except Exception as e:
                logging.error(f"获取文件 {path} 大小失败：{e}")
                return False
        # Check total size
        if total_size > MAX_TOTAL_SIZE:
            logging.error(f"错误: 选择的文件总大小不能超过 {MAX_TOTAL_SIZE / (1024 * 1024):.2f} MB。")
            return False
        return True

    @staticmethod
    def get_file_and_sheet_info(file_paths):
        """
        获取文件和工作表信息，同时返回sheet_index (0-based)
        Uses pandas just for convenience of reading sheet names.
        :param file_paths: 文件路径列表
        :return: 包含文件basename和其sheet信息的字典 {file_name: [(sheet_index, sheet_name), ...]}
        """
        file_sheet_info = {}
        for file_path in file_paths:
            if not os.path.exists(file_path):
                logging.warning(f"文件 {file_path} 未找到，跳过处理。")
                continue
            try:
                # Using pandas just to get sheet names and indices easily
                excel_file = pd.ExcelFile(file_path)
                sheet_names = excel_file.sheet_names
                file_name = os.path.basename(file_path)
                # Store sheet name and corresponding 0-based index
                sheet_info = [(index, sheet_name) for index, sheet_name in enumerate(sheet_names)]
                file_sheet_info[file_name] = sheet_info
                logging.debug(f"读取文件 {file_name} 的 sheet 信息成功。")
            except Exception as e:
                logging.error(f"读取文件 {file_path} 的 sheet 信息出错: {e}")
        return file_sheet_info

    def embed_images(self, file_paths, sheets_to_process_map):
        """
        嵌入图片到Excel文件中
        :param file_paths: 原始文件路径列表
        :param sheets_to_process_map: 包含需要处理的工作表索引的字典
               格式: {file_basename: [sheet_index, ...], ...}
        :return: None. Logs success/failure.
        """
        start_time = time.time()
        logging.info("\n \n")
        logging.info("-------------- 开始图片嵌入处理 --------------")
        logging.info(f"待处理文件: {[os.path.basename(p) for p in file_paths]}")
        logging.info(f"待处理 sheets (索引): {sheets_to_process_map}")

        output_dir = "excel_with_images"  # Define a dedicated output directory
        os.makedirs(output_dir, exist_ok=True)

        total_files_processed = 0
        total_successful_files = 0

        for file_path in file_paths:
            file_basename = os.path.basename(file_path)
            sheets_to_process = sheets_to_process_map.get(file_basename, [])

            if not sheets_to_process:
                logging.info(f"文件 {file_basename} 没有选定的 sheet 进行处理，跳过。")
                continue

            total_files_processed += 1
            logging.info(f"-> 开始处理文件: {file_basename}")
            self._successfully_downloaded_urls.clear()  # Clear downloaded URLs per file

            try:
                # Use read-only=False to allow modifications
                wb = load_workbook(file_path)  # Read workbook once
                sheet_names = wb.sheetnames

                urls_to_download = set()
                url_save_path_map = {}  # Map URL to its intended save path

                # --- Step 1: Collect all unique image URLs from selected sheets ---
                logging.info(f"--- 收集文件 {file_basename} 中选定 sheets 的图片链接 ---")
                for sheet_index in sheets_to_process:
                    if not (0 <= sheet_index < len(sheet_names)):
                        logging.warning(f"文件 {file_basename}: 指定的 Sheet 索引 {sheet_index} (0-based) 不存在，跳过。")
                        continue
                    sheet_name = sheet_names[sheet_index]
                    logging.debug(f"正在收集 Sheet: {sheet_name} (Index: {sheet_index}) 的链接...")
                    ws = wb[sheet_name]
                    for row_index, row in enumerate(ws.iter_rows()):
                        for col_index, cell in enumerate(row):
                            if self.is_image_url(cell.value):
                                url = cell.value.strip()  # Trim whitespace
                                if url not in urls_to_download:
                                    urls_to_download.add(url)
                                    # Define a save path based on URL or a hash
                                    # Using a simple hash to avoid invalid filename characters from URL
                                    import hashlib
                                    url_hash = hashlib.md5(url.encode('utf-8')).hexdigest()
                                    # Try to keep original extension if possible
                                    ext = os.path.splitext(url.lower())[1]
                                    if not ext or not any(image_ext in ext for image_ext in
                                                          ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg', '.webp')):
                                        ext = '.jpg'  # Default extension if none found or recognized
                                    img_filename = f"{url_hash}{ext}"
                                    save_path = os.path.join("downloaded_images", img_filename)
                                    url_save_path_map[url] = save_path  # Map URL to its save path

                if not urls_to_download:
                    logging.info(f"文件 {file_basename} 的选定 sheets 中没有找到图片链接，跳过下载和嵌入。")
                else:
                    # --- Step 2: Download images concurrently ---
                    logging.info(f"--- 开始下载文件 {file_basename} 的图片 ({len(urls_to_download)} 张) ---")
                    download_results = {}  # {url: downloaded_path or None}
                    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_CONCURRENT_DOWNLOADS) as executor:
                        # Submit download tasks
                        future_to_url = {executor.submit(self._download_image, url, url_save_path_map[url]): url for url
                                         in urls_to_download}
                        for future in concurrent.futures.as_completed(future_to_url):
                            url = future_to_url[future]
                            try:
                                save_path = future.result()
                                download_results[url] = save_path  # Store the resulting path (or None if failed)
                            except Exception as exc:
                                logging.error(f'{url} 下载时出现异常: {exc}')
                                download_results[url] = None  # Mark as failed download

                    successful_downloads = sum(1 for path in download_results.values() if path is not None)
                    failed_downloads = len(urls_to_download) - successful_downloads
                    logging.info(
                        f"--- 文件 {file_basename} 图片下载完成：成功 {successful_downloads} 张，失败 {failed_downloads} 张 ---")

                    # --- Step 3: Embed downloaded images into selected sheets ---
                    logging.info(f"--- 开始嵌入文件 {file_basename} 中选定 sheets 的图片 ---")
                    successful_embeds = 0
                    failed_embeds = 0
                    total_attempted_embeds = 0  # Count how many cells *contained* a URL

                    for sheet_index in sheets_to_process:
                        if not (0 <= sheet_index < len(sheet_names)):
                            continue  # Already logged warning above
                        sheet_name = sheet_names[sheet_index]
                        logging.debug(f"正在嵌入 Sheet: {sheet_name} (Index: {sheet_index}) 的图片...")
                        ws = wb[sheet_name]
                        for row_index, row in enumerate(ws.iter_rows()):
                            for col_index, cell in enumerate(row):
                                if self.is_image_url(cell.value):
                                    url = cell.value.strip()
                                    total_attempted_embeds += 1
                                    downloaded_path = download_results.get(url)  # Get path from results

                                    if downloaded_path:  # If download was successful
                                        if self._embed_image_to_cell(ws, downloaded_path, row_index, col_index):
                                            successful_embeds += 1
                                        else:
                                            failed_embeds += 1  # Embedding failed even with file
                                    else:
                                        # Log failure if download itself failed for this URL
                                        logging.error(
                                            f"在单元格 {chr(65 + col_index)}{row_index + 1} 嵌入图片时出错: 图片 {url} 下载失败。")
                                        failed_embeds += 1  # Count as embedding failure due to download failure

                    logging.info(f"--- 文件 {file_basename} 图片嵌入完成：成功 {successful_embeds} 张，失败 {failed_embeds} 张，"
                                 f"共检查 {total_attempted_embeds} 个包含图片链接的单元格 ---")

                    # --- Step 4: Save the modified Excel file ---
                    if successful_embeds > 0 or failed_embeds > 0:  # Save if any embeds were attempted
                        new_file_name = f"{os.path.splitext(file_basename)[0]}_with_images.xlsx"
                        new_file_path = os.path.join(output_dir, new_file_name)
                        try:
                            wb.save(new_file_path)
                            logging.info(f"-> 已保存修改后的文件到 {new_file_path}")
                            total_successful_files += 1
                        except Exception as e:
                            logging.error(f"保存修改后的文件 {new_file_path} 时出错: {e}")
                    else:
                        logging.info(f"文件 {file_basename} 中选定的 sheets 没有图片被处理或尝试嵌入，不生成新文件。")

            except FileNotFoundError:
                logging.error(f"错误: 处理文件时 {file_path} 未找到。")
            except Exception as e:
                logging.error(f"处理文件 {file_path} 时发生意外错误: {e}")

        end_time = time.time()
        logging.info("-------------- 图片嵌入处理结束 --------------")
        logging.info(f"总共处理了 {total_files_processed} 个文件，成功生成 {total_successful_files} 个带图片的输出文件。")
        logging.info(f"总处理耗时: {end_time - start_time:.2f} 秒")
