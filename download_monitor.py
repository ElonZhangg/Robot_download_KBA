# -*- coding: utf-8 -*-
"""
下载文件监控工具 - 等待新文件下载完成
"""
import os
import time
import glob

class DownloadMonitor:
    def __init__(self, download_dir):
        self.download_dir = download_dir
        if not os.path.exists(download_dir):
            os.makedirs(download_dir)

    def wait_for_download(self, timeout=30, poll_interval=1):
        """
        等待下载目录中出现新文件并返回文件路径
        原理：记录下载前目录中的文件集合，下载后等待新文件出现且大小稳定
        """
        # 获取下载前文件列表
        before_files = set(os.listdir(self.download_dir))
        start_time = time.time()
        while time.time() - start_time < timeout:
            after_files = set(os.listdir(self.download_dir))
            new_files = after_files - before_files
            if new_files:
                # 取最新的文件（可能有多个，取第一个）
                new_file = max([os.path.join(self.download_dir, f) for f in new_files], 
                               key=os.path.getctime)
                # 等待文件大小稳定（连续两次检查大小不变）
                stable = False
                last_size = -1
                for _ in range(5):
                    if os.path.exists(new_file):
                        current_size = os.path.getsize(new_file)
                        if current_size == last_size and current_size > 0:
                            stable = True
                            break
                        last_size = current_size
                    time.sleep(1)
                if stable:
                    return new_file
            time.sleep(poll_interval)
        raise TimeoutError(f"下载超时（{timeout}秒），未检测到新文件")

    def clear_old_files(self):
        """清空下载目录（可选，避免重复文件干扰）"""
        for f in glob.glob(os.path.join(self.download_dir, "*")):
            try:
                os.remove(f)
            except:
                pass