# src/utils.py
import os
import shutil
from datetime import datetime

def move_files(source_folder, destination_folder, download_start_time):
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)

    for file_name in os.listdir(source_folder):
        file_path = os.path.join(source_folder, file_name)
        if os.path.getmtime(file_path) >= download_start_time.timestamp():
            shutil.move(file_path, os.path.join(destination_folder, file_name))