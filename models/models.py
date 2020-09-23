"""
models.py will contain all models associated with this app
"""
import os
import sys
import pandas as pd


def verify_target_scraper_workbook():
    try:
        file_target_scrapes = 'scrape_targets.xlsx'
        scrape_df = pd.read_excel(file_target_scrapes, sheet_name='target_data')
        return scrape_df
    except FileNotFoundError:
        #TODO print message
        return None


def verify_sql_connection():
    conn_string = ""


def get_dir_and_workbooks(filepath=None):
    """
    Recursive function
    Asks for user input until valid dir path is provided
    Also ensures the dir path has at least one supported Excel format present
    """

    list_supported_extensions = ['.xlsx', '.xlsm', '.xltx', '.xltm']

    def check_for_xl_files(dirpath):
        """
        Returns False if directory contains no supported extensions
        Else returns tuple, (filepath, supported_files_found)
        """
        valid_xl_files = [f for f in os.listdir(dirpath) if os.path.splitext(f)[1].lower() in list_supported_extensions]
        if len(valid_xl_files) < 1:
            return False
        else:
            return valid_xl_files

    if filepath is None:
        filepath = str(input(r'Paste directory of Excel workbooks: '))
        return get_dir_and_workbooks(filepath=filepath)
    elif filepath is not None:
        if os.path.isdir(filepath) is True:
            supported_files_found = check_for_xl_files(dirpath=filepath)
            if supported_files_found is False:
                print(f"No supported Excel files found, script supports: {list_supported_extensions}")
                return get_dir_and_workbooks(filepath=None)
            else:
                return filepath, supported_files_found
        else:
            # executes if no valid filepath is given in first instance
            print("Filepath not found")
            return get_dir_and_workbooks(filepath=None)
