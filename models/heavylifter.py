from openpyxl import load_workbook
import time
import os


class HeavyLifter:
    """
    This is the main bulk of script.
    It takes in the list of files to be extracted and does the extracting
    """

    def __init__(self, target_dir, target_files, scraper_df, server_conn_string):
        # __init__ variables
        self.target_dir = target_dir
        self.target_files = target_files
        self.scraper_df = scraper_df
        self.server_conn_string = server_conn_string
        # executing methods
        self.run()

    def run(self):
        # change to dir with workbooks
        os.chdir(self.target_dir)

        # TODO: implement try/excepts

        # iterate over all workbooks
        for file in self.target_files:
            # load workbook
            wb = load_workbook(file)
            print(f"1 FILE: {file}")

            # iterate over each worksheet in scraper_df
            for worksheet in self.scraper_df['source_sheet'].unique():
                # load worksheet
                ws = wb[worksheet]
                print(f"2 WORKSHEET: {worksheet}")

                # working hierarchically, iterate over all tables, columns, cells and create dataframe
                for sql_table in self.scraper_df['sql_table'].unique():
                    print("____________________________")
                    print(sql_table)
                    print("____________________________")
                    # data_dict will be the dataframe for each table that will be sent to the server
                    data_dict = dict()

                    print(f"3 SQL TABLE: {sql_table}")
                    # filter df so only looking at this particular SQL table
                    df_filt_sql = self.scraper_df[self.scraper_df['sql_table'] == sql_table]

                    for column_name in df_filt_sql['column_name']:
                        print(f"4 COLUMN NAME. {column_name}")

                        # filtering for the row
                        df_filt_col = df_filt_sql[df_filt_sql['column_name'] == column_name]
                        range_or_cell = df_filt_col['is_target_a_cell_or_range_of_cells'].values[0]
                        row_or_column = df_filt_col['range_type'].values[0]

                        print(f"5. RANGE OR CELL, ROW OR COLUMN: {range_or_cell}, {row_or_column}")

                        if range_or_cell == 'cell':
                            print("DO CELL EXTRACTION")
                        elif range_or_cell == 'range':
                            if row_or_column == 'row':
                                print("--ROW EXTRACTION")
                            elif row_or_column == 'column':
                                print("--COLUMN EXTRACTION")

                # break out of SQL table and move to next
                break
