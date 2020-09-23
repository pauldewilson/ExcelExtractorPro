from server.server import ServerConnection
from models.models import (get_dir_and_workbooks,
                           verify_target_scraper_workbook)

# generate the master scrape_df from scrape_targets.xlsx
scrape_df = verify_target_scraper_workbook()

# verify SQL connection and create connection string
server = ServerConnection()
server.postgres_connection()
cs = server.conn_string

# get from user and verify target directory, returns tuple (r'target directory', [target files])
tgt_dir_and_files_tuple = get_dir_and_workbooks()

#TODO create scrape_targets pandas table from scrape_targets.xlsx to include info the script will need
# things like the table_n since the script will have to run mathematically to allow for scalability
# and also concatenating any range start and end cells into one 'start:end' for easier use in openpyxl

#TODO begin iterating over all target workbooks
# note some messaging re CSV being unsupported, save as xlsx
# need to differentiate between cells/ranges if in the same table
# probably easier to do ranges first since inserting cells to a Series is a one-liner (think: filename)
# insert filename, sheetname, and target cell/range into each table, and scrape timestamp

#TODO option to save as CSV or send to SQL

#TODO create UI elements such as printing messages to terminal
# best to create a printing function somewhere to handle formatting

#TODO consider any decorators that could be useful