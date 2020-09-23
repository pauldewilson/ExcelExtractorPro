from server.server import ServerConnection
from models.heavylifter import HeavyLifter
from models.models import (get_dir_and_workbooks,
                           verify_target_scraper_workbook,
                           scrape_df_formatter)

# generate the master scrape_df from scrape_targets.xlsx
# formatting is done later to save because it is unnecessary if other, shorter checks fail first
scrape_df = verify_target_scraper_workbook()

# verify SQL connection and create connection string
server = ServerConnection()
server.postgres_connection()
cs = server.conn_string

# get from user and verify target directory, returns tuple (r'target directory', [target files])
tgt_dir_and_files_tuple = get_dir_and_workbooks()

# make any necessary changes to the scrape_df
# this is not done before now because it is unnecessary if other, shorter checks fail
scrape_df = scrape_df_formatter(scrape_df)

# begin iterating over all target workbooks and extract the data
HeavyLifter(target_dir=tgt_dir_and_files_tuple[0],
            target_files=tgt_dir_and_files_tuple[1],
            scraper_df=scrape_df,
            server_conn_string=cs)

# TODO option to save as CSV or send to SQL

# TODO create UI elements such as printing messages to terminal
# best to create a printing function somewhere to handle formatting

# TODO consider any decorators that could be useful
# TODO consider whether it's worth exceptions throwing sys.exit
