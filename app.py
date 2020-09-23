# verify scrape_targets.xlsx

# verify target directory

# create scrape_targets pandas table from scrape_targets.xlsx to include info the script will need
# things like the table_n since the script will have to run mathematically to allow for scalability
# and also concatenating any range start and end cells into one 'start:end' for easier use in openpyxl

# begin iterating over all target workbooks
# note some messaging re CSV being unsupported, save as xlsx
# need to differentiate between cells/ranges if in the same table
# probably easier to do ranges first since inserting cells to a Series is a one-liner (think: filename)
# insert filename, sheetname, and target cell/range into each table, and scrape timestamp

# option to save as CSV or send to SQL
