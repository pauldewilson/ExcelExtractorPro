from sqlalchemy import create_engine


class ServerConnection:
    """
    Create instance of class and call relevant connection to generate the necessary conn_string
    Takes server variables from user to ensure correction connection variables
    Each method attempts server connection before returning connection string
    """

    def __init__(self):
        self.conn_string = ""

    def postgres_connection(self, db_name='excel_extractor', user='postgres', password='postgres', host='localhost'):
        self.conn_string = f"postgresql://{user}:{password}@{host}/{db_name}"
        try:
            engine = create_engine(self.conn_string)
            engine.connect()
        except:
            print("Postgres Connection failed")

    def mssql_connection(self, db_name, host, driver='SQL+Server+Native+Client+11.0', trusted_connection='yes'):
        """
        Trusted connection is windows auth.
        If this is to change to SQL login then it must be changed
        """
        self.conn_string = f"mssql+pyodbc://@{host}/{db_name}?driver={driver}?trusted_connection={trusted_connection}"
        try:
            engine = create_engine(self.conn_string)
            engine.connect()
        except:
            print("MSSQL Connection failed")
