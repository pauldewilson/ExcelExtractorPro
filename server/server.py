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
            # TODO print msg
        except:
            print("Connection failed")
            # TODO print msg

    def mssql_connection(self):
        # TODO implement mssql server connection
        pass
