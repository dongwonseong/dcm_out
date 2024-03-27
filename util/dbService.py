import os
import re
from util.dbConnect import DatabaseConnector
from dotenv import load_dotenv
from datetime import datetime, timedelta
from util.logMaster import Logger

class DatabaseService:
    def __init__(self):
        self.db = DatabaseConnector()
        self.log = Logger('MYSQL')

    def execute_query_for_type(self, type_value):
        self.db.connect()
        today = datetime.now().strftime('%Y%m%d')
        query = (f"SELECT * "
                 f"FROM DCM_DAILY "
                 f"WHERE bond_type = '{type_value}' and issue_date = '{today}' "
                 f"ORDER BY issue_date asc")

        query_results = self.db.execute_query(query)
        self.log.info(f"query = "
                      f"{query}")
        for result in query_results:
            self.log.info(f"Find Bond [{type_value}] [{result['bond_name']}]")

        self.db.close()

        return query_results
