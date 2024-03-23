import os
import re
from util.dbConnect import DatabaseConnector
from dotenv import load_dotenv
from datetime import datetime

class DatabaseService:
    def __init__(self):
        self.db = DatabaseConnector()

    def execute_query_for_type(self, type_value):
        self.db.connect()

        query = (f"SELECT * FROM DCM_DAILY WHERE bond_type = '{type_value}' ORDER BY issue_date asc")
        query_results = self.db.execute_query(query)

        self.db.close()

        return query_results
