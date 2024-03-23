import pymysql
from dotenv import load_dotenv
import os

# 환경 변수 로드
load_dotenv()

class DatabaseConnector:
    def __init__(self):
        self.host = os.getenv('DB_HOST')
        self.port = int(os.getenv('DB_PORT'))
        self.user = os.getenv('DB_USER')
        self.password = os.getenv('DB_PASS')
        self.database = os.getenv('DB_NAME')
        self.connection = None

    def connect(self):
        if not self.connection or not self.connection.open:
            try:
                self.connection = pymysql.connect(
                    host=self.host,
                    port=self.port,
                    user=self.user,
                    password=self.password,
                    database=self.database,
                    cursorclass=pymysql.cursors.DictCursor
                )

            except pymysql.MySQLError as e:
                print(f"Error connecting to MySQL database: {e}")


    def execute_query(self, query, params=None, query_type='read'):
        self.connect()
        with self.connection.cursor() as cursor:
            cursor.execute(query, params)
            if query_type == 'write':
                self.connection.commit()
                return cursor.rowcount
            else:
                return cursor.fetchall()


    def close(self):
        if self.connection and self.connection.open:
            self.connection.close()

    def loadQuery(self, query_id, date, fileName):
        queryPath = os.getenv('QueryPath')
        with open(queryPath, 'r', encoding='utf-8') as file:
            queries = file.read().split("-- END --")
        for q in queries:
            if query_id in q:
                queryFormat = q.strip().split('\n', 1)[1].strip()
                query = queryFormat.format(date=date)
                queryResult = self.execute(query)

                if queryResult:
                    return self.toCsv(queryResult, fileName)
                else:
                    return False

        return False



