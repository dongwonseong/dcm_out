from dotenv import load_dotenv
from util.dbConnect import DatabaseConnector
from excel.parser import ExcelProcess
from util.logMaster import Logger


load_dotenv()

def main(type_value):

    log = Logger('EXCEL')
    log.line()
    log.info(f"{type_value} Excel Parsing Start")

    if type_value == 'ABS':
        excel = ExcelProcess(type_value)
        excel.excel_process()

    elif type_value == 'FB':
        excel = ExcelProcess(type_value)
        excel.excel_process()

    elif type_value == 'SB':
        excel = ExcelProcess(type_value)
        excel.excel_process()

    else:
        return None

    log.info(f" {type_value} Excel Parsing Success !")


if __name__ == "__main__":
    db = DatabaseConnector()
    db.connect()

    query = ("SELECT TYPE FROM DCM_DAILY_LIST GROUP BY type")
    query_results = db.execute_query(query)
    db.close()

    for result in query_results:
        type_value = result['TYPE']
        main(type_value)