from dotenv import load_dotenv
from util.dbConnect import DatabaseConnector
from excel.parser import ExcelProcess

load_dotenv()

def main(type_value):
    print(type_value)

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

if __name__ == "__main__":
    db = DatabaseConnector()
    db.connect()

    query = ("SELECT TYPE FROM DCM_DAILY_LIST GROUP BY type")
    query_results = db.execute_query(query)
    db.close()

    # test = 'SB'
    # main(test)


    for result in query_results:
        type_value = result['TYPE']
        print(type_value)
        main(type_value)