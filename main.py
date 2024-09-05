from dotenv import load_dotenv
from excel.parser import ExcelProcess
from util.logMaster import Logger
import argparse

load_dotenv()


def main(type_value, date):

    log = Logger('EXCEL')
    log.line()
    log.info(f"{type_value} Excel Parsing Start")

    excel = ExcelProcess(type_value, date)
    excel.excel_process()

    log.info(f" {type_value} Excel Parsing Success !")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Process some integers.")
    parser.add_argument('-d', '--date', help="Date parameter")
    parser.add_argument('-t', '--type', help="type")
    args = parser.parse_args()

    if args.type == 'daily':
        types = ['ABS', 'SB', 'FB']
        for type in types:
            main(type, args.date)
    else:
        main(args.type, args.date)
