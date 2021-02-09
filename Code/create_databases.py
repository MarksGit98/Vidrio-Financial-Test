import pandas as pd
from openpyxl import load_workbook
from sqlalchemy import create_engine

class ReadandWrite:
    def __init__(self):
        self.engine = create_engine('mysql+pymysql://root:root@localhost:3306/vidrio_test')
        self.con = self.engine.connect()

    def excel_to_db(self, filenames):
        for filename in filenames:
            wb = load_workbook(filename)
            xls = pd.ExcelFile(filename)
            dataframes = []
            for worksheet in wb.sheetnames:
                dataframes.append((pd.read_excel(xls, worksheet), worksheet))

            for df, name in dataframes:
                df.columns = df.columns.str.strip()
                df.to_sql(name, con=self.con)
                # print(df)

    def generate_portfolio_valuation_file1(self, filename):
        data = {
            "Reference Day": [],
            "Periodicity": [],
            "Investor Account UID": [],
            "Investment Account UID": [],
            "Investment Account Long Name": [],
            "Attribution Gross": [],
            "Attribition Net": [],
            "Opening Allocation": [],
            "Closing Allocation": [],
            "Opening Equity": [],
            "Closing Allocation": [],
            "Opening Equity": [],
            "Closing Equity": [],
            "Investment Performance": [],
            "Investment Adj Opening Balance": [],
            "Investment Closing Balance": [],
            "Portfolio Opening Balance": [],
            "Portfolio Closing Balance": [],
        }

    def generate_portfolio_valuation_file2(self, filename):
        data = {
            "Portfolio Account UID": [],
            "Account Long Name": [],
            "Date": [],
            "NAV/Share": [],
            "Final": []
        }

    def run_query(self, query):
        queryset = self.con.execute(f"SELECT * FROM constituents")
        print(queryset.fetchall()[0])