# IMPORT

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference

WORKSHEET = "PythonSheets.xlsx"
SHEETS = ["Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5"]
MSHEET = "MasterSheet"


# init method is called and reading the excel file

class Aggregator:
    def __init__(self, worksheet, sheets):
        self.worksheet, self.sheets = worksheet, sheets
        self.dfs = pd.read_excel(worksheet, sheet_name=sheets, parse_dates=False)

    #  Validation (PS Number) from all sheets

    def get_input(self, c=0):
        query = input("Enter PS Number / Email / Name: ")        # query is what we've provided in input
        if query and c < 3:  # check if user has entered something or not
            try:
                query = int(query)              # change ps no as integer
                searchid = "Ps No"              # ps no is saved as searchid
            except ValueError:
                if "@" in query:
                    searchid = "Email"
                else:
                    searchid = "Name"
            return query, searchid
        elif c == 3:
            print("Too many wrong attempts, try again later.")
            exit()
        else:
            print("No input found, try again!")
            return self.get_input(c + 1)

    # values of dataframe gets stored in x and we're updating the fields.

    def search(self, query, searchid):
        print(f"Searching {searchid.lower()} `{query}`...")
        fields = {}
        for x in self.dfs.values():
            fields.update(
                x[x[searchid] == query].to_dict(orient="list"))  # checking if the query exist in that col or not
        # print(fields)

        if fields[searchid]:
            print("Found.")
            return pd.DataFrame.from_dict(fields)  # changing the dict into dataframe

        print(f"Couldn't find the {searchid.lower()} in sheets!")
        exit()

    def add_to_master(self, df):

        # df["Entry Time"] = df["Entry Time"].dt.strftime("%I:%S %p")
        # df["Exit Time"] = df["Exit Time"].dt.strftime("%H:%S %p")
        df["Start Date"] = df["Start Date"].dt.strftime("%d/%m/%Y")
        df["End Date"] = df["End Date"].dt.strftime("%d/%m/%Y")

        book = load_workbook(self.worksheet)  # loading the workbook
        with pd.ExcelWriter(self.worksheet, engine="openpyxl", mode="a") as writer:  # opening a Excel Writer instance
            writer.book = book  # changing the workbook of the writer to our current workbook
            writer.sheets = {ws.title: ws for ws in book.worksheets}  # adding the worksheets to it
            # getting the last row if not found set it to 0
            try:
                startrow = writer.sheets[MSHEET].max_row
            except KeyError:
                startrow = 0
            # checking if we have already have a master sheet
            if MSHEET in writer.book.sheetnames:
                df.to_excel(
                    writer,
                    index=False,
                    header=False,
                    sheet_name=MSHEET,
                    startrow=startrow,
                )
            else:
                df.to_excel(writer, index=False, sheet_name=MSHEET, startrow=startrow)
        print(f"Added to {MSHEET}.")

        # for bar graph

    def BarChart(self):

        book = load_workbook(self.worksheet)
        chart1 = BarChart()
        chart1.type = "col"  # taken output of any two cols
        chart1.style = 10
        chart1.title = "Bar Chart"
        chart1.y_axis.title = 'Test number'
        chart1.x_axis.title = 'Sample length (mm)'

        data = Reference(book[MSHEET], min_col=1, min_row=4,
                         max_col=2, max_row=book[MSHEET].max_row)
        cats = Reference(book[MSHEET], min_col=1, min_row=2, max_row=7)
        chart1.add_data(data, titles_from_data=True)
        chart1.set_categories(cats)
        chart1.shape = 4
        book[MSHEET].add_chart(chart1, "A10")
        book.save(WORKSHEET)


no_of_inputs = int(input("Select the number of inputs: "))
for i in range(no_of_inputs):

    # main function

    if __name__ == "__main__":
        agg = Aggregator(WORKSHEET, SHEETS)

        query, searchid = agg.get_input()

        df_m = agg.search(query, searchid)

        agg.add_to_master(df_m)

        agg.BarChart()
