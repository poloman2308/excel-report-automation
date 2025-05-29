import pandas as pd
import os
from datetime import datetime
from .utils import auto_size_columns


class BaseReport:
    def __init__(self, input_file, output_dir, logo_path=None):
        self.input_file = input_file
        self.output_dir = output_dir
        self.logo_path = logo_path
        self.output_file = self.generate_output_filename()
        self.df = None
        self.issues = {}
        
    def generate_output_filename(self):
        filename = f"Sales_Report_{datetime.now().strftime('%Y_%m_%d')}.xlsx"
        return os.path.join(self.output_dir, filename)

    def load_data(self):
        self.df = pd.read_csv(self.input_file, parse_dates=['OrderDate'])
        self.df['Revenue'] = self.df['Quantity'] * self.df['UnitPrice']
        self.df['Month'] = self.df['OrderDate'].dt.to_period('M')
        
        if self.df.isnull().any().any():
            self.issues["missing] = self.df[self.df.isnull().any(axis=1)]"]
            
        dupes = self.df.duplicated()
        if dupes.any():
            self.issues["duplicates"] = self.df[dupes]
            
        z_scores = (self.df['Revenue'] - self.df['Revenue'].mean()) / self.df['Revenue'].std()
        outliers = z_scores.abs() > 3
        if outliers.any():
            self.issues["outliers"] = self.df[outliers]

    def generate(self):
        self.load_data()
        
        with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as self.writer:
            self.workbook = self.writer.book
            self.write_raw_data()
            self.write_summary()
            self.write_pivot()
            self.insert_chart()
            self.write_issues()

    def write_raw_data(self):
        sheet_name = "RawData"
        self.df.to_excel(self.writer, sheet_name=sheet_name, index=False)
        worksheet = self.writer.sheets[sheet_name]
        header_format = self.workbook.add_format({"bold": True, "bg_color": "#D9EAD3", "border": 1})

        for col_num, value in enumerate(self.df.columns.values):
            worksheet.write(0, col_num, value, header_format)

        self.auto_size_columns(worksheet, self.df)
        
    def write_issues(self):
        for issue_type, issue_df in self.issues.items():
            sheet_name = f"Issues_{issue_type}"
            issue_df.to_excel(self.writer, sheet_name=sheet_name, index=False)
            worksheet = self.writer.sheets[sheet_name]
            header_format = self.workbook.add_format({"bold": True, "bg_color": "#FCE5CD", "border": 1})

            for col_num, value in enumerate(issue_df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            self.auto_size_columns(worksheet, issue_df)
            
            # Apply conditional formatting for issues
            if issue_type == "missing":
                worksheet.conditional_format(
                    f"A2:{chr(65 + len(issue_df.columns) - 1)}{len(issue_df) + 1}",
                    {
                        "type": "blanks",
                        "format": self.workbook.add_format({"bg_color": "#FFC7CE"})
                    }
                )

            elif issue_type == "duplicates":
                worksheet.conditional_format(
                    f"A2:{chr(65 + len(issue_df.columns) - 1)}{len(issue_df) + 1}",
                    {
                        "type": "no_errors",  # Entire row will be highlighted
                        "format": self.workbook.add_format({"bg_color": "#F9CB9C"})
                    }
                )

            elif issue_type == "outliers" and "Revenue" in issue_df.columns:
                col_index = issue_df.columns.get_loc("Revenue")
                col_letter = chr(65 + col_index)
                worksheet.conditional_format(
                    f"{col_letter}2:{col_letter}{len(issue_df) + 1}",
                    {
                        "type": "cell",
                        "criteria": ">",
                        "value": self.df["Revenue"].mean() + 3 * self.df["Revenue"].std(),
                        "format": self.workbook.add_format({"bg_color": "#FFEB9C"})
                    }
                )
            
    def auto_size_columns(self, worksheet, dataframe):
        for idx, col in enumerate(dataframe.columns):
            max_len = max(
                dataframe[col].astype(str).map(len).max(),
                len(str(col)) + 2
            )
            worksheet.set_column(idx, idx, max_len)

    def write_summary(self):
        raise NotImplementedError("Must implement write_summary in subclass")

    def write_pivot(self):
        raise NotImplementedError("Must implemented write_pivot in subclass")

    def insert_chart(self):
        raise NotImplementedError("Must implemented insert_chart in subclass")
