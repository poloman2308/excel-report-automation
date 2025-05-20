import pandas as pd
import os
from datetime import datetime
from .utils import auto_size_columns


class BaseReport:
    def __init__(self, input_file, output_dir, logo_path=None):
        self.input_file = input_file
        self.output_dir = output_dir
        self.logo_path = logo_path
        self.today_str = datetime.today().strftime('%Y_%m_%d')
        self.output_file = os.path.join(output_dir, f"Sales_Report_{self.today_str}.xlsx")

        self.df = None
        self.writer = None
        self.workbook = None

    def load_data(self):
        self.df = pd.read_csv(self.input_file, parse_dates=['OrderDate'])
        self.df['Revenue'] = self.df['Quantity'] * self.df['UnitPrice']
        self.df['Month'] = self.df['OrderDate'].dt.to_period('M')

    def generate(self):
        self.load_data()
        os.makedirs(self.output_dir, exist_ok=True)
        with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as self.writer:
            self.workbook = self.writer.book
            self.write_raw_data()
            self.write_summary()
            self.write_pivot()
            self.insert_chart()
        print(f"✅ Report saved: {self.output_file}")

    def write_raw_data(self):
        self.df.to_excel(self.writer, sheet_name='RawData', index=False)
        
    def write_raw_data(self):
        self.df.to_excel(self.writer, sheet_name='RawData', index=False)
        worksheet = self.writer.sheets['RawData']
        auto_size_columns(worksheet, self.df)

    def write_summary(self):
        raise NotImplementedError("Must be implemented by subclass")

    def write_pivot(self):
        raise NotImplementedError("Must be implemented by subclass")

    def insert_chart(self):
        raise NotImplementedError("Must be implemented by subclass")
