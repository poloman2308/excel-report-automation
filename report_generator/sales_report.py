import os
import pandas as pd
from .base_report import BaseReport
from report_generator.utils import auto_size_columns, currency_format, top_5_highlight_format

class SalesReport(BaseReport):
    def generate(self):
        self.load_data()

        with pd.ExcelWriter(self.output_file, engine='xlsxwriter') as self.writer:
            self.workbook = self.writer.book
            self.write_raw_data()
            self.write_summary()
            self.write_pivot()
            self.insert_chart()
            
            if self.issues:
                self.write_issues()
            self.export_for_powerbi()
    
    def write_summary(self):
        summary = self.df.groupby(['Region', 'Product'])['Revenue'].sum().reset_index()
        summary.to_excel(self.writer, sheet_name='Summary', index=False)
        worksheet = self.writer.sheets['Summary']
        auto_size_columns(worksheet, summary)

        currency_fmt = self.workbook.add_format({'num_format': '$#,##0.00'})
        worksheet.set_column('C:C', 15, currency_fmt)

        # Conditional formatting
        worksheet.conditional_format('C2:C100', {
            'type': 'top',
            'value': 5,
            'format': self.workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        })

        # Logo
        if self.logo_path and os.path.exists(self.logo_path):
            worksheet.insert_image('E1', self.logo_path, {'x_scale': 0.5, 'y_scale': 0.5})

    def write_pivot(self):
        self.pivot_df = self.df.pivot_table(
            index='Region',
            columns='Product',
            values='Revenue',
            aggfunc='sum',
            fill_value=0
        )
        self.pivot_df.to_excel(self.writer, sheet_name='Pivot', startrow=0, startcol=0)
        worksheet = self.writer.sheets['Pivot']
        auto_size_columns(worksheet, self.pivot_df.reset_index())

    def insert_chart(self):
        pivot_ws = self.writer.sheets['Pivot']
        chart = self.workbook.add_chart({'type': 'column'})
        num_regions = self.pivot_df.shape[0]

        for i, product in enumerate(self.pivot_df.columns):
            chart.add_series({
                'name': ['Pivot', 0, i + 1],  # Header in row 0
                'categories': ['Pivot', 1, 0, num_regions, 0],  # Region names from row 1 to N
                'values': ['Pivot', 1, i + 1, num_regions, i + 1],  # Revenue values
            })

        chart.set_title({'name': 'Revenue by Region and Product'})
        chart.set_x_axis({'name': 'Region'})
        chart.set_y_axis({'name': 'Revenue'})
        chart.set_legend({'position': 'bottom'})
        pivot_ws.insert_chart('F2', chart)
        
    def export_for_powerbi(self):
        summary = self.df.groupby(['Region', 'Product'])['Revenue'].sum().reset_index()
        pivot = self.df.pivot_table(
            index='Region',
            columns='Product',
            values='Revenue',
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        export_dir = os.path.join(self.output_dir, 'powerbi')
        os.makedirs(export_dir, exist_ok=True)

        summary.to_csv(os.path.join(export_dir, 'summary.csv'), index=False)
        pivot.to_csv(os.path.join(export_dir, 'pivot.csv'), index=False)

