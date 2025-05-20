from datetime import datetime
import os

def get_today_string(fmt="%Y_%m_%d"):
    """Return today's date as a string in given format."""
    return datetime.today().strftime(fmt)

def currency_format(workbook):
    """Return a standard currency cell format."""
    return workbook.add_format({'num_format': '$#,##0.00'})

def top_5_highlight_format(workbook):
    """Return a format for conditional highlighting (top 5 values)."""
    return workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})

def ensure_dir(path):
    """Create directory if it doesn't exist."""
    os.makedirs(path, exist_ok=True)
    
def auto_size_columns(worksheet, dataframe):
    """Auto-size columns based on max data length in a dataframe."""
    for i, column in enumerate(dataframe.columns):
        max_length = max(
            dataframe[column].astype(str).map(len).max(),
            len(column)
        ) + 2  # optional padding
        worksheet.set_column(i, i, max_length)
