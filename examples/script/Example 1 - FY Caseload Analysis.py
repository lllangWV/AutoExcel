from datetime import datetime
from autoexcel.main import fy_analysis
from autoexcel.utils.config import config
import os


config.logging_config.loggers.autoexcel.level='DEBUG'
config.apply()


raw_xlsx = os.path.join(config.data_dir, 'fy_analysis', 'raw', 'Raw Data 9-19-2024.xlsx')
template_xlsx = os.path.join(config.data_dir, 'fy_analysis', 'templates', 'template_processed_workbook.xlsx')
processed_dir = os.path.join(config.data_dir, 'fy_analysis', 'processed')
assigned_date_filter=[datetime(2023, 7, 1), None]

fy_analysis(raw_xlsx, template_xlsx, processed_dir=processed_dir, assigned_date_filter=assigned_date_filter)