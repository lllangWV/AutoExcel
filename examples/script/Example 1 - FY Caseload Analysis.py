from datetime import datetime
from autoexcel.main import fy_analysis
from autoexcel.utils.config import config
import os
import numpy as np

# np.busday_count

config.logging_config.loggers.autoexcel.level='DEBUG'
config.apply()


raw_dir = os.path.join(config.data_dir, 'fy_analysis_test', 'raw')
template_dir = os.path.join(config.data_dir, 'fy_analysis_test', 'templates')
processed_dir = os.path.join(config.data_dir, 'fy_analysis_test', 'processed')
assigned_date_filter=[datetime(2023, 7, 1), None]

fy_analysis(raw_dir, processed_dir, assigned_date_filter=assigned_date_filter)