import os
from datetime import datetime

from autoexcel.main import fy_analysis
from autoexcel.utils.config import config

config.logging_config.loggers.autoexcel.level = "INFO"
config.apply()

old_negotiators = [
    "Abigail Gallagher",
    "Eric Divito",
    "Eric Winaught",
    "Jillian Corbett",
    "Huron – New",
]

raw_dir = os.path.join(config.data_dir, "fy_analysis", "raw")
processed_dir = os.path.join(config.data_dir, "fy_analysis", "processed")

assigned_date_filter = [datetime(2023, 7, 1), None]

fy_analysis(
    raw_dir,
    processed_dir,
    assigned_date_filter=assigned_date_filter,
    old_negotiators=old_negotiators,
)
