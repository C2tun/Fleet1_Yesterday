from datetime import datetime, timedelta
import pytz

import os
from tkcalendar import DateEntry
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import StringVar, ttk

from influxdb_client import InfluxDBClient, Point, WritePrecision
from influxdb_client.client.write_api import SYNCHRONOUS

import pandas as pd
import plotly.express as px
import pytz
import csv
import datetime

import openpyxl
from datetime import datetime, timedelta

import threading
import time

from itertools import groupby
start_datetime = None
end_datetime = None

output_file = 'CELL_OUPUT.csv'

DataBP = False

token = "O5BEs5j6viF_UCR2Ll5cYAHJVXtN8jyZoD4ZjLtTvpQRWLyjoVXzWVfMXXGV1of0wrhrSUKGrpVyWn7g7rhRYQ=="
org = "DataLogger"
bucket = "DataLogger"
client = InfluxDBClient(url="https://datalogger-influxdb.minesmartferry.com", token=token)


# Specify the local time zone (GMT+7 in this case)
local_tz = pytz.timezone('Asia/Bangkok')  # Replace with your actual time zone

# Get the current date and time in the local time zone
current_datetime = datetime.now(local_tz)

# Calculate yesterday's date and time
yesterday_datetime = current_datetime - timedelta(days=1)

# Set the desired time for yesterday (22:00:00.000 in local time)
desired_time1 = yesterday_datetime.replace(hour=22, minute=0, second=0, microsecond=0)

# Set another desired time (05:00:00.000 in local time)
desired_time2 = yesterday_datetime.replace(hour=5, minute=0, second=0, microsecond=0)

# Create a naive datetime object without a time zone
naive_datetime = datetime(desired_time1.year, desired_time1.month, desired_time1.day, desired_time1.hour, desired_time1.minute, desired_time1.second, desired_time1.microsecond)

# Localize the naive datetime to the specified time zone (GMT+7)
localized_desired_time = local_tz.localize(naive_datetime)

# Convert the localized time to UTC
utc_time = localized_desired_time.astimezone(pytz.utc)

# Create a naive datetime object without a time zone for desired_time2
naive_datetime2 = datetime(desired_time2.year, desired_time2.month, desired_time2.day, desired_time2.hour, desired_time2.minute, desired_time2.second, desired_time2.microsecond)

# Localize the naive datetime to the specified time zone (GMT+7)
localized_desired_time2 = local_tz.localize(naive_datetime2)

# Convert the localized time to UTC for desired_time2
utc_time2 = localized_desired_time2.astimezone(pytz.utc)

# Format both UTC datetimes as strings
formatted_utc_time1 = utc_time.strftime('%Y-%m-%d %H:%M:%S.%f')  # Include microseconds for desired_time1
formatted_utc_time2 = utc_time2.strftime('%Y-%m-%d %H:%M:%S.%f')  # Include microseconds for desired_time2

print("Original Date and Time 1 (Local):", desired_time1)
print("Equivalent Date and Time 1 in UTC:", formatted_utc_time1)

print("Original Date and Time 2 (Local):", desired_time2)
print("Equivalent Date and Time 2 in UTC:", formatted_utc_time2)

# Replace ".000000" with "Z"
end_time1 = formatted_utc_time1.replace(".000000", "Z")
start_time1 = formatted_utc_time2.replace(".000000", "Z")

end_t = end_time1.replace(" ","T")
start_t = start_time1.replace(" ","T")
print("Original Date and Time 1 (Local):", desired_time1)
print("Equivalent Date and Time 1 in UTC:", end_t)

print("Original Date and Time 2 (Local):", desired_time2)
print("Equivalent Date and Time 2 in UTC:", start_t)


def Get_ferry_id():
    global ferry_ids
    ferry_ids =[] 
      
    query3  = f' from(bucket:"DataLogger")\
    |> range(start:{start_t}, stop:{end_t})\
    |> filter(fn:(r) => r._measurement == "sbcu")\
    |> filter(fn:(r) => r._field == "0x180a0001_S1_BatPack_Current" )\
    |> group(columns: ["ferry_id"])\
    |> distinct(column: "ferry_id")'

    resulta = client.query_api().query(org = org,query=query3)
    ferry_ids =[]
    for table in resulta:
        for record in table.records:
            # Get the tag value from the "_measurement" column
            ferry_id = record.get_value()
            ferry_ids.append(ferry_id)    
            print(ferry_id)        
    return ferry_ids

ferry = Get_ferry_id()
print(ferry)
