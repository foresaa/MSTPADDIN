from datetime import datetime

def convert_duration_to_days(duration, hours_per_day):
    return duration / (hours_per_day * 60)

def convert_to_string(date):
    if isinstance(date, datetime):
        return date.strftime("%Y-%m-%d %H:%M:%S")
    return date
