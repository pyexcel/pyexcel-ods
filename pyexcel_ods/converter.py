import sys
import datetime

PY2 = sys.version_info[0] == 2


def float_value(value):
    """convert a value to float"""
    ret = float(value)
    return ret


def date_value(value):
    """convert to data value accroding ods specification"""
    ret = "invalid"
    try:
        # catch strptime exceptions only
        if len(value) == 10:
            ret = datetime.datetime.strptime(
                value,
                "%Y-%m-%d")
            ret = ret.date()
        elif len(value) == 19:
            ret = datetime.datetime.strptime(
                value,
                "%Y-%m-%dT%H:%M:%S")
        elif len(value) > 19:
            ret = datetime.datetime.strptime(
                value[0:26],
                "%Y-%m-%dT%H:%M:%S.%f")
    except ValueError:
        pass
    if ret == "invalid":
        raise Exception("Bad date value %s" % value)
    return ret


def ods_date_value(value):
    return value.strftime("%Y-%m-%d")


def time_value(value):
    """convert to time value accroding the specification"""
    import re
    results = re.match('PT(\d+)H(\d+)M(\d+)S', value)
    hour = int(results.group(1))
    minute = int(results.group(2))
    second = int(results.group(3))
    if hour < 24:
        ret = datetime.time(hour, minute, second)
    else:
        ret = datetime.timedelta(hours=hour, minutes=minute, seconds=second)
    return ret


def ods_time_value(value):
    return value.strftime("PT%HH%MM%SS")


def boolean_value(value):
    """get bolean value"""
    if value == "true":
        ret = True
    else:
        ret = False
    return ret


def ods_bool_value(value):
    """convert a boolean value to text"""
    if value is True:
        return "true"
    else:
        return "false"


def ods_timedelta_value(cell):
    """convert a cell value to time delta"""
    hours = cell.days * 24 + cell.seconds // 3600
    minutes = (cell.seconds // 60) % 60
    seconds = cell.seconds % 60
    return "PT%02dH%02dM%02dS" % (hours, minutes, seconds)


ODS_FORMAT_CONVERSION = {
    "float": float,
    "date": datetime.date,
    "time": datetime.time,
    'timedelta': datetime.timedelta,
    "boolean": bool,
    "percentage": float,
    "currency": float
}


ODS_WRITE_FORMAT_COVERSION = {
    float: "float",
    int: "float",
    str: "string",
    datetime.date: "date",
    datetime.time: "time",
    datetime.timedelta: "timedelta",
    bool: "boolean"
}

if PY2:
    ODS_WRITE_FORMAT_COVERSION[unicode] = "string"

VALUE_CONVERTERS = {
    "float": float_value,
    "date": date_value,
    "time": time_value,
    "timedelta": time_value,
    "boolean": boolean_value,
    "percentage": float_value,
    "currency": float_value
}

ODS_VALUE_CONVERTERS = {
    "date": ods_date_value,
    "time": ods_time_value,
    "boolean": ods_bool_value,
    "timedelta": ods_timedelta_value
}


VALUE_TOKEN = {
    "float": "value",
    "date": "date-value",
    "time": "time-value",
    "boolean": "boolean-value",
    "percentage": "value",
    "currency": "value",
    "timedelta": "time-value"
}
