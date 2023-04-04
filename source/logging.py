# This module handles logging
import sys
from pathlib import Path
import datetime
import os
import os.path

def start(error=True, add_time=True, add_name=True, plan_breaks=True, change_task=True, export_to_excel=True, start_stop=True):
    # start logging

    if error:
    # error logging
        path = Path(__file__).parent / "log"
        try:
            files = os.listdir(path)
            for file in files:
                size = os.path.getsize(f'{path}/{file}')
                if size == 0:
                    os.remove(f'{path}/{file}')
        except:
            pass
        try:
            os.mkdir(path)
        except:
            pass
        time = datetime.datetime.now()
        file = f'error{time.year}-{time.month}-{time.day}_{time.hour}-{time.minute}-{time.second}.log'
        sys.stderr = open(f'{path}/{file}', 'w')

    log = {}
    path = Path(__file__).parent / "log"
    time = datetime.datetime.now()
    file = f'log{time.year}-{time.month}-{time.day}.log'
    log_something = False

    if add_time:
        log_something = True
    elif add_time:
        log_something = True
    elif plan_breaks:
        log_something = True
    elif change_task:
        log_something = True
    elif export_to_excel:
        log_something = True
    elif start_stop:
        log_something = True

    if log_something:
        try:
            os.mkdir(path)
        except:
            pass
        log_file = open(f'{path}/{file}', 'a')

    log['add_time'] = add_time
    log['add_name'] = add_name
    log['plan_breaks'] = plan_breaks
    log['change_task'] = change_task
    log['export_to_excel'] = export_to_excel
    log['start_stop'] = start_stop

    return log, log_file

