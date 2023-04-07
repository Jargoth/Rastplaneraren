# This module handles logging
import sys
from pathlib import Path
import datetime
import os
import os.path

def start(error=True, add_time=True, add_name=True, plan_breaks=True, change_task=True, export_to_excel=True,
          start_stop=True, set_default_task=True, show_announcements=True, delete_announcements=True, load_task=True,
          load_breaks=True, load_workers_min=True, load_working_time=True, load_employees=True, load_announcements=True,
          load_excel_templates=True, load_excel_selected=True, simultaneous_breaks=True, workers_minimum_override=True):
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
    elif set_default_task:
        log_something = True
    elif show_announcements:
        log_something = True
    elif delete_announcements:
        log_something = True
    elif load_task:
        log_something = True
    elif load_breaks:
        log_something = True
    elif load_workers_min:
        log_something = True
    elif load_working_time:
        log_something = True
    elif load_employees:
        log_something = True
    elif load_announcements:
        log_something = True
    elif load_excel_templates:
        log_something = True
    elif load_excel_selected:
        log_something = True
    elif simultaneous_breaks:
        log_something = True
    elif workers_minimum_override:
        log_something = True

    if log_something:
        try:
            os.mkdir(path)
        except:
            pass
        log_file = f'{path}/{file}'

        log['add_time'] = add_time
        log['add_name'] = add_name
        log['plan_breaks'] = plan_breaks
        log['change_task'] = change_task
        log['export_to_excel'] = export_to_excel
        log['start_stop'] = start_stop
        log['set_default_task'] = set_default_task
        log['show_announcements'] = show_announcements
        log['delete_announcements'] = delete_announcements
        log['load_task'] = load_task
        log['load_breaks'] = load_breaks
        log['load_workers_min'] = load_workers_min
        log['load_working_time'] = load_working_time
        log['load_employees'] = load_employees
        log['load_announcements'] = load_announcements
        log['load_excel_templates'] = load_excel_templates
        log['load_excel_selected'] = load_excel_selected
        log['simultaneous_breaks'] = simultaneous_breaks
        log['workers_minimum_override'] = workers_minimum_override
    else:
        return log, None

    return log, log_file

