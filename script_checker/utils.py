import json
import logging
import os
import sys
from datetime import datetime
from pathlib import Path
import shutil

logger = logging.getLogger(__name__)
def load(path, keys=''):
    '''Load data from json file'''
    logger.debug("Load data from %s", Path(path).name)

    def filter(data):
        for k in keys.split('.'):
            data = data.get(k, {}) if k != '' else data
        return data
    try:
        keys = '' if keys is None else keys.strip()

        with open(path, encoding='shift-jis', errors='ignore') as fp:
            data = json.load(fp)
    except Exception as e:
        data = {}
        if Path(path).is_file() is True:
            logger.exception(e)
    finally:
        return filter(data)

def write(data, path):
    '''Write dict to json'''
    logger.debug("Write data to %s", Path(path).name)
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, encoding='shift-jis', errors='ignore', mode='w') as fp:
        json.dump(data, fp, indent=4, sort_keys=True)


def read_file(path):
    with open(path, encoding='shift-jis', errors='ignore') as fp:
        return fp.readlines()

def filter_tbl(data, opts):
    '''Filter rows from table data'''
    def match(row):
        dct = dict(zip(data[0], row))
        return all([dct.get(col) == value for col, value in opts.items()])

    logger.debug("Filter table by %s", opts)
    return [row for row in data if match(row)]

def copy(src, dst):
    '''Copy file'''
    logger.debug("Copy %s to %s", Path(src).name, Path(dst).parent.name)
    Path(dst).parent.mkdir(parents=True, exist_ok=True)
    if Path(dst).absolute() != Path(src).absolute():
        shutil.copy2(src, dst)

def delete(filepath):
    '''Delete file if exist'''
    if Path(filepath).is_file():
        Path(filepath).unlink()

def get_label_list(labels=[]):
    '''Get label list'''
    logger.debug("Get label list")
    try:
        lst = load(CONST.SETTING).get('labelList')
        lst_count = [len(set(labels) & set(l)) for l in lst]
        index = lst_count.index(max(lst_count))
        data = lst[index]
    except Exception as e:
        logger.exception(e)
        data = []
    finally:
        return data

def collapse_list(lst):
    '''Collapse the list of number'''
    def text(a, b):
        return str(a) if a == b else '{0}~{1}'.format(a, b)

    if len(lst) == 0:
        return ''
    lst.sort()
    rst = []
    t = 0
    for i in range(len(lst)):
        if i > 1 and lst[i] - lst[i-1] > 1:
            rst.append(text(lst[t], lst[i-1]))
            t = i
    rst.append(text(lst[t], lst[i]))

    return ', '.join(rst)

def is_simulink(path):
    '''Check source code is Simulink model'''
    try:
        rst = False
        with open(path, encoding='shift-jis', errors='ignore') as fp:
            for line in fp.readlines()[:100]:
                if 'Simulink model' in line:
                    rst = True
                    break
    except:
        rst = None
    finally:
        return rst


def scan_files(directory, ext='.txt'):
    '''Scan all file that has extension in directory'''
    logger.debug("Scan directory %s %s", directory, ext)
    data = []
    latest = None
    for root, _, files in os.walk(directory):
        for filename in files:
            if filename.endswith(ext):
                filepath = Path(root).joinpath(filename)
                data.append(filepath)

                # if (latest is None or
                        # Path(latest).stat().st_mtime < filepath.stat().st_mtime):
                    # latest = filepath

    return data, latest
