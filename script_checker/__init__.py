# -*- coding: utf-8 -*-

import json
import logging
import shutil
from datetime import datetime
from pathlib import Path

import const as CONST

try:
    with open(CONST.CONFIG) as fp:
        config = json.load(fp)
except:
    config = {}

# Logging
CONST.LOGS.parent.mkdir(parents=True, exist_ok=True)
if Path(CONST.LOGS).is_file() and Path(CONST.LOGS).stat().st_size > 9000000:
    now = datetime.now().strftime("%y%m%d%H%M")
    old = "{0}.{1}".format(str(CONST.LOGS), now)
    shutil.move(CONST.LOGS, old)

logging.basicConfig(
    level=config.get('logging', logging.INFO),
    format='%(asctime)s %(lineno)4s:%(name)-12s %(levelname)-8s %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler(CONST.LOGS),
        logging.StreamHandler()
    ]
)