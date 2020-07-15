# -*- coding: utf-8 -*-

import logging
from pathlib import Path

NAME = 'Checker'

HOME = Path.home().joinpath(NAME)
CONFIG = HOME.joinpath('config.json')
LOGS = HOME.joinpath('logs', 'messages')

DATA = HOME.joinpath('db')

ASSET = Path(__file__).parent.joinpath('assets')
SETTING = ASSET.joinpath('settings.json')
VBA = ASSET.joinpath('vba.xlsm')
VERSION = ASSET.joinpath('version.json')
TEMPLATE = Path(__file__).parent.joinpath('template')