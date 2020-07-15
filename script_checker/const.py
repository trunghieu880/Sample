# -*- coding: utf-8 -*-

import logging
from pathlib import Path

NAME = 'Checker'

LOGS = HOME.joinpath('logs', 'messages')

ASSET = Path(__file__).parent.joinpath('assets')
SETTING = ASSET.joinpath('settings.json')
VERSION = ASSET.joinpath('version.json')
TEMPLATE = Path(__file__).parent.joinpath('template')
