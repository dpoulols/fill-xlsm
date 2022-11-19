# -*- coding: utf-8 -*-

import configparser
import logging

import os
import json

from logging import config as logging_config

from reader import Reader

class xlsmHandler:

    def __init__(self, file: str = None):
        self.file = file
    
    def run(self):
        r  = Reader(self.file)

        r.tryToOpen()

        r.execute()

        r.save(self.file)
