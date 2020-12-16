#!/usr/bin/python3
# -*- coding: utf-8 -*-

import logging

FILE_FORMATER = logging.Formatter(
        fmt="%(asctime)s %(levelname)s: %(funcName)s - %(message)s",
        datefmt="%d-%m-%Y %H:%M:%S"
    )
CONSOLE_FORMATER = logging.Formatter(fmt="%(levelname)s: %(lineno)d: %(funcName)s - %(message)s")
LOG_FILE = "info.log"


def get_logger(logger_name):
    _log = logging.getLogger(logger_name)
    _log.addHandler(_configure_file_logging())
    _log.addHandler(_configure_stream_logging())
    return _log


def _configure_file_logging():
    file_handler = logging.FileHandler(filename=LOG_FILE, encoding="utf8")
    file_handler.setFormatter(FILE_FORMATER)
    file_handler.setLevel(logging.DEBUG)
    return file_handler


def _configure_stream_logging():
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(CONSOLE_FORMATER)
    stream_handler.setLevel(logging.INFO)
    return stream_handler
