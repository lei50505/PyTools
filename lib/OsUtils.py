#! /usr/bin/env python
# -*- coding: utf-8 -*-

import sys
sys.path.append(".")

import os

def osPathIsFile(path):
    return os.path.isfile(path)

def removeFile(filePath):
    if os.path.isfile(filePath):
        os.remove(filePath)
