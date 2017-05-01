#! /usr/bin/env python
# -*- coding: utf-8 -*-

import sys
sys.path.append(".")

import codecs

def openFile(filePath,operType):
    return codecs.open(filePath, mode=operType, encoding="UTF-8", \
                       errors='strict', buffering=1)
