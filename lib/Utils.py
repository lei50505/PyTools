#! /usr/bin/env python
# -*- coding: utf-8 -*-

import sys
sys.path.append(".")

def isNum(val):
    if val == None:
        return False
    try:
        int(val)
    except:
        try:
            float(val)
        except:
            return False
    return True


