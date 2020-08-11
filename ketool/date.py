#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2020-08-12 01:55:53
# @Author  : 流柯
# @Version : V
# @desc :

import time

# Gets the current timestamp


def timeStamp(ms=True):
    '''ms is True return Millisecond time stamp,default  True'''
    if ms:
        return int(time.time() * 1000)
    else:
        return int(time.time())
