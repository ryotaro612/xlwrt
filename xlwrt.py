#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re
def to_row_num(row_alpha):
    i =  0
    row = 0
    for l in [(ord(a)-ord('A') + 1) for a in row_alpha[::-1]]:
        row += l * pow(26,i)
        i+=1
    return row

# Deprecated
def to_row_alph(row_num):
    l = row_num
    third = l // pow(26,2)
    l -= third * pow(26,2)
    second = l // 26
    l -= second * 26
    return ''.join(map(lambda n: "" if n == 0 else chr(n + ord('A') - 1), [third, second, l]))


def filled_lines(alpha, size):
    a = to_row_num(alpha)
    return [to_row_alph(c+a) for c in range(size)]

#print (to_row_alph(to_row_num('Z')))
#print (filled_lines('Z', 3))

# http://openpyxl.readthedocs.org/en/default/tutorial.html

import sys
import csv
lines = [ l for l in csv.reader(sys.stdin.readlines())]
row_siz=len(lines)
col_siz=len(lines[0])

