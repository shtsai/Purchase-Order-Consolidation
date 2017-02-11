# -*- coding: utf-8 -*-

def eval0(num):
    '''
    This function parse the following non-standard integer characters
    '''
    try:
        result = eval(num)
    except SyntaxError:
        if num == "１":
            result = 1
        elif num == "２":
            result = 2
        elif num == "３":
            result = 3
        elif num == "４":
            result = 4
        elif num == "５":
            result = 5
        elif num == "６":
            result = 6
        else:
            return 0
    return result

