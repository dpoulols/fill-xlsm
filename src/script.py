# -*- coding: utf-8 -*-

import argparse

from handler import xlsmHandler

def start():

    #############ARG_parse#############

    parser = argparse.ArgumentParser(prog='script',description='Fill with data a xlsm template.')

    parser.add_argument('file',help='Archivo de origen XLSM')

    args=parser.parse_args()

    ###################################

    command = xlsmHandler(args.file)

    command.run()

if __name__ == '__main__':
    start()
