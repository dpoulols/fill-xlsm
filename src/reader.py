# -*- coding: utf-8 -*-

from openpyxl import load_workbook
from openpyxl import Workbook
from datetime import datetime as date

import logging

class Reader:
    
    def __init__(self,path: str = None):
        self.path = path

        self.__wb = None
        self.__cs = None
        self.__values = ["¿Sabias que?", "Antonio","es","una","pedazo","de","foca","?","13","14","15","16","17",self.todaysDate(),"19","20","21","22"]

        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.DEBUG)
    
    @property
    def wb(self):
        return self.__wb
    
    @property
    def isOpen(self):
        return self.__wb is not None
    
    def __reset(self):
        if self.__cs is not None:
            self.__cs = None
        
        if self.__wb is not None:
            self.__wb.close()
            self.__wb = None
        
    
    def tryToOpen(self,path: str = None):

        if path is not None:
            self.path = path
        
        if self.path is None:
            raise Exception('Ruta no definida')
        
        self.__wb = load_workbook(self.path, read_only=False)


    def execute(self):

        if self.isOpen is False:
            raise Exception('Archivo no abierto')
        
        self.printValues()


    def save(self,path: str):
        
        self.__wb.save(path)



    def todaysDate(self):
        return date.today().strftime('%d/%m/%Y')

    
    def printValues(self):
        #le devuelvo un diccionario
        self.__cs = self.__wb.active
        row = 5
        for value in self.__values:
            cell = "B" + str(row)
            self.__cs[cell] = value
            row += 1
