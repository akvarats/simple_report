#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''

from abc import ABCMeta, abstractmethod


class IReport(object):
    
    @DeprecationWarning    
    def show(self, *args, **kwargs):
        self.build(*args, **kwargs)
        
    @abstractmethod
    def build(self, *args, **kwargs):
        pass


class IDocumentReport(IReport):
    __meta__ = ABCMeta
    
    @abstractmethod
    def build(self, params, file_type):
        """
        Генерирует выходной файл в нужном формате
        """
    
    
class ISpreadsheetReport(IReport):
    __meta__ = ABCMeta

    @abstractmethod
    def get_sections(self):
        """
        Возвращает все секции
        """
    
    @abstractmethod
    def get_section(self, section_name):
        """
        Возвращает секцию по имени
        """
        
    @abstractmethod
    def build(self, file_type):
        """
        Генерирует выходной файл в нужном формате
        """
        
        
class ISpreadsheetSection(object):
    __meta__ = ABCMeta
    
        
    # Тип разворота секции
    VERTICAL = 0
    GORIZONTAL = 1
    
    @abstractmethod
    def flush(self, params, oriented=VERTICAL):
        pass