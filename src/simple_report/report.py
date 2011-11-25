#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''

import abc

from interface import ISpreadsheetSection, ISpreadsheetReport,\
    IDocumentReport
from utils import FileProxy, FileConverter
    
                
class Report(object):
    """
    Абстрактный класс отчета
    """
    
    __meta__ = abc.ABCMeta
    
    def __init__(self, ffile):
        """
        """
        self.file = FileProxy(ffile)
        
        
class DocumentReport(Report, IDocumentReport):

    DOCX = FileConverter.DOCX

    def build(self, params, file_type=DOCX):
        """
        Генерирует выходной файл в нужном формате
        """
    
    
class SpreadsheetSection(ISpreadsheetSection):
    """
    Объект - секция отчета
    """    
    
    def __init__(self, report):
        assert isinstance(report, SpreadsheetReport)
        self._report = report
    
    def flush(self, params, oriented=ISpreadsheetSection.VERTICAL):
        """
        Установка параметров в секции
        """
        
        
class SpreadsheetReport(Report, ISpreadsheetReport):
    
    XLSX = FileConverter.XLSX
    
    def __init__(self, *args, **kwargs):
        super(SpreadsheetReport, self).__init__(*args, **kwargs)
        xlsx_file = FileConverter(self.file).build(self.XLSX)
    
    def get_sections(self):
        """
        Возвращает все секции
        """
    
    def get_section(self, section_name):
        """
        Возвращает секцию по имени
        """
        
    
    def build(self, file_type=XLSX):
        """
        Генерирует выходной файл в нужном формате
        """
