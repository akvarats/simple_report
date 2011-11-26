#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''

import abc

from interface import ISpreadsheetSection, ISpreadsheetReport,\
    IDocumentReport
from utils import FileProxy, FileConverter
from xlsx import XLSXWrapper
    
                
class Report(object):
    u"""
    Абстрактный класс отчета
    """
    
    __meta__ = abc.ABCMeta
    
    def __init__(self, ffile):
        """
        """
        self.file = FileProxy(ffile)
        
        
class DocumentReport(Report, IDocumentReport):

    DOCX = FileConverter.DOCX

    def build(self, dst_file_path, params, file_type=DOCX):
        u"""
        Генерирует выходной файл в нужном формате
        """
    
    
class SpreadsheetSection(ISpreadsheetSection):
    u"""
    Объект - секция отчета
    """    
    
    def __init__(self, report):
        assert isinstance(report, SpreadsheetReport)
        self._report = report
    
    def flush(self, params, oriented=ISpreadsheetSection.VERTICAL):
        u"""
        Установка параметров в секции
        """
        
        
class SpreadsheetReport(Report, ISpreadsheetReport):
    
    XLSX = FileConverter.XLSX
    
    def __init__(self, *args, **kwargs):
        super(SpreadsheetReport, self).__init__(*args, **kwargs)
        
        xlsx_file = FileConverter(self.file).build(self.XLSX)
        
        self.report = XLSXWrapper( FileProxy(xlsx_file) )
    
    def get_sections(self):
        u"""
        Возвращает все секции
        """
        
        return self.report.get_sections()
    
    def get_section(self, section_name):
        u"""
        Возвращает секцию по имени
        """
        self.report.get_section(section_name)
        
    
    def build(self, dst_file_path, file_type=XLSX):
        u"""
        Генерирует выходной файл в нужном формате
        """
        ffile = FileProxy(dst_file_path, new_file=True)
        self.report.pack(ffile)
        return FileConverter(ffile).build(file_type)
