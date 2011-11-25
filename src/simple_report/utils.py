#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''


import zipfile

class ZipProxy(object):
    """
    Распаковка/упаковка Open XML
    """
    def __init__(self, ffile):
        self.file = FileProxy(ffile)
        
    def to_extract(self, path):
        """
        Распоковывает zip файл
        """
        
    def to_pack(self, file_name, path):
        """
        Запаковывает zip файл
        """
        
class FileConverter(object):
    """
    Конвертирует файл из одного формата в другой
    """
    
    # open xml formats
    DOCX = 'docx'
    XLSX = 'xlsx'
    
    HTML = 'html'
    DOC = 'doc'
    XLS = 'xls'
    PDF = 'pdf'
    RTF = 'rtf'
    
    # OpenOffice Formats
    ODT = 'odt'
    ODS = 'ods'
    ODF = 'odf'
        
    def __init__(self, ffile):
        self.file = FileProxy(ffile)
    
    def build(self, to_format):
        pass


    def xls2xlsx(self): pass
    
    def xlsx2xls(self): pass
    def xlsx2pdf(self): pass
    def xlsx2html(self): pass
    


class FileProxy(object):
    """
    """
    def __init__(self, file_like_object):
        assert file_like_object
        
        self.file = file_like_object
    
    def get_path(self):
        """
        Возвращает путь до файла
        """
    
    def get_file_like_object(self):
        """
        Возвращает открытый файл
        """
    
    
         
        
        