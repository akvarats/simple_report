#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''

import os
import zipfile

    
class FileException(Exception):
    pass


class ZipProxy(object):
    u"""
    Распаковка/упаковка Open XML
    """            
    
    @classmethod    
    def to_extract(cls, src_file_path, dst_files_path):
        u"""
        Распоковывает zip файл
        """
        with zipfile.ZipFile(src_file_path, 'r') as zip_file:
            zip_file.extractall(dst_files_path)
        
    @classmethod
    def to_pack(cls, dst_file_path, src_files_path):
        u"""
        Запаковывает zip файл
        """
        with zipfile.ZipFile(dst_file_path, 'w') as zip_file:
            for root, _, file_names in os.walk(src_files_path):
                for file_name in file_names:
                    # Абсолютный путь до файла
                    abs_path = os.path.join(root, file_name)
                    
                    # Путь до директории 
                    dir_path = abs_path[len(src_files_path) + len(os.sep):]
                    
                    zip_file.write(abs_path, dir_path, compress_type=zipfile.ZIP_DEFLATED)
                            
        
class FileConverter(object):
    u"""
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
        assert isinstance(ffile, FileProxy)
        self.file = ffile
    
    def build(self, to_format):
        # Заглушка для конвертера
        return self.file.get_path()


    def xls2xlsx(self): pass
    
    def xlsx2xls(self): pass
    def xlsx2pdf(self): pass
    def xlsx2html(self): pass

class FileProxy(object):
    u"""
    """
    def __init__(self, file_like_object, new_file=False):
        self.is_file_like_object = False
        if hasattr(file_like_object, 'read'):
            raise FileException("File like object temporarily not supported.")
            #self.is_file = True
        
        if not os.path.exists(file_like_object) and not new_file:
            raise FileException('File "%s" not found.' % file_like_object)
            
        if not os.path.isfile(file_like_object) and not new_file:
            raise FileException('"%s" is not file' % file_like_object)
        
        self.file = file_like_object
    
    def get_path(self):
        u"""
        Возвращает путь до файла
        """
        return self.file
    
    def get_file_like_object(self):
        u"""
        Возвращает открытый файл
        """
        return self.file
    
    def get_file_name(self):
        u"""
        Возвращает только имя файла
        """        
        if self.is_file_like_object:
            file_name = self.file.name
        else:
            file_name = self.file

        return os.path.split(file_name)[-1]
        
        