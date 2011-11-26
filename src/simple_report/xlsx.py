#coding: utf-8
'''
Created on 24.11.2011

@author: prefer
'''
import os
import uuid
import shutil
from tempfile import gettempdir

from utils import ZipProxy, FileProxy


class XLSXWrapper(object):
    u"""
    """
    def __init__(self, ffile):
        assert isinstance(ffile, FileProxy)       
        
        self.extract_folder = os.path.join( gettempdir(), 
                    '_'.join([str(uuid.uuid4())[:8], ffile.get_file_name()]))
        
        ZipProxy.to_extract(ffile.get_path(), self.extract_folder)
        
    def pack(self, dst_file):
        u"""
        Запаковать в файл
        """
        assert isinstance(dst_file, FileProxy)
        
        ZipProxy.to_pack(dst_file.get_path(), self.extract_folder)
        shutil.rmtree(self.extract_folder)
    
    
class DOCXWrapper(object):
    """
    """