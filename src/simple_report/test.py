'''
Created on 26.11.2011

@author: prefer
'''

import os
import unittest

from report import SpreadsheetReport

class TestSpreadsheetReport(unittest.TestCase):
    
    def setUp(self):
        self.src_file_path = os.path.join( 
                    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                    'test_data', 'simple_test.xlsx')
        self.dst_file_path = os.path.join(
                    os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 
                    'test_data', 'simple_test_result.xlsx')
        
        print self.dst_file_path

    def test_report(self):        
        if os.path.exists(self.dst_file_path):
            os.remove(self.dst_file_path)
        
        report = SpreadsheetReport(self.src_file_path)
        report.build(self.dst_file_path)
        
        if not os.path.exists(self.dst_file_path):
            raise Exception('File not create')
            
        
if __name__ == '__main__':
    unittest.main()