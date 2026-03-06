import os
import sys
import unittest
import shutil

sys.path.insert(0, "./program")
import FileMangement as FM

class test_read_files(unittest.TestCase):
   
    global file_pth 
    file_pth = "/Test"

    #TODO implement
    def test_content_as_expected(self):
      self.fail("not implemented")
      
    
    
    def test_if_no_file_return_dict(self):

        expected = {}

        res = FM.Get_locs(file_pth + "/Does_not_exist.pkl")
        
        self.assertEqual(res, expected)


class test_write_files(unittest.TestCase):
   
    global file_pth 
    file_pth = "Test/tempTestFiles"
    global mockData
    mockData = {'name': 'Alice', 'age': '25', 'city': 'New York'}

    ### --- Arrange --- ###
    def setUp(self):
        os.mkdir(file_pth)

    def tearDown(self):
        shutil.rmtree(file_pth)

    def test_if_file_gets_created(self):
      ### --- Arrange --- ###
      filename = file_pth + "/person.pkl"

      ### --- Act --- ###
      FM.Set_locs(mockData, filename)

      ### --- Assert --- ###
      self.assertTrue(os.path.exists(filename))

