import unittest
from src.sharepoint.connector import SharePointConnector

class TestSharePointConnector(unittest.TestCase):

    def setUp(self):
        self.connector = SharePointConnector()

    def test_connect(self):
        result = self.connector.connect()
        self.assertTrue(result)

    def test_list_files(self):
        files = self.connector.list_files()
        self.assertIsInstance(files, list)

    def test_download_file(self):
        file_name = "test_file.xlsx"
        result = self.connector.download_file(file_name)
        self.assertTrue(result)

if __name__ == '__main__':
    unittest.main()