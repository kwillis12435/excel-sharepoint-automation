import unittest
from src.excel.parser import ExcelParser
from src.excel.compiler import ResultsCompiler

class TestExcelAutomation(unittest.TestCase):

    def setUp(self):
        self.parser = ExcelParser()
        self.compiler = ResultsCompiler()

    def test_parse_file(self):
        # Test parsing of a sample Excel file
        data = self.parser.parse_file('path/to/sample.xlsx')
        self.assertIsNotNone(data)
        self.assertIsInstance(data, list)

    def test_extract_data(self):
        # Test extraction of data from parsed content
        sample_data = [{'column1': 'value1', 'column2': 'value2'}]
        extracted_data = self.parser.extract_data(sample_data)
        self.assertEqual(len(extracted_data), 1)
        self.assertIn('column1', extracted_data[0])
        self.assertIn('column2', extracted_data[0])

    def test_compile_results(self):
        # Test compiling results from multiple data sources
        data_sources = [
            [{'column1': 'value1', 'column2': 'value2'}],
            [{'column1': 'value3', 'column2': 'value4'}]
        ]
        compiled_results = self.compiler.compile_results(data_sources)
        self.assertEqual(len(compiled_results), 2)

    def test_save_results(self):
        # Test saving compiled results to a file
        compiled_data = [{'column1': 'value1', 'column2': 'value2'}]
        result = self.compiler.save_results(compiled_data, 'path/to/output.xlsx')
        self.assertTrue(result)

if __name__ == '__main__':
    unittest.main()