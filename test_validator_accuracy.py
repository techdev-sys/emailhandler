import unittest
from unittest.mock import MagicMock, patch
import sys
import os

# Mock zipfile before importing excel_validator because it imports zipfile at top level
sys.modules['zipfile'] = MagicMock()

# Import the function to test
# We need to be careful about the import because we mocked zipfile
import excel_validator

class TestBSDValidator(unittest.TestCase):

    def setUp(self):
        # Reset the mock for each test
        self.mock_zip = sys.modules['zipfile'].ZipFile.return_value
        self.mock_enter = self.mock_zip.__enter__.return_value

    def create_mock_xml_content(self, text):
        # Helper to mock the read().decode() chain
        mock_file = MagicMock()
        mock_file.read.return_value.decode.return_value = text
        return mock_file

    @patch('excel_validator.is_actually_bsd4')
    def test_bsd4_rejection_by_filename(self, mock_is_bsd4):
        # Should reject immediately if filename says BSD 4
        self.assertFalse(excel_validator.is_valid_bsd_return("Report_BSD4.xlsx"))
        self.assertFalse(excel_validator.is_valid_bsd_return("BSD 4 Return.xlsx"))

    def test_bsd2_acceptance_by_content(self):
        # Scenario: Filename is generic, but content has BSD 2 title
        filename = "Generic_Return.xlsx"
        
        # content with valid BSD 2 markers
        xml_content = """
        <t>FOREIGN EXCHANGE MARKET ACTIVITY</t>
        <t>Some other data</t>
        """
        
        # Setup mock behavior
        self.mock_enter.namelist.return_value = ['xl/sharedStrings.xml']
        self.mock_enter.open.return_value.__enter__.return_value = self.create_mock_xml_content(xml_content)

        # We need to bypass the is_actually_bsd4 check which tries to open the file again
        # Ideally we'd mock that too, but for integration logic:
        with patch('excel_validator.is_actually_bsd4', return_value=False):
            result = excel_validator.is_valid_bsd_return(filename)
            self.assertTrue(result, "Should accept file with 'FOREIGN EXCHANGE MARKET ACTIVITY'")

    def test_bsd4_rejection_by_marker(self):
        # Scenario: Content has BSD 4 specific headers
        filename = "Suspicious_File.xlsx"
        
        xml_content = """
        <t>ASSETS</t>
        <t>LIABILITIES</t>
        <t>NET FX ASSETS</t>
        <t>BSD 4</t>
        """
        
        self.mock_enter.namelist.return_value = ['xl/sharedStrings.xml']
        self.mock_enter.open.return_value.__enter__.return_value = self.create_mock_xml_content(xml_content)

        # The validator calls is_actually_bsd4 first. 
        # But let's assume it passes the first check and goes to main logic
        # We need to assert that the main logic ALSO catches it if it sees "BSD 4" inside
        
        with patch('excel_validator.is_actually_bsd4', return_value=False):
            # If is_actually_bsd4 returns False (missed it), the main loop should still catch "BSD 4" in content
            # *However*, our current code calls is_actually_bsd4. logic is split.
            # actually, our updated code has `if "BSD 4" in content: return False` inside the main loop too.
            
            result = excel_validator.is_valid_bsd_return(filename)
            self.assertFalse(result, "Should reject content containing 'BSD 4'")

    def test_ambiguous_rejection(self):
        # Scenario: File has 'SPOT' but no 'BSD 2' or 'BSD 3' title
        filename = "Ambiguous.xlsx"
        
        xml_content = """
        <t>SPOT TRANSACTIONS</t>
        <t>Total</t>
        """
        
        self.mock_enter.namelist.return_value = ['xl/sharedStrings.xml']
        self.mock_enter.open.return_value.__enter__.return_value = self.create_mock_xml_content(xml_content)

        with patch('excel_validator.is_actually_bsd4', return_value=False):
            result = excel_validator.is_valid_bsd_return(filename)
            self.assertFalse(result, "Should reject file that only has 'SPOT TRANSACTIONS' without specific BSD 2/3 title")

if __name__ == '__main__':
    unittest.main()
