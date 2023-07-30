import unittest
from unittest.mock import patch
from helper_functions import (
    read_last_roll_number,
    save_last_roll_number,
    get_valid_number,
    map_course_selection,
    map_hostel_requirement
)


class TestMyModule(unittest.TestCase):
    def test_read_last_roll_number(self):
        with patch('builtins.open', mock_open(read_data='10\n')) as mock_file:
            result = read_last_roll_number('file_path')
            self.assertEqual(result, 10)

    def test_save_last_roll_number(self):
        with patch('builtins.open', mock_open()) as mock_file:
            save_last_roll_number('file_path', 20)
            mock_file.assert_called_with('file_path', 'w')
            mock_file().write.assert_called_with('20')

    def test_get_valid_number(self):
        with patch('builtins.input', side_effect=['12345', 'abc', '123456']):
            result = get_valid_number('Enter a 5-digit number: ', 5)
            self.assertEqual(result, '12345')

            # Invalid input, should prompt again
            result = get_valid_number('Enter a 5-digit number: ', 5)
            self.assertEqual(result, '123456')

    def test_map_course_selection(self):
        result = map_course_selection('BSCN')
        self.assertEqual(result, 'BSC Nursing')

        result = map_course_selection('DP')
        self.assertEqual(result, 'DPharma')

        result = map_course_selection('ANM')
        self.assertEqual(result, 'ANM')

    def test_map_hostel_requirement(self):
        result = map_hostel_requirement('Y')
        self.assertEqual(result, 'Yes')

        result = map_hostel_requirement('N')
        self.assertEqual(result, 'No')

        result = map_hostel_requirement('X')
        self.assertEqual(result, 'X')


if __name__ == '__main__':
    unittest.main()