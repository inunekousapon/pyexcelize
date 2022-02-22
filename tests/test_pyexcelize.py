import unittest
import os
import shutil

import pyexcelize as pe


class ASeriesOfTest(unittest.TestCase):

    def setUp(self):
        try:
            os.mkdir('__tmp')
        except FileExistsError as e:
            pass

    def test_a_series_of_test(self):
        index = pe.new_file()
        self.assertEqual(1, index)
        pe.save_as(index, '__tmp/test.xlsx')
        self.assertTrue(os.path.exists('__tmp/test.xlsx'))
        pe.close(index)

        index = pe.open_file('./__tmp/test.xlsx')
        self.assertEqual(2, index)
        new_sheet = pe.new_sheet(index, 'Sheet2')
        to_sheet = pe.copy_sheet(index, 1, new_sheet)
        pe.set_active_sheet(index, new_sheet)
        self.assertEqual(to_sheet, pe.get_active_sheet_index(index))
        pe.delete_sheet(index, 'Sheet1')

        pe.set_cell_int(index, "Sheet2", "A1", 1)
        pe.set_cell_int(index, "Sheet2", "A2", 2)
        pe.set_cell_int(index, "Sheet2", "A3", 3)
        pe.set_cell_str(index, "Sheet2", "A4", "hello")

        self.assertEqual('1', pe.get_cell_value(index, "Sheet2", "A1"))
        self.assertEqual('2', pe.get_cell_value(index, "Sheet2", "A2"))
        self.assertEqual('3', pe.get_cell_value(index, "Sheet2", "A3"))
        self.assertEqual('hello', pe.get_cell_value(index, "Sheet2", "A4"))

        style_index = pe.get_cell_style(index, "Sheet2", "A1")
        pe.set_cell_style(index, "Sheet2", "A2", "A2", style_index)
        pe.save(index)


    def tearDown(self):
        shutil.rmtree('__tmp')


if __name__ == '__main__':
    unittest.main()