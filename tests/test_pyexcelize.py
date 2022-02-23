import unittest
import os
import shutil
import resource
from os.path import getsize
from datetime import datetime

import pyexcelize as pe


def get_maxrss() -> float:
    r = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return r // 1024 // 1024  # bytes on MacOS


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


class PerformanceTest(unittest.TestCase):

    def test_performance(self):
        print(f"{datetime.now().isoformat()} init: {get_maxrss()}MB")
        index = pe.new_file()

        txt = "1234567890" * 10

        writer_index = pe.new_stream_writer(index, "Sheet1")
        for row in range(1,50000):
            params = []
            for col in "ABCDEFGHIJKLMNOPQRST":
                params.append(txt)
            pe.set_row(writer_index, f"A{row}", params)

        pe.flush(writer_index)

        print(f"{datetime.now().isoformat()} writed: {get_maxrss()}MB")

        pe.save_as(index, 'output.xlsx')

        print(f"{datetime.now().isoformat()} saved: {get_maxrss()}MB")

        pe.close(index)

        print(f"{datetime.now().isoformat()} closed: {get_maxrss()}MB")


if __name__ == '__main__':
    unittest.main()