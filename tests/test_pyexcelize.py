import unittest
import os
import shutil
import resource
from os.path import getsize
from datetime import datetime
import random

from faker.factory import Factory

import pyexcelize as pe


Faker = Factory.create
fake = Faker()
fake.seed(0)
fake = Faker("ja_JP")


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

    def setUp(self):
        try:
            os.mkdir('__tmp')
        except FileExistsError as e:
            pass

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

        pe.save_as(index, './__tmp/output.xlsx')

        print(f"{datetime.now().isoformat()} saved: {get_maxrss()}MB")

        pe.close(index)

        print(f"{datetime.now().isoformat()} closed: {get_maxrss()}MB")

    def test_template_performance(self):

        print(f"{datetime.now().isoformat()} init: {get_maxrss()}MB")

        index = pe.open_file('./tests/template.xlsx')
        writer_index = pe.new_stream_writer(index, "Sheet1")
        headers = [
            "employee name",
            "company",
            "salary",
        ]
        pe.set_row(writer_index, "A1", headers)
        for row in range(2,500000):
            params = [
                fake.name(),
                random.choice(["Google", "Microsoft", "Apple", "Toyota", "Meta"]),
                random.randint(10000, 10000000),
            ]
            pe.set_row(writer_index, f"A{row}", params)
        pe.add_table(writer_index, "A1", "C499999", dict(
            table_name="テーブル1",
            table_style="TableStyleMedium2",
            show_first_column=True,
            show_last_column=True,
            show_row_stripes=True,
            show_column_stripes=False,
        ))
        pe.flush(writer_index)
        print(f"{datetime.now().isoformat()} writed: {get_maxrss()}MB")
        pe.save_as(index, './__tmp/output.xlsx')
        print(f"{datetime.now().isoformat()} saved: {get_maxrss()}MB")
        pe.close(index)
        print(f"{datetime.now().isoformat()} closed: {get_maxrss()}MB")

    def tearDown(self):
        shutil.rmtree('__tmp')

if __name__ == '__main__':
    unittest.main()