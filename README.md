pyexcelize
===========

## Introduction

Creating formatting in Python when outputting to Excel can be tedious. Usually, you will want to write the formatting in Excel and then just output the data.

It is a good idea to use OpenPyXL when working with Excel in Python. However, OpenPyXL uses a very large amount of memory.

OpenPyXL is not suitable for writing data to a loaded Excel. However, there is a library in the Go language called excelize that is suitable for writing data.

pyexcelize wraps Go's excelize and makes it usable.


## Basic Usage

```python

## Create New File
index = pe.new_file()

## Save as File
pe.save_as(index, '__tmp/test.xlsx')

## Close Workbook
pe.close(index)

## Open Exist Excel file
index = pe.open_file('./__tmp/test.xlsx')

## Create New Sheet
new_sheet = pe.new_sheet(index, 'Sheet2')

## Copy New Sheet
to_sheet = pe.copy_sheet(index, 1, new_sheet)

## Set Active Sheet
pe.set_active_sheet(index, new_sheet)

## Delete Sheet
pe.delete_sheet(index, 'Sheet1')

## Set Cell Value
pe.set_cell_int(index, "Sheet2", "A1", 1)
pe.set_cell_int(index, "Sheet2", "A2", 2)
pe.set_cell_int(index, "Sheet2", "A3", 3)
pe.set_cell_str(index, "Sheet2", "A4", "hello")

## Get Cell Value
pe.get_cell_value(index, "Sheet2", "A1")
pe.get_cell_value(index, "Sheet2", "A2")
pe.get_cell_value(index, "Sheet2", "A3")
pe.get_cell_value(index, "Sheet2", "A4")

## Get Cell Style
style_index = pe.get_cell_style(index, "Sheet2", "A1")

## Copy Cell Style
pe.set_cell_style(index, "Sheet2", "A2", "A2", style_index)

## Save Update
pe.save(index)

pe.close(index)
```

Stream Writer with Template xlsx.  
If you don't use Stream Writer, it will use more than 2GB of memory.  
Stream Writer will only consume less than 100MB of memory.

```python
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
pe.save_as(index, './__tmp/output.xlsx')
pe.close(index)
```

## Memo

Go Build Command
```
go build -buildmode=c-shared -o py-excelize.so main.go
```

Test
```
python -m unittest
```