py-excelize
===========

## Introduction

If you want to handle Excel with formatting in Python, you can use OpenPyXL.
However, OpenPyXL is very slow and uses a lot of memory, making it useless in some cases.

OpenPyXL has a lot of features, and no alternative library was found in Python.
On the other hand, Go had a highly functional library called Excelize.

This package indirectly calls Excelize in Go from Python.

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

## Memo

Go Build Command
```
go build -buildmode=c-shared -o py-excelize.so main.go
```

Test
```
python -m unittest
```