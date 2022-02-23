pyexcelize
===========

## Introduction

If you want to handle Excel with formatting in Python, you can use OpenPyXL.
However, OpenPyXL is very slow and uses a lot of memory, making it useless in some cases.

OpenPyXL has a lot of features, and no alternative library was found in Python.
On the other hand, Go had a highly functional library called Excelize.

This package indirectly calls Excelize in Go from Python.

<img src="https://user-images.githubusercontent.com/6970513/155269441-0d93b900-cd4d-43c4-be76-0a744786c2c2.png" width=480 alt="Writer Benchmark of Generation 50000 * 20(Time costs)">

<img src="https://user-images.githubusercontent.com/6970513/155269464-df7b8e6d-4463-4a08-92e2-e8492fd0db04.png" width=480 alt="Writer Benchmark of Generation 50000 * 20(Memory Usage)">

Note: I am using StreamWriter for writing.

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

## Benchmark

Environment

- MacBook Pro (13-inch, 2019)
- CPU 1.4 GHz Quad core Intel Core i5
- Memory 16 GB 2133 MHz LPDDR3

### OpenpyXL Code

```python
import resource
from os.path import getsize
from datetime import datetime

from openpyxl import Workbook


def get_maxrss() -> float:
    r = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return r // 1024 // 1024  # bytes on MacOS


print(f"{datetime.now().isoformat()} init: {get_maxrss()}MB")
wb = Workbook()
ws = wb.worksheets[0]
txt = "1234567890" * 10

for row in range(1,50000):
    for col in range(1, 20):
        ws.cell(row=row, column=col, value=txt)

print(f"{datetime.now().isoformat()} writed: {get_maxrss()}MB")
wb.save('./output.xlsx')
print(f"{datetime.now().isoformat()} saved: {get_maxrss()}MB")
wb.close()
print(f"{datetime.now().isoformat()} closed: {get_maxrss()}MB")
```


### Go Only

```go
package main

import (
	"fmt"
	"syscall"
	"time"

	"github.com/xuri/excelize/v2"
)

func GetStackMemory() int64 {
    var rusageInfo syscall.Rusage
    syscall.Getrusage(syscall.RUSAGE_SELF, &rusageInfo)
    return rusageInfo.Maxrss
}

func main() {

    now := time.Now()
    fmt.Printf("%s init: %dMB\n", now.Format(time.RFC3339), GetStackMemory()/1024/1024)

    f := excelize.NewFile()

    txt := "1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890"
    writer, err := f.NewStreamWriter("Sheet1")
    if err != nil {
        fmt.Println(err)
        syscall.Exit(0)
    }

    for i := 1; i < 100000; i++ {
        row := make([]interface{}, 20)
        for colID := 0; colID < 20; colID++ {
            row[colID] = txt
        }
        cell, _ := excelize.CoordinatesToCellName(1, i)
        writer.SetRow(cell, row)
    }
    writer.Flush()

    now = time.Now()
    fmt.Printf("%s writed: %dMB\n", now.Format(time.RFC3339), GetStackMemory()/1024/1024)

    f.SaveAs("output.xlsx")

    now = time.Now()
    fmt.Printf("%s saved: %dMB\n", now.Format(time.RFC3339), GetStackMemory()/1024/1024)

    f.Close()

    now = time.Now()
    fmt.Printf("%s closed: %dMB\n", now.Format(time.RFC3339), GetStackMemory()/1024/1024)
}
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