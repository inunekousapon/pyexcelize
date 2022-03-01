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