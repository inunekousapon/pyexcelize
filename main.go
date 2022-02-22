package main

//#include <stdlib.h>
import "C"

import (
	"bytes"
	"unsafe"

	"github.com/xuri/excelize/v2"
)


var files map[int]*excelize.File
var fIndex = 0

/////////////////////////////////////////////////////////
//                      WorkBook
/////////////////////////////////////////////////////////

//export NewFile
func NewFile() int {
	f := excelize.NewFile()
	if files == nil {
		files = make(map[int]*excelize.File)
	}
	fIndex += 1
	files[fIndex] = f
	return fIndex;
}

//export OpenFile
func OpenFile(filename *C.char) int {
	f, err:= excelize.OpenFile(C.GoString(filename))
	if err != nil {
		return -1;
	}
	if files == nil {
		files = make(map[int]*excelize.File)
	}
	fIndex += 1
	files[fIndex] = f
	return fIndex;
}

//export Save
func Save(fIndex int) {
	files[fIndex].Save()
}

//export SaveAs
func SaveAs(fIndex int, filename *C.char) {
	files[fIndex].SaveAs(C.GoString(filename))
}

//export Close
func Close(fIndex int) {
	files[fIndex].Close()
	delete(files, fIndex)
}

//export NewSheet
func NewSheet(fIndex int, sheetname *C.char) int {
	return files[fIndex].NewSheet(C.GoString(sheetname))
}

//export DeleteSheet
func DeleteSheet(fIndex int, sheetname *C.char) {
	files[fIndex].DeleteSheet(C.GoString(sheetname))
}

//export CopySheet
func CopySheet(fIndex int, fromSheet int, toSheet int) int {
	err:=files[fIndex].CopySheet(fromSheet, toSheet)
	if err != nil {
		return -1
	}
	return toSheet
}

//GroupSheets
//UnGroupSheets
//SetSheetBackground

//export SetActiveSheet
func SetActiveSheet(fIndex int, sheetIndex int) {
	files[fIndex].SetActiveSheet(sheetIndex)
}

//export GetActiveSheetIndex
func GetActiveSheetIndex(fIndex int) int {
	return files[fIndex].GetActiveSheetIndex()
}

//SetSheetVisible
//SetSheetFormatPr
//GetSheetFormatPr
//SetSheetViewOptions
//GetSheetViewOptions
//SetPageLayout
//GetPageLayout
//SetPageMargins
//GetPageMargins
//SetWorkbookPrOptions
//GetWorkbookPrOptions
//SetHeaderFooter
//GetDefinedName
//DeleteDefinedName
//SetAppProps
//GetAppProps
//SetDocProps
//GetDocProps


/////////////////////////////////////////////////////////
//                      WorkSheet
/////////////////////////////////////////////////////////

//SetColVisible
//SetColWidth
//SetRowHeight
//SetRowVisible
//GetSheetName
//GetColVisible
//GetColWidth
//GetRowHeight
//GetRowVisible
//GetSheetIndex
//GetSheetMap
//GetSheetList
//SetSheetName
//SetSheetPrOptions
//GetSheetPrOptions
//InsertCol
//InsertRow
//DuplicateRow
//DuplicateRowTo
//SetRowOutlineLevel
//SetColOutlineLevel
//GetRowOutlineLevel
//GetColOutlineLevel
//Cols
//Rows
//Next
//Error
//Rows
//Columns
//Next
//Error
//Close
//SearchSheet
//ProtectSheet
//UnprotectSheet
//RemoveCol
//RemoveRow
//SetSheetRow
//InsertPageBreak
//RemovePageBreak


/////////////////////////////////////////////////////////
//                        Cell
/////////////////////////////////////////////////////////

//SetCellValue
//SetCellBool
//SetCellDefault

//export SetCellInt
func SetCellInt(fIndex int, sheetname *C.char, axis *C.char, value int) {
	files[fIndex].SetCellInt(C.GoString(sheetname), C.GoString(axis), value)
}

//export SetCellStr
func SetCellStr(fIndex int, sheetname *C.char, axis *C.char, value *C.char) {
	files[fIndex].SetCellStr(C.GoString(sheetname), C.GoString(axis), C.GoString(value))
}

//export SetCellStyle
func SetCellStyle(fIndex int, sheetname *C.char, hCell *C.char, vCell *C.char, styleIndex int) {
	files[fIndex].SetCellStyle(C.GoString(sheetname), C.GoString(hCell), C.GoString(vCell), styleIndex)
}

//SetCellHyperLink
//SetCellRichText
//GetCellRichText
//export GetCellValue
func GetCellValue(fIndex int, sheetname *C.char, axis *C.char, out *byte, outN int64) *byte {
	val, err := files[fIndex].GetCellValue(C.GoString(sheetname), C.GoString(axis))
	if err != nil {
		return nil
	}
	outBytes := unsafe.Slice(out, outN)[:0]
	buf := bytes.NewBuffer(outBytes)
	buf.WriteString(val)
	buf.WriteByte(0) // Null terminator
	return out
}
//GetCellType
//GetCols
//GetRows
//GetCellHyperLink

//export GetCellStyle
func GetCellStyle(fIndex int, sheetname *C.char, axis *C.char) int {
	index, err := files[fIndex].GetCellStyle(C.GoString(sheetname), C.GoString(axis))
	if err != nil {
		return -1
	}
	return index;
}

//MergeCell
//UnmergeCell
//GetMergeCells
//GetCellValue
//GetStartAxis
//GetEndAxis
//AddComment
//GetComments
//SetCellFormula
//GetCellFormula
//CalcCellValue


/////////////////////////////////////////////////////////
//                        Graph
/////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////
//                        Image
/////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////
//                        Shape
/////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////
//                      Sparkline
/////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////
//                        Style
/////////////////////////////////////////////////////////

//NewStyle
//SetColStyle
//SetRowStyle
//SetDefaultFont
//GetDefaultFont


/////////////////////////////////////////////////////////
//                     StreamWriter
/////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////
//　                  Data Validation
/////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////
//　                    Pivot Table
/////////////////////////////////////////////////////////


/////////////////////////////////////////////////////////
//　                       Tools
/////////////////////////////////////////////////////////

func main() {
}
