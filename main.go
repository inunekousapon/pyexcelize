package main

/*#include <stdlib.h>
#include <time.h>
struct ExcelValue {
	int int_value;
	float float_value;
	char* string_value;
	int value_type;
};
*/
import "C"

import (
	"bytes"
	"time"
	"unsafe"

	"github.com/xuri/excelize/v2"
)

const ValueType_Int C.int  = 0
const ValueType_Float C.int  = 1
const ValueType_String C.int  = 2
const ValueType_Bool C.int  = 3
const ValueType_Time C.int  = 4
const ValueType_Nil C.int  = 5

var files map[int]*excelize.File
var writers map[int]*excelize.StreamWriter
var fIndex = 0
var wIndex = 0


func convert_excelvalue(val *C.struct_ExcelValue) interface{} {
	switch val.value_type {
	case ValueType_Int:
		return val.int_value
	case ValueType_Float:
		return val.float_value
	case ValueType_String:
		return C.GoString(val.string_value)
	case ValueType_Bool:
		return val.int_value == 1
	case ValueType_Time:
		return time.Unix(int64(val.int_value), 0)
	case ValueType_Nil:
		return nil
	default:
		return nil
	}
}


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
//export NewStreamWriter
func NewStreamWriter(fIndex int, sheetname *C.char) int {
	writer, err := files[fIndex].NewStreamWriter(C.GoString(sheetname))
	if err != nil {
		return -1
	}
	wIndex ++
	if writers == nil {
		writers = make(map[int]*excelize.StreamWriter)
	}
	writers[wIndex] = writer
	return wIndex
}

//export SetRow
func SetRow(wIndex int, axis *C.char, rowPtr *C.struct_ExcelValue, length int) int {
	values := make([]interface{}, length)
	row := unsafe.Slice(rowPtr, length)
	for i, x := range row {
		values[i] = convert_excelvalue(&x)
	}
	err := writers[wIndex].SetRow(C.GoString(axis), values)
	if err != nil {
		return -1;
	}
	return length
}

//AddTable
//MergeCell
//SetColWidth
//export Flush
func Flush(wIndex int) {
	writers[wIndex].Flush()
	delete(writers, wIndex)
}


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
