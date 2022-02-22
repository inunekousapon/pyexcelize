/* Code generated by cmd/cgo; DO NOT EDIT. */

/* package command-line-arguments */


#line 1 "cgo-builtin-export-prolog"

#include <stddef.h> /* for ptrdiff_t below */

#ifndef GO_CGO_EXPORT_PROLOGUE_H
#define GO_CGO_EXPORT_PROLOGUE_H

#ifndef GO_CGO_GOSTRING_TYPEDEF
typedef struct { const char *p; ptrdiff_t n; } _GoString_;
#endif

#endif

/* Start of preamble from import "C" comments.  */


#line 3 "main.go"
#include <stdlib.h>

#line 1 "cgo-generated-wrapper"


/* End of preamble from import "C" comments.  */


/* Start of boilerplate cgo prologue.  */
#line 1 "cgo-gcc-export-header-prolog"

#ifndef GO_CGO_PROLOGUE_H
#define GO_CGO_PROLOGUE_H

typedef signed char GoInt8;
typedef unsigned char GoUint8;
typedef short GoInt16;
typedef unsigned short GoUint16;
typedef int GoInt32;
typedef unsigned int GoUint32;
typedef long long GoInt64;
typedef unsigned long long GoUint64;
typedef GoInt64 GoInt;
typedef GoUint64 GoUint;
typedef __SIZE_TYPE__ GoUintptr;
typedef float GoFloat32;
typedef double GoFloat64;
typedef float _Complex GoComplex64;
typedef double _Complex GoComplex128;

/*
  static assertion to make sure the file is being used on architecture
  at least with matching size of GoInt.
*/
typedef char _check_for_64_bit_pointer_matching_GoInt[sizeof(void*)==64/8 ? 1:-1];

#ifndef GO_CGO_GOSTRING_TYPEDEF
typedef _GoString_ GoString;
#endif
typedef void *GoMap;
typedef void *GoChan;
typedef struct { void *t; void *v; } GoInterface;
typedef struct { void *data; GoInt len; GoInt cap; } GoSlice;

#endif

/* End of boilerplate cgo prologue.  */

#ifdef __cplusplus
extern "C" {
#endif

extern GoInt NewFile();
extern GoInt OpenFile(char* filename);
extern void Save(GoInt fIndex);
extern void SaveAs(GoInt fIndex, char* filename);
extern void Close(GoInt fIndex);
extern GoInt NewSheet(GoInt fIndex, char* sheetname);
extern void DeleteSheet(GoInt fIndex, char* sheetname);
extern GoInt CopySheet(GoInt fIndex, GoInt fromSheet, GoInt toSheet);
extern void SetActiveSheet(GoInt fIndex, GoInt sheetIndex);
extern GoInt GetActiveSheetIndex(GoInt fIndex);
extern void SetCellInt(GoInt fIndex, char* sheetname, char* axis, GoInt value);
extern void SetCellStr(GoInt fIndex, char* sheetname, char* axis, char* value);
extern void SetCellStyle(GoInt fIndex, char* sheetname, char* hCell, char* vCell, GoInt styleIndex);

//SetCellHyperLink
//SetCellRichText
//GetCellRichText
extern GoUint8* GetCellValue(GoInt fIndex, char* sheetname, char* axis, GoUint8* out, GoInt64 outN);
extern GoInt GetCellStyle(GoInt fIndex, char* sheetname, char* axis);

#ifdef __cplusplus
}
#endif
