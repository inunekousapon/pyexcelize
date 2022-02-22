py-excelize
===========

## Introduction

If you want to handle Excel with formatting in Python, you can use OpenPyXL.
However, OpenPyXL is very slow and uses a lot of memory, making it useless in some cases.

OpenPyXL has a lot of features, and no alternative library was found in Python.
On the other hand, Go had a highly functional library called Excelize.

This package indirectly calls Excelize in Go from Python.

## Basic Usage

T.B.D

## Memo

Go Build Command
```
go build -buildmode=c-shared -o py-excelize.so main.go
```

Test
```
python -m unittest
```