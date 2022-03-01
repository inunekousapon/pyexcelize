from datetime import datetime, date
from ctypes import CDLL, c_int, c_char_p, create_string_buffer, Structure, c_float, sizeof


lib = CDLL('py-excelize.so')
ENCODE = 'utf-8'


__version__ = '0.1.0'


class PyExcelizeError(Exception):
    pass


class ExcelValue(Structure):
    _fields_ = [
        ('int_value', c_int),
        ('float_value', c_float),
        ('str_value', c_char_p),
        ('value_type', c_int),
    ]

ValueType_Int = 0
ValueType_Float = 1
ValueType_String = 2
ValueType_Bool = 3
ValueType_Time = 4
ValueType_Nil = 5


def _param_to_excel_value(value, param):
    if isinstance(value, int):
        param.int_value = value
        param.value_type = ValueType_Int
    elif isinstance(value, str):
        param.str_value = value.encode(ENCODE)
        param.value_type = ValueType_String
    elif isinstance(value, float):
        param.float_value = value
        param.value_type = ValueType_Float
    elif isinstance(value, bool):
        param.int_value = 1 if value else 0
        param.value_type = ValueType_Bool
    elif isinstance(value, (datetime, date)):
        param.str_value = value.isoformat().encode(ENCODE)
        param.value_type = ValueType_Time
    elif value is None:
        param.value_type = ValueType_Nil
    else:
        import pdb;pdb.set_trace()
        raise PyExcelizeError('unsupported type')


def new_file() -> int:
    return lib.NewFile()


def open_file(filename: str) -> int:
    return lib.OpenFile(filename.encode(ENCODE))


def save(file_index: int) -> None:
    lib.Save(file_index)


def save_as(file_index: int, filename: str) -> None:
    lib.SaveAs(file_index, filename.encode(ENCODE))


def close(file: int) -> None:
    lib.Close(file)


def new_sheet(file_index: int, sheet_name: str) -> None:
    lib.NewSheet(file_index, sheet_name.encode(ENCODE))


def delete_sheet(file_index: int, sheet_name: str) -> None:
    lib.DeleteSheet(file_index, sheet_name.encode(ENCODE))


def copy_sheet(file_index: int, fromIndex: int, toIndex: int) -> int:
    result = lib.CopySheet(file_index, fromIndex, toIndex)
    if result < 0:
        raise PyExcelizeError('copy sheet error')
    return result


def set_active_sheet(file_index: int, index: int) -> None:
    lib.SetActiveSheet(file_index, index)


def get_active_sheet_index(file_index: int) -> int:
    lib.GetActiveSheetIndex.restype = c_int
    return lib.GetActiveSheetIndex(file_index)


def set_sheet_row(file_index:int, sheet_name: str, axis: str, values):
    ExcelValues = ExcelValue * len(values)
    param = ExcelValues()
    for v, p in zip(values, param):
        _param_to_excel_value(v, p)
    lib.SetSheetRow(file_index, sheet_name.encode(ENCODE), axis.encode(ENCODE), param, len(values))


def set_cell_int(file_index:int, sheet_name: str, axis: str, value: int) -> None:
    lib.SetCellInt(file_index, sheet_name.encode(ENCODE), axis.encode(ENCODE), value)


def set_cell_str(file_index: int, sheet_name: str, axis: str, value: str) -> None:
    lib.SetCellStr(file_index, sheet_name.encode(ENCODE), axis.encode(ENCODE), value.encode(ENCODE))


def set_cell_style(file_index: int, sheet_name: str, h_cell: str, v_cell:str, style: int) -> None:
    lib.SetCellStyle(file_index, sheet_name.encode(ENCODE), h_cell.encode(ENCODE), v_cell.encode(ENCODE), style)


def get_cell_style(file_index: int, sheet_name: str, axis: str) -> int:
    lib.GetCellStyle.restype = c_int
    return lib.GetCellStyle(file_index, sheet_name.encode(ENCODE), axis.encode(ENCODE))


def get_cell_value(file_index: int, sheet_name: str, axis: str) -> str:
    buf_size = 2048
    buf = create_string_buffer(buf_size)
    lib.GetCellValue.restype = c_char_p
    return lib.GetCellValue(file_index, sheet_name.encode(ENCODE), axis.encode(ENCODE), buf, buf_size).decode(ENCODE)


def new_stream_writer(file_index: int, sheet_name: str) -> int:
    return lib.NewStreamWriter(file_index, sheet_name.encode(ENCODE))


def set_row(writer_index: int, axis: str, row: list) -> None:
    ExcelValues = ExcelValue * len(row)
    param = ExcelValues()
    for v, p in zip(row, param):
        _param_to_excel_value(v, p)
    lib.SetRow(writer_index, axis.encode(ENCODE), param, len(row))


def flush(writer_index: int) -> None:
    lib.Flush(writer_index)