from ctypes import CDLL, c_int, c_char_p, create_string_buffer


lib = CDLL('py-excelize.so')
ENCODE = 'utf-8'


__version__ = '0.1.0'


class PyExcelizeError(Exception):
    pass


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


if __name__ == '__main__':
    p = lib.NewFile()
    lib.SaveAs(p, b'./hello.xlsx')
