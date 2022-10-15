from typing import Iterator
from excelsigncheck.signature import ExcelSignature

class ExcelSignCheck:
    import win32com.client
    import pythoncom
    __file_path: str
    __dispatch = None
    __workbook = None

    def __init__(self, file_path: str):
        self.__file_path = file_path
        self.pythoncom.CoInitialize()

    def get_dispatch(self):
        if self.__dispatch is None:
            self.__dispatch = self.win32com.client.DispatchEx('Excel.Application')
        return self.__dispatch

    def open(self):
        if self.__workbook is None:
            self.__workbook = self.get_dispatch().Workbooks.Open(self.__file_path)

    @property
    def signatures(self) -> Iterator[ExcelSignature]:
        return map(lambda sign: ExcelSignature(sign), self.__workbook.Signatures)

    def __enter__(self):
        self.open()
        return self

    def close(self):
        try:
            self.__workbook.Close()
        except AttributeError:
            # AttributeError: Open.Close
            pass
        try:
            self.get_dispatch().Quit()
        except AttributeError:
            # AttributeError: Excel.Application.Quit
            pass
        self.pythoncom.CoUninitialize()

    def __exit__(self, type, value, traceback):
        self.close()
