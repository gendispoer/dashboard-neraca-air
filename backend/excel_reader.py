from data_processor import parse_wb_workbook

class ExcelReader:
    def validate_excel(self, content: bytes):
        if not content:
            raise ValueError("File kosong")
        parse_wb_workbook(content)

    def read_input(self, content: bytes):
        return parse_wb_workbook(content)
