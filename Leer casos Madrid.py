from src.excel_writer import Excel_writer
from src.pdf_reader import PDF_Reader


def main():
    last_data = PDF_Reader()
    list_data = last_data.read_file()
    del last_data

    excel = Excel_writer()
    excel.write_data(list_data)
    del excel


if __name__ == "__main__":
    main()
