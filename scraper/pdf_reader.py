import PyPDF2


class PDFReader:

    pdf_reader = None

    def __init__(self, pdf_file_path: str):
        pdf_file = open(pdf_file_path, 'rb')
        self.pdf_reader = PyPDF2.PdfFileReader(pdf_file)


    def get_page(self, page: int):
        return self.pdf_reader.getPage(page)

    def get_page_contents(self, page: int):
        return self.get_page(page).extractText()

