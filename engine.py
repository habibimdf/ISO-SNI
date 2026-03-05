import fitz
from pdf2docx import Converter

class PDFConverterEngine:
    def __init__(self, tesseract_path=None):
        if tesseract_path:
            import pytesseract
            pytesseract.pytesseract.tesseract_cmd = tesseract_path

    def is_scanned_pdf(self, pdf_path):
        doc = fitz.open(pdf_path)
        for page in doc:
            if page.get_text().strip():
                doc.close()
                return False
        doc.close()
        return True

    def convert(self, pdf_path, docx_path):
        try:
            is_scan = self.is_scanned_pdf(pdf_path)
            cv = Converter(pdf_path)
            if is_scan:
                cv.convert(docx_path, is_scan=True, ocr=1)
            else:
                cv.convert(docx_path, start=0, end=None)
            cv.close()
            return True, docx_path
        except Exception as e:
            return False, str(e)