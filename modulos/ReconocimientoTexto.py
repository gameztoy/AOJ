from PIL import Image
from pytesseract import pytesseract

class ExtraccionTexto():
    def __init__(self,app):
        self.app=app

    def extraerTexto(self, imagen):
        # Abrimos la imagen
        img = Image.open(imagen)

        # Localizamos el exe de donde se encuentra instalado Tesseract
        pytesseract.tesseract_cmd = "C:\\Program Files\\Tesseract-OCR\\tesseract.exe"
        
        # Obtenemos el texto y lo retornamos
        text = pytesseract.image_to_string(img)
        return text[:-1]
