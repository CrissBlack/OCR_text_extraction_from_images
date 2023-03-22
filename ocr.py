from PIL import Image
from pytesseract import pytesseract
import PySimpleGUI as sg


class OCR:
    def __init__(self, tesseract_bin_path):
        self.tesseract = pytesseract
        self.tesseract.tesseract_cmd = tesseract_bin_path

    def process_image(self, img_path):
        """
        Transcribe the image located at {img_path}
        :param img_path: Image path
        :return: Transcribed text.
        """
        img = self.open_image(img_path)

        try:
            text = self.tesseract.image_to_string(img)
        except self.tesseract.TesseractNotFoundError:
            sg.PopupError('Tesseract.exe not found.')
        except TypeError:
            sg.PopupError(f'Could not process file\n{img_path}')
        else:
            return text

    @staticmethod
    def open_image(img_path):
        try:
            img = Image.open(img_path)
        except FileNotFoundError:
            sg.PopupError(f'File\n{img_path}\nmoved or deleted.')
        else:
            return img
