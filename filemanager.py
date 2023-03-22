from os.path import isfile, join, exists
import os
from PIL import Image
import docx
import PySimpleGUI as sg
from pathlib import Path


class FileManager:
    def __init__(self):
        self.settings_path = './settings.ini'
        self.tesseract_path = 'c:/program files/tesseract/tesseract.exe'
        self.input_dir = str(Path(__file__).parent.absolute())
        self.output_dir = str(Path(__file__).parent.absolute())
        if isfile(self.settings_path):
            self.read_settings_file()
        else:
            self.create_settings_file()

    def create_settings_file(self):
        with open(self.settings_path, 'w') as settings_file:
            settings_file.write(f'{self.tesseract_path}\n{self.input_dir}\n{self.output_dir}')

    def read_settings_file(self):
        settings_file = open(self.settings_path)
        lines = settings_file.readlines()
        try:
            tesseract_bin_path = lines[0].strip()
            last_open_dir = lines[1].strip()
            last_output_dir = lines[2].strip()
        except IndexError:
            settings_file.close()
            os.remove(self.settings_path)
        else:
            if exists(tesseract_bin_path):
                self.tesseract_path = tesseract_bin_path
            if exists(last_open_dir):
                self.input_dir = last_open_dir
            if exists(last_output_dir):
                self.output_dir = last_output_dir
            settings_file.close()

    def update_settings_file(self, values):
        tesseract_path = values["-PATH_TESSERACT-"]
        input_dir = values["-INPUT_DIR-"]
        output_dir = values["-OUTPUT_DIR-"]
        with open(self.settings_path, 'w') as settings_file:
            settings_file.write(f'{tesseract_path}\n{input_dir}\n{output_dir}')

    @staticmethod
    def parse_folder(path):
        """
        Parses selected folder for valid images.
        :param path: selected folder path.
        :return: list of valid images.
        """
        valid_pictures = []
        picture_names = [f for f in os.listdir(path)
                         if isfile(join(path, f))
                         and os.path.splitext(f)[1].lower() in ['.jpg', '.gif', '.jpeg', '.jfif']]
        for name in picture_names:
            try:
                im = Image.open(f'{path}/{name}')
                im.verify()
                im.close()
            except IOError:
                pass
            else:
                valid_pictures.append(name)
        return valid_pictures

    @staticmethod
    def save_file(transcribed_text, values, img_name):
        """
        Saves OCR text to file.
        :param transcribed_text: Transcribed text.
        :param values: PySimpleGui window values.
        :param img_name: Image name.
        """
        output_dir = values['-OUTPUT_DIR-']
        file_format = values['-FILE_FORMAT-']
        save_path = f'{output_dir}/{img_name}.{file_format}'

        if isfile(save_path):
            choice = sg.PopupYesNo(f'File {img_name}.{file_format} already exists.\nOverwrite?', modal=True)
            if choice == 'No':
                return

        if file_format == 'txt':
            FileManager.save_as_txt(save_path, transcribed_text)

        elif file_format == 'docx':
            FileManager.save_as_docx(save_path, transcribed_text)


    @staticmethod
    def save_as_txt(save_path, transcribed_text):
        try:
            with open(save_path, 'w') as output_file:
                output_file.write(transcribed_text)
        except FileNotFoundError:
            sg.PopupError('Invalid output directory.')

    @staticmethod
    def save_as_docx(save_path, transcribed_text):
        document = docx.Document()
        for paragraph in transcribed_text.split('\n\n'):
            document.add_paragraph(paragraph)
        try:
            document.save(save_path)
        except FileNotFoundError:
            sg.PopupError('Invalid output directory.')
