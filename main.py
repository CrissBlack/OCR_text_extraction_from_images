from PIL import Image, ImageTk
import PySimpleGUI as sg
from os.path import join
from filemanager import FileManager
from ocr import OCR
from pathlib import Path


filemanager = FileManager()
ocr_engine = OCR(filemanager.tesseract_path)


def populate_file_list(path_to_folder, window):
    """
    Populates the listbox with valid images.
    :param path_to_folder: Currently selected folder.
    :param window: PySimpleGui Window element.
    """
    filenames = filemanager.parse_folder(path_to_folder)

    window['-DISPLAYED_PATH-'].update(path_to_folder
                                      if len(path_to_folder) < 50
                                      else f'.../{path_to_folder[-45:]}')
    window['-FILE_LIST-'].update(filenames)


def load_image(window, path):
    """
    Displays the selected image.
    :param window: PySimpleGui Window element.
    :param path: Selected image path.
    :return:
    """
    image = Image.open(path)
    image.thumbnail(size=(700, 600))
    photo_img = ImageTk.PhotoImage(image=image)
    window['-IMAGE-'].update(data=photo_img)


def main():
    """
    Main PySimpleGui Window.
    """
    location = 0    # current location in selected images if multiple selected
    sg.theme('DarkGrey')
    print(filemanager.input_dir)
    column_upper_left = [
        [sg.Input(key='-INPUT_DIR-', enable_events=True, visible=False),
         sg.FolderBrowse('Select source folder', key='-INPUT_DIR-', initial_folder=filemanager.input_dir)
         ],

        [sg.Text(filemanager.input_dir, key='-DISPLAYED_PATH-')],

        [sg.Listbox([], enable_events=True, key="-FILE_LIST-", size=(40, 30), select_mode='extended')],
        [sg.Button('OCR Selected', key='-BTN_OCR-'), sg.Push(), sg.Button('Preview', key='-PREVIEW-', visible=False)]
    ]

    column_upper_right = [
        [sg.Image(size=(700, 600), key='-IMAGE-')],
    ]

    column_lower_left = [
        [sg.Text('Select output folder'),
         sg.InputText(filemanager.output_dir, size=(35, 1), key='-OUTPUT_DIR-', enable_events=True),
         sg.FolderBrowse('Browse', initial_folder=filemanager.output_dir)],

        [sg.Text(text='Save files as          '),
         sg.Combo(['txt', 'docx'], default_value='txt', size=(33, 1), readonly=True, key='-FILE_FORMAT-')]
    ]

    column_lower_right = [
        [sg.VPush(), sg.Button('Prev', key='-BTN_PREV-', visible=False),
         sg.Button('Next', key='-BTN_NEXT-', visible=False)],

        [sg.Text('Pytesseract binary location'),
         sg.InputText(filemanager.tesseract_path, size=(35, 1), key='-PATH_TESSERACT-', enable_events=True),
         sg.FileBrowse('Browse', initial_folder=filemanager.tesseract_path)]
    ]

    layout = [
        [sg.Column(layout=column_upper_left), sg.VSeparator(), sg.Push(), sg.Column(layout=column_upper_right),
         sg.Push()],

        [sg.HorizontalSeparator()],

        [sg.Column(layout=column_lower_left), sg.VSeparator(), sg.Push(),
         sg.Column(layout=column_lower_right, vertical_alignment='r')]

    ]

    window = sg.Window('OCR', layout=layout, size=(1100, 700), resizable=True, grab_anywhere=True, finalize=True)
    populate_file_list(filemanager.input_dir, window)
    window['-INPUT_DIR-'].update(filemanager.input_dir)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break

        # Open directory button
        if event == '-INPUT_DIR-':
            path_to_folder = values['-INPUT_DIR-']
            populate_file_list(path_to_folder, window)

        # Displays image list for the current folder.
        if event == '-FILE_LIST-':
            selected_paths = [join(values['-INPUT_DIR-'], x) for x in values['-FILE_LIST-']]

            window['-BTN_PREV-'].update(visible=True if len(selected_paths)>1 else False)
            window['-BTN_NEXT-'].update(visible=True if len(selected_paths)>1 else False)
            window['-PREVIEW-'].update(visible=True if len(selected_paths) == 1 else False)

            if selected_paths:
                load_image(window, selected_paths[0])

        # If multiple images selected, cycle through them.
        if event == '-BTN_PREV-':
            if location == 0:
                location = len(selected_paths) - 1
            else:
                location -= 1
            load_image(window, selected_paths[location])

        # If multiple images selected, cycle through them.
        if event == '-BTN_NEXT-':
            if location == len(selected_paths) - 1:
                location = 0
            else:
                location += 1
            load_image(window, selected_paths[location])

        # Opens file browser for selecting tesseract binary location.
        if event == '-PATH_TESSERACT-':
            ocr_engine.tesseract.tesseract_cmd = values['-PATH_TESSERACT-']

        # Popup window, displays OCR text for the selected image.
        if event == '-PREVIEW-':
            directory = values['-INPUT_DIR-']
            img_name = values['-FILE_LIST-'][0]
            img_path = join(directory, img_name)

            transcribed_text = ocr_engine.process_image(img_path)
            if transcribed_text:

                preview_window(img_path)

        # OCR selected images and save to output folder.
        if event == '-BTN_OCR-':
            selected_paths = [f'{values["-INPUT_DIR-"]}/{x}' for x in values['-FILE_LIST-']]

            if not selected_paths:
                sg.PopupError('Nothing selected', grab_anywhere=True)
                continue

            for img_path in selected_paths:
                transcribed_text = ocr_engine.process_image(img_path)
                img_name = Path(img_path).stem

                if transcribed_text:
                    filemanager.save_file(transcribed_text, values, img_name)

            filemanager.update_settings_file(values)


def preview_window(img_path):
    """
    Popup window, displays OCR text for the selected image.
    :param img_path: Selected image path.
    """
    img_name = Path(img_path).name
    layout = [
        [sg.Text(text=f'Preview for: {img_name}')],
        [sg.Multiline(default_text=ocr_engine.process_image(img_path), expand_y=True, expand_x=True)],
        [sg.Push(), sg.Button(button_text='Close', key='-CLOSE_PREVIEW-', size=(11, 2)), sg.Push()]
    ]
    window = sg.Window('Preview', size=(800, 700), layout=layout, resizable=True, modal=True, grab_anywhere=True)
    while True:
        event, values = window.read()
        if event == 'Exit' or event == sg.WINDOW_CLOSED or event == '-CLOSE_PREVIEW-':
            break
    window.close()


if __name__ == '__main__':
    main()
