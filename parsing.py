import os
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import zipfile


def show_directory(path):
    files = os.listdir(path)
    for file in files:
        yield file


def parse_excel(to_replace, replacement, start_path, result_path, message_screen):
    is_replaced = False
    try:
        workbook = load_workbook(start_path)
        sheet = workbook.sheetnames[0]
        worksheet = workbook[sheet]
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value == to_replace:
                    cell.value = replacement
                    is_replaced = True
        if is_replaced:
            workbook.save(result_path)
            if message_screen:
                message_screen.appendHtml('<span style="color: green;">Заменено</span>')
            else:
                print('REPLACED')
    except InvalidFileException:
        if message_screen:
            message_screen.appendHtml('<span style="color: red;">Не могу заменить из-за старого формата Excel</span>')
        else:
            print('OLD FORMAT')
    except zipfile.BadZipFile:
        pass


def run(start_directory, result_directory, text_to_replace, result_text, message_screen=None):
    for file_name in show_directory(start_directory):
        if message_screen:
            message_screen.appendHtml("Чтение файла <b>{}</b>".format(file_name))
        else:
            print("Чтение файла <b>{}</b>".format(file_name))
        path = os.path.join(start_directory, file_name)
        result_path = os.path.join(result_directory, file_name)
        parse_excel(text_to_replace, result_text, path, result_path, message_screen)


if __name__ == '__main__':
    run(
        start_directory='documents',
        result_directory='new_documents',
        text_to_replace='to replace',
        result_text='replaced'
    )
