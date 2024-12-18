file_path = r"D:\ngu\test\Радіостанції_2024-11-23.xlsx"

import win32clipboard
import os
import struct


def copy_file_to_clipboard(file_path):
    """
    Копіює файл до буфера обміну.

    Args:
      file_path: Шлях до файлу.
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Файл не знайдено: {file_path}")

    # Створюємо структуру DROPFILES
    #   dwSize: Розмір структури в байтах.
    #   pFiles: Зміщення до списку файлів відносно початку структури (в даному випадку 0).
    #   fNC: Прапор, який вказує, чи потрібно копіювати чи переміщувати файли (0 для копіювання).
    #   fWide: Прапор, який вказує, чи використовуються імена файлів в Юнікоді (1 для так).
    #   шлях до файлу з нульовим символом в кінці (в Юнікоді).
    #   нульовий символ в кінці списку файлів (в Юнікоді).
    # Використовуємо 'H' як специфікатор формату для кожного символу в шляху до файлу
    data = struct.pack('IIIHHI' + str(len(file_path)) + 'H' + 'H',
                       20 + len(file_path) * 2 + 2, 0, 0, 1, 0,
                       *[ord(c) for c in file_path], 0, 0)

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_HDROP, data)
    win32clipboard.CloseClipboard()


def is_file_in_clipboard():
    """
    Перевіряє, чи є файл в буфері обміну.

    Returns:
      True, якщо в буфері обміну є файл, False - інакше.
    """

    try:
        win32clipboard.OpenClipboard()
        result = win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_HDROP)
        win32clipboard.CloseClipboard()
        print(result, "RESULT")
        print(win32clipboard.GetClipboardData(win32clipboard.CF_HDROP))
        return result
    except Exception as e:
        print(f"Помилка: {e}")
        return False


def is_clip():
    win32clipboard.OpenClipboard()
    if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_HDROP):
        # print(win32clipboard.GetClipboardData(win32clipboard.CF_HDROP))
        return win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
    else:
        return "Нема даних в буфері!"


def get_clipboard_files():
    """
    Отримує список файлів з буфера обміну.

    Returns:
      Список шляхів до файлів або None, якщо в буфері обміну немає файлів.
    """
    try:
        win32clipboard.OpenClipboard()
        if is_file_in_clipboard():
            data = win32clipboard.GetClipboardData(win32clipboard.CF_HDROP)
            # Розібрати структуру DROPFILES
            files = []
            offset = struct.unpack_from('I', data, 16)[0]
            while offset < len(data) - 2:
                # Отримуємо шлях до файлу в кодуванні UTF-16
                path = data[offset:data.find(b'\x00\x00', offset)].decode('utf-16le')
                files.append(os.path.normpath(path))
                offset += 2 * (len(path) + 1)
            win32clipboard.CloseClipboard()
            return files
        else:
            return None
    except Exception as e:
        print(f"Помилка: {e}")
        return None


# Приклад використання
# copy_file_to_clipboard(file_path)

print(is_clip())
