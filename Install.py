import os
import sys
import shutil
import subprocess
import ctypes
from win32com.client import Dispatch

# Пути для установки программ
nexus_winstep_path = os.path.join(os.getcwd(), "Installers\\NexusSetup.exe")
rainmeter_path = os.path.join(os.getcwd(), "Installers\\Rainmeter-4.5.18.exe")

# Пути к файлам и папкам
wallpaper_file = "Wallpaper.png"
wallpaper_dest = os.path.join("C:\\", "Wallpaper.png")

winstep_archive = "Archives\\Winstep.7z"
winstep_dest = os.path.join("C:\\", "Users", "Public", "Documents", "Winstep")

rainmeter_archive = "Archives\\Rainmeter.7z"
rainmeter_user_dir = os.path.expanduser("~")
rainmeter_dest = os.path.join(rainmeter_user_dir, "Documents", "Rainmeter")

# Путь к 7za.exe
sevenzip_path = os.path.join(os.getcwd(), "7za.exe")

# Функция для тихой установки программ
def install_program(file_path, silent_args):
    try:
        subprocess.run([file_path] + silent_args, check=True)
        print(f"Программа {file_path} успешно установлена.")
    except subprocess.CalledProcessError as e:
        print(f"Ошибка при установке программы {file_path}: {e}")

def move_file(src, dst):
    try:
        shutil.move(src, dst)
        print(f"Файл {src} успешно перемещен в {dst}.")
    except (shutil.Error, IOError) as e:
        print(f"Ошибка при перемещении файла {src}: {e}")

def extract_archive(archive_path, dest_dir):
    try:
        subprocess.run([sevenzip_path, "x", archive_path, "-o" + dest_dir, "-y"], check=True)
        print(f"Архив {archive_path} успешно распакован в {dest_dir}.")
    except subprocess.CalledProcessError as e:
        print(f"Ошибка при распаковке архива {archive_path}: {e}")

def change_wallpaper(image_path):
    """
    Функция, которая меняет обои рабочего стола в Windows.
    """
    try:
        ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 0)
        print(f"Обои рабочего стола успешно изменены на {image_path}")
    except Exception as e:
        print(f"Ошибка при изменении обоев рабочего стола: {e}")

def create_desktop_shortcuts():
    # Create a shortcut for "Start Hacker Desktop"
    shell = Dispatch("WScript.Shell")
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "Start Hacker Desktop.lnk")
    shortcut = shell.CreateShortcut(desktop_path)
    shortcut.TargetPath = os.path.join(os.getcwd(), "Start Hacker Desktop.bat")
    shortcut.IconLocation = r"C:\\Program Files\\Rainmeter\\Rainmeter.exe"
    shortcut.WorkingDirectory = os.getcwd()
    shortcut.Save()

    # Create a shortcut for "Exit Hacker Desktop"
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop", "Exit Hacker Desktop.lnk")
    shortcut = shell.CreateShortcut(desktop_path)
    shortcut.TargetPath = os.path.join(os.getcwd(), "Exit Hacker Desktop.bat")
    shortcut.IconLocation = r"C:\\Program Files\\Rainmeter\\Rainmeter.exe"
    shortcut.WorkingDirectory = os.getcwd()
    shortcut.Save()

if __name__ == "__main__":
    # Устанавливаем Nexus Winstep в тихом режиме
    install_program(nexus_winstep_path, ["/SILENT", "/VERYSILENT"])

    # Устанавливаем Rainmeter в тихом режиме
    install_program(rainmeter_path, ["/S"])

    print("Установка завершена.")

    # Убедитесь, что 7za.exe находится в том же каталоге, где находится Python-файл
    if not os.path.exists(sevenzip_path):
        print("Ошибка: 7za.exe не найден. Убедитесь, что он находится в том же каталоге, где находится Python-файл.")
        exit(1)

    # Перемещаем Wallpaper.png в корень диска C
    move_file(wallpaper_file, wallpaper_dest)

    # Распаковываем Winstep.7z в C:\Users\Public\Documents\Winstep
    extract_archive(winstep_archive, winstep_dest)

    # Распаковываем Rainmeter.7z в папку Documents текущего пользователя
    extract_archive(rainmeter_archive, rainmeter_dest)

    # Путь к изображению, которое нужно установить в качестве обоев
    image_path = r"C:\\Wallpaper.png"

    # Вызываем функцию для изменения обоев
    change_wallpaper(image_path)

    # Создает ярлыки на рабочий стол
    create_desktop_shortcuts()

    print("Все действия выполнены успешно.")
