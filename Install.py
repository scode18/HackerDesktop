import os
import sys
import shutil
import subprocess
import ctypes
from win32com.client import Dispatch

# Путь к утилите busybox.exe
busybox_path = r"Utils\\busybox.exe"

# Список URL для скачивания файлов
file_urls = [
    "https://github.com/scode18/Installers-for-HackerDesktop/releases/download/v1/Installers.7z",
    "https://github.com/scode18/HackerDesktop/raw/main/Archives.7z",
    "https://github.com/scode18/HackerDesktop/raw/main/Files.7z"
]

# Команды для скачивания файлов с помощью busybox.exe wget
download_commands = [[busybox_path, "wget", url] for url in file_urls]

# Выполнение команд для скачивания файлов
for i, command in enumerate(download_commands):
    subprocess.run(command, check=True)
    print(f"Файл {file_urls[i].split('/')[-1]} успешно скачан.")

# Путь к утилите 7za.exe
seven_zip_path = r"Utils\\7za.exe"

# Путь исходных архивов и файлов
source_files = [
    "Installers.7z"
]

# Распаковка всех скачанных файлов
for file in source_files:
    extract_path = "."
    if file.endswith(".7z"):
        extract_path = file.split(".")[0]  # Для архивов создаем папку с именем без расширения
    extract_command = [seven_zip_path, "x", file, f"-o{extract_path}"]
    subprocess.run(extract_command, check=True)
    print(f"Файл {file} успешно распакован в папку {extract_path}.")

print("Все файлы успешно скачаны и распакованы.")

# Список имен файлов для перемещения
files_to_move = [
    "Wallpaper.png",
    "finder.png",
    "Trash Dark.png",
    "Trash Dark Full.png"
]

# Путь назначения
destination_path = r"C:\\"

# Проверка существования файлов и их перемещение
for file_name in files_to_move:
    if os.path.exists(file_name):
        shutil.move(file_name, os.path.join(destination_path, file_name))
        print(f"Файл {file_name} успешно перемещен в {destination_path}")
    else:
        print(f"Файл {file_name} не найден и не был перемещен.")

print("Все операции завершены.")

# # Путь к утилите busybox.exe
# busybox_path = r"Utils\\busybox.exe"

# # URL для скачивания файла Installers.7z
# file_url_installers = "https://github.com/scode18/Installers-for-HackerDesktop/releases/download/v1/Installers.7z"

# # Команда для скачивания файла Installers.7z с помощью busybox.exe wget
# download_command_installers = [busybox_path, "wget", file_url_installers]

# # Выполнение команды скачивания Installers.7z
# subprocess.run(download_command_installers, check=True)

# print("Файл Installers.7z успешно скачан с помощью busybox.exe wget и будет распакован в папку Installers.")

# Путь к утилите 7za.exe
seven_zip_path = r"Utils\\7za.exe"

# Путь к скачанному архиву Installers.7z
archive_file_installers = "Installers.7z"

# Путь для распаковки архива Installers.7z
extract_path_installers = "."

# Команда для распаковки архива Installers.7z с помощью 7za.exe
extract_command_installers = [seven_zip_path, "x", archive_file_installers, f"-o{extract_path_installers}"]

# Выполнение команды распаковки Installers.7z
subprocess.run(extract_command_installers, check=True)

print("Файл Installers.7z успешно распакован.")

# Путь к скачанному архиву Files.7z
archive_file_installers = "Files.7z"

# Путь для распаковки архива Files.7z
extract_path_installers = "."

# Команда для распаковки архива Files.7z с помощью 7za.exe
extract_command_installers = [seven_zip_path, "x", archive_file_installers, f"-o{extract_path_installers}"]

# Выполнение команды распаковки Files.7z
subprocess.run(extract_command_installers, check=True)

print("Файл Files.7z успешно распакован.")

# # URL для скачивания файла Archives.7z
# file_url_archives = "https://github.com/scode18/HackerDesktop/raw/main/Archives.7z"

# # Команда для скачивания файла Archives.7z с помощью busybox.exe wget
# download_command_archives = [busybox_path, "wget", file_url_archives]

# # Выполнение команды скачивания Archives.7z
# subprocess.run(download_command_archives, check=True)

# print("Файл Archives.7z успешно скачан с помощью busybox.exe wget и будет распакован в текущую директорию.")

# Путь к скачанному архиву Archives.7z
archive_file_archives = "Archives.7z"

# Путь для распаковки архива Archives.7z (в текущую директорию)
extract_path_archives = "."

# Команда для распаковки архива Archives.7z с помощью 7za.exe
extract_command_archives = [seven_zip_path, "x", archive_file_archives, f"-o{extract_path_archives}"]

# Выполнение команды распаковки Archives.7z
subprocess.run(extract_command_archives, check=True)

print("Файл Archives.7z успешно распакован в текущую директорию.")

# Пути для установки программ
nexus_winstep_path = os.path.join(os.getcwd(), "Installers\\NexusSetup.exe")
rainmeter_path = os.path.join(os.getcwd(), "Installers\\Rainmeter-4.5.18.exe")

# Пути к файлам и папкам
# wallpaper_file = "Wallpaper.png"
# wallpaper_dest = os.path.join("C:\\", "Wallpaper.png")

winstep_archive = "Archives\\Winstep.7z"
winstep_dest = os.path.join("C:\\", "Users", "Public", "Documents", "Winstep")

rainmeter_archive = "Archives\\Rainmeter.7z"
rainmeter_user_dir = os.path.expanduser("~")
rainmeter_dest = os.path.join(rainmeter_user_dir, "Documents", "Rainmeter")

# Путь к 7za.exe
sevenzip_path = os.path.join(os.getcwd(), "Utils\\7za.exe")

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
        print("Ошибка: 7za.exe не найден. Убедитесь, что он находится в каталоге Utils.")
        exit(1)

    # Перемещаем Wallpaper.png в корень диска C
    # move_file(wallpaper_file, wallpaper_dest)

    # Распаковываем Winstep.7z в C:\Users\Public\Documents\Winstep
    extract_archive(winstep_archive, winstep_dest)

    # Распаковываем Rainmeter.7z в папку Documents текущего пользователя
    extract_archive(rainmeter_archive, rainmeter_dest)

    os.system("start catppuccin_1.4.1.rmskin")

    # Путь к изображению, которое нужно установить в качестве обоев
    image_path = r"C:\\Wallpaper.png"

    # Вызываем функцию для изменения обоев
    change_wallpaper(image_path)

    # Создает ярлыки на рабочий стол
    create_desktop_shortcuts()

    print("Все действия выполнены успешно.")
