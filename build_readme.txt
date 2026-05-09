# ============================================================
# СБОРКА ИСПОЛНЯЕМОГО ФАЙЛА
# ============================================================
# Требования: pip install pyinstaller openpyxl python-pptx pillow
# (pandas и numpy НЕ нужны — убраны из зависимостей,
#  это сильно уменьшает размер и ускоряет запуск exe)
# pillow нужен для определения размеров изображений и поворота
# вытянутых картинок, а также для поддержки WebP.
#
# ВАЖНО: --onedir (папка) запускается намного быстрее, чем --onefile,
# потому что --onefile каждый раз распаковывает архив во временную директорию.
# Рекомендуется --onedir, если нужен один файл — заархивируйте папку dist/.
#
# ЗАГЛУШКА ДЛЯ ОТСУТСТВУЮЩИХ КАРТИНОК:
# Положите файл fallback.png (или любое изображение с таким именем) в папку
# рядом с main.py ДО сборки — он будет автоматически включён в exe.
# При запуске exe заглушка берётся из bundle. Если файла нет — программа
# сгенерирует серый PNG-квадрат автоматически.
# ============================================================

# ── Рекомендуемый вариант: папка (быстрый запуск) ───────────
pyinstaller --name hello_kitty --hiddenimport openpyxl --hiddenimport PIL --hiddenimport PIL.Image --hiddenimport PIL.WebPImagePlugin --collect-submodules pptx --collect-data pptx --exclude-module pandas --exclude-module numpy --exclude-module matplotlib --exclude-module scipy --exclude-module IPython --exclude-module tkinter --version-file file_version_info.txt --icon=icon_file.ico main.py

# ── Один файл .exe (медленнее — ~3-5 сек на распаковку) ─────
pyinstaller --name hello_kitty --hiddenimport openpyxl --hiddenimport PIL --hiddenimport PIL.Image --hiddenimport PIL.WebPImagePlugin --collect-submodules pptx --collect-data pptx --exclude-module pandas --exclude-module numpy --exclude-module matplotlib --exclude-module scipy --exclude-module IPython --exclude-module tkinter -F --onefile --version-file file_version_info.txt --icon=icon_file.ico main.py

# ── Через spec-файл (рекомендуется, если нужна fallback.png в bundle) ──
pyinstaller main.spec
