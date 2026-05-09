# ============================================================
# СБОРКА ИСПОЛНЯЕМОГО ФАЙЛА
# ============================================================
# Требования: pip install pyinstaller openpyxl python-pptx
# (pandas и numpy больше НЕ нужны — убраны из зависимостей,
#  это сильно уменьшает размер и ускоряет запуск exe)
#
# ВАЖНО: --onedir (папка) запускается намного быстрее, чем --onefile,
# потому что --onefile каждый раз распаковывает архив во временную директорию.
# Рекомендуется --onedir, если нужен один файл — заархивируйте папку dist/.
# ============================================================

# ── Рекомендуемый вариант: папка (быстрый запуск) ───────────
pyinstaller --hiddenimport openpyxl --collect-submodules pptx --collect-data pptx --exclude-module pandas --exclude-module numpy --exclude-module matplotlib --exclude-module scipy --exclude-module IPython --exclude-module tkinter --windowed --version-file file_version_info.txt --icon=icon_file.ico main.py

# ── Один файл .exe (медленнее — ~3-5 сек на распаковку) ─────
pyinstaller --hiddenimport openpyxl --collect-submodules pptx --collect-data pptx --exclude-module pandas --exclude-module numpy --exclude-module matplotlib --exclude-module scipy --exclude-module IPython --exclude-module tkinter -F --windowed --onefile --version-file file_version_info.txt --icon=icon_file.ico main.py
