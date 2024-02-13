from cx_Freeze import setup, Executable

base = None

executables = [Executable("main.py", base=base)]

packages = ["idna", "pptx", "os", "pandas"]
options = {
    'build_exe': {
        'packages': packages,
    },
}

setup(
    name="pptx_builder",
    options=options,
    version="0.0.1",
    description='csv_to_pptx_converter',
    executables=executables
)

# pyinstaller --hiddenimport pptx --collect-submodules pptx --collect-data pptx -F --windowed --onefile --version-file file_version_info.txt --icon=icon_file.ico main.py
# pyinstaller --hiddenimport pptx --collect-submodules pptx --collect-data pptx -F --onefile --version-file file_version_info.txt --icon=icon_file.ico main.py
