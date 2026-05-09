### velocity pptx auto generator

Читает `example_folder/velocity.xlsx` и генерирует презентацию `velocity.pptx` рядом с программой.

---

## Структура файлов

```
velocity.exe  (или main.py)
example_folder/
    velocity.xlsx
    <sku>.jpg / <sku>.png / ...   ← изображения, имя файла = значение столбца sku
```

---

## Слайды

**Слайд 1** — XY-диаграмма рассеяния: ось X = `nd`, ось Y = `ros`, подписи = `sku`.
Медиана отображается красным перекрестием (горизонтальная + вертикальная линии).

**Слайд 2** — те же данные, но вместо точек — картинки SKU, расположенные внутри рамки.
Размеры рамки задаются в `Sheet2` ячейки `A2` (ширина, см) и `B2` (высота, см); по умолчанию 20 × 15 см.

**Слайд 3** — таблица, отсортированная по столбцу `rank` (по возрастанию), с картинками слева.
Столбцы: Рейтинг, Изм vs. LY, Продажи на ТТ, ND LM, Микс цена.
Строки раскрашиваются по категориям, заданным в `Sheet2` ячейки `C2:E2` — пороговые значения rank для 4 цветовых групп (зелёный / синий / жёлтый / красный).

---

## Формат xlsx

### Sheet1 — обязательные столбцы (порядок не важен):
| sku | nd | ros | rank | change | price |

### Sheet2 — настройки (необязательно):
| A1 | B1 | C1 | D1 | E1 |
|---|---|---|---|---|
| любое | любое | любое | любое | любое |
| ширина рамки (см) | высота рамки (см) | порог кат.1 | порог кат.2 | порог кат.3 |

Если ячейки A2/B2 пустые — используются значения по умолчанию 20 × 15 см.
Если C2/D2/E2 пустые — строки таблицы не раскрашиваются.

---

## Поддерживаемые форматы изображений

`.jpg` `.jpeg` `.png` `.gif` `.bmp` `.tiff` `.tif` `.webp` `.heic` `.avif`

---

## Сборка исполняемого файла

Требования: `pip install pyinstaller openpyxl python-pptx`
(**pandas и numpy не нужны**)

### Рекомендуется: папка `dist/main/` — быстрый запуск

```
pyinstaller --hiddenimport openpyxl --collect-submodules pptx --collect-data pptx --exclude-module pandas --exclude-module numpy --exclude-module matplotlib --exclude-module scipy --exclude-module IPython --exclude-module tkinter --windowed --version-file file_version_info.txt --icon=icon_file.ico main.py
```

Результат: папка `dist/main/` с `main.exe` и всеми зависимостями (~40 MB).
**Переносится на ПК без Python целиком — копировать всю папку `dist/main/`.**

### Альтернатива: один файл `.exe` — медленнее на запуск

```
pyinstaller --hiddenimport openpyxl --collect-submodules pptx --collect-data pptx --exclude-module pandas --exclude-module numpy --exclude-module matplotlib --exclude-module scipy --exclude-module IPython --exclude-module tkinter -F --windowed --onefile --version-file file_version_info.txt --icon=icon_file.ico main.py
```

> `--onefile` при каждом запуске распаковывается во временную папку (~3–10 сек).
> `--onedir` (без `-F`) запускается мгновенно — рекомендуется для распространения.
