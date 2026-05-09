import os
import sys
import traceback
import statistics
from datetime import datetime

from pptx import Presentation
from pptx.chart.data import XyChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor

import openpyxl

# ---------------------------------------------------------------------------
# Paths — resolve relative to the executable or script, not cwd
# ---------------------------------------------------------------------------
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

BASE_FOLDER = os.path.join(APP_DIR, 'example_folder')
XLSX_PATH   = os.path.join(BASE_FOLDER, 'velocity.xlsx')
OUTPUT_PATH = os.path.join(APP_DIR, 'velocity.pptx')
ERROR_PATH  = os.path.join(APP_DIR, 'velocity.txt')

SHEET1 = 'Sheet1'
SHEET2 = 'Sheet2'

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
IMAGE_EXTENSIONS = ('.jpg', '.jpeg', '.gif', '.png', '.bmp',
                    '.tiff', '.tif', '.webp', '.heic', '.avif')

COLUMN_RENAME = {
    'rank':   'Рейтинг',
    'change': 'Изм vs. LY',
    'ros':    'Продажи на ТТ',
    'nd':     'ND, LM',
    'price':  'Микс цена',
}

CATEGORY_COLORS = [
    RGBColor(0x00, 0x6D, 0x50),
    RGBColor(0x44, 0x70, 0xF3),
    RGBColor(0xFF, 0xC0, 0x00),
    RGBColor(0xEB, 0x23, 0x16),
]

FONT         = 'Calibri'
SLIDE_WIDTH  = Cm(29.7)
SLIDE_HEIGHT = Cm(21.0)
DEFAULT_GW   = Cm(20)
DEFAULT_GH   = Cm(15)


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------
def uniquify(path):
    base, ext = os.path.splitext(path)
    counter = 1
    while os.path.exists(path):
        path = f"{base}_{counter}{ext}"
        counter += 1
    return path


def write_error(user_message, technical_detail=None):
    lines = [f"ОШИБКА: {user_message}\n"]
    if technical_detail:
        lines.append(f"\nТехническая информация:\n{technical_detail}\n")
    lines.append(f"\nВремя: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    content = "".join(lines)
    try:
        with open(ERROR_PATH, 'w', encoding='utf-8') as f:
            f.write(content)
    except Exception:
        pass
    print(content)


def get_sku_images_map():
    """Return {sku_name_without_ext: full_path} for all images in BASE_FOLDER."""
    result = {}
    if not os.path.isdir(BASE_FOLDER):
        return result
    for fname in os.listdir(BASE_FOLDER):
        if fname.lower().endswith(IMAGE_EXTENSIONS):
            result[os.path.splitext(fname)[0]] = os.path.join(BASE_FOLDER, fname)
    return result


def get_blank_layout(prs):
    for layout in prs.slide_layouts:
        if layout.name == 'Blank':
            return layout
    return prs.slide_layouts[-1]


# ---------------------------------------------------------------------------
# Data reading
# ---------------------------------------------------------------------------
def _parse_nd(val):
    if isinstance(val, str):
        return float(val.strip().strip('%'))
    return float(val) if val is not None else 0.0


def read_sheet1():
    wb = openpyxl.load_workbook(XLSX_PATH, read_only=True, data_only=True)

    if SHEET1 not in wb.sheetnames:
        wb.close()
        raise KeyError(
            f"Лист '{SHEET1}' не найден в файле velocity.xlsx. "
            f"Убедитесь, что в файле есть лист с именем Sheet1."
        )

    ws = wb[SHEET1]
    rows_iter = ws.iter_rows(values_only=True)
    header_row = next(rows_iter, None)

    if header_row is None:
        wb.close()
        raise ValueError(
            "Лист Sheet1 пустой — нет строк с данными. "
            "Добавьте заголовки и данные в таблицу на листе Sheet1."
        )

    headers = [str(h).strip().lower() if h is not None else '' for h in header_row]
    required = ['sku', 'nd', 'ros', 'rank', 'change', 'price']
    missing = [c for c in required if c not in headers]
    if missing:
        wb.close()
        raise KeyError(
            f"В таблице не найдены столбцы: {', '.join(missing)}. "
            f"На листе Sheet1 должны быть столбцы: sku, nd, ros, rank, change, price "
            f"(регистр не важен). Проверьте названия заголовков."
        )

    idx = {h: i for i, h in enumerate(headers)}
    data = []

    for row in rows_iter:
        sku_val = row[idx['sku']]
        if sku_val is None:
            continue
        data.append({
            'sku':    str(sku_val).strip(),
            'nd':     _parse_nd(row[idx['nd']]),
            'ros':    float(row[idx['ros']])    if row[idx['ros']]    is not None else 0.0,
            'rank':   int(row[idx['rank']])     if row[idx['rank']]   is not None else 0,
            'change': row[idx['change']],
            'price':  row[idx['price']],
        })

    wb.close()
    return data


def read_sheet2():
    """Return (graph_width, graph_height, thresholds_list)."""
    try:
        wb = openpyxl.load_workbook(XLSX_PATH, read_only=True, data_only=True)
        if SHEET2 not in wb.sheetnames:
            wb.close()
            return DEFAULT_GW, DEFAULT_GH, []

        ws = wb[SHEET2]

        # A2, B2 — graph frame dimensions
        row2 = next(ws.iter_rows(min_row=2, max_row=2, min_col=1, max_col=5, values_only=True), None)
        wb.close()

        if row2 is None:
            return DEFAULT_GW, DEFAULT_GH, []

        a2, b2, c2, d2, e2 = (row2 + (None, None, None, None, None))[:5]
        gw = Cm(float(a2)) if a2 is not None else DEFAULT_GW
        gh = Cm(float(b2)) if b2 is not None else DEFAULT_GH

        thresholds = []
        for val in [c2, d2, e2]:
            if val is not None:
                thresholds.append(int(val))
            else:
                break

        return gw, gh, thresholds

    except Exception:
        return DEFAULT_GW, DEFAULT_GH, []


# ---------------------------------------------------------------------------
# Drawing helpers
# ---------------------------------------------------------------------------
def add_cross(slide, left, top, width, height):
    """Red X cross — used when an image is missing."""
    for x1, y1, x2, y2 in [
        (left, top, left + width, top + height),
        (left + width, top, left, top + height),
    ]:
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
        ln.line.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        ln.line.width = Pt(1.5)


def draw_border(slide, width, height):
    """Rectangular border starting at 0,0."""
    for x1, y1, x2, y2 in [
        (0, 0, width, 0),
        (width, 0, width, height),
        (width, height, 0, height),
        (0, height, 0, 0),
    ]:
        slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)


def add_median_crosshair(slide, chart_frame, median_nd, median_ros, max_nd, max_ros):
    """
    Draws red horizontal + vertical lines at the median position.
    The plot-area offsets are heuristic fractions of the chart bounding box.
    """
    if max_nd <= 0 or max_ros <= 0:
        return

    left_frac, right_frac = 0.14, 0.03
    top_frac, bottom_frac = 0.05, 0.14

    cl, ct = chart_frame.left, chart_frame.top
    cw, ch = chart_frame.width, chart_frame.height

    pl = cl + int(cw * left_frac)
    pt = ct + int(ch * top_frac)
    pw = int(cw * (1 - left_frac - right_frac))
    ph = int(ch * (1 - top_frac - bottom_frac))

    y_med = pt + ph - int(median_ros / max_ros * ph)
    h = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, pl, y_med, pl + pw, y_med)
    h.line.color.rgb = RGBColor(0xE8, 0x00, 0x00)
    h.line.width = Pt(1.5)

    x_med = pl + int(median_nd / max_nd * pw)
    v = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x_med, pt, x_med, pt + ph)
    v.line.color.rgb = RGBColor(0xE8, 0x00, 0x00)
    v.line.width = Pt(1.5)


def _set_cell_text(cell, text, font_size=Pt(8), bold=False):
    tf = cell.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = str(text) if text is not None else ''
    run.font.name = FONT
    run.font.size = font_size
    run.font.bold = bold


# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------
def build_slide_one(prs, sku_data):
    """Slide 1: XY scatter chart (nd vs ros) with red median crosshair."""
    slide = prs.slides.add_slide(get_blank_layout(prs))

    chart_data = XyChartData()
    for rec in sku_data:
        s = chart_data.add_series(str(rec['sku']))
        s.add_data_point(rec['nd'], rec['ros'])

    chart_frame = slide.shapes.add_chart(
        XL_CHART_TYPE.XY_SCATTER, 0, 0, SLIDE_WIDTH, SLIDE_HEIGHT, chart_data
    )

    chart = chart_frame.chart
    chart.value_axis.minimum_scale    = 0.0
    chart.category_axis.minimum_scale = 0.0

    nd_vals  = [r['nd']  for r in sku_data]
    ros_vals = [r['ros'] for r in sku_data]
    med_nd   = statistics.median(nd_vals)
    med_ros  = statistics.median(ros_vals)

    add_median_crosshair(
        slide, chart_frame,
        med_nd, med_ros,
        max(nd_vals), max(ros_vals),
    )


def build_slide_two(prs, sku_data, graph_w, graph_h, sku_images_map):
    """Slide 2: images placed by nd/ros coordinates inside a border frame."""
    slide = prs.slides.add_slide(get_blank_layout(prs))
    draw_border(slide, graph_w, graph_h)

    nd_vals  = [r['nd']  for r in sku_data]
    ros_vals = [r['ros'] for r in sku_data]
    max_nd   = max(nd_vals) if nd_vals else 1
    max_ros  = max(ros_vals) if ros_vals else 1
    img_h    = Cm(2)
    not_found = set()

    for rec in sku_data:
        sku = rec['sku']
        x = int(rec['nd']  / max_nd  * graph_w)
        y = int(graph_h - rec['ros'] / max_ros * graph_h)
        img_path = sku_images_map.get(sku)
        if img_path:
            try:
                slide.shapes.add_picture(img_path, x, y, height=img_h)
                continue
            except Exception:
                pass
        not_found.add(sku)
        add_cross(slide, x, y, img_h, img_h)

    return not_found


def build_slide_three(prs, sku_data, thresholds, sku_images_map):
    """Slide 3: table sorted by rank with images in the left column."""
    slide = prs.slides.add_slide(get_blank_layout(prs))

    sorted_data = sorted(sku_data, key=lambda r: r['rank'])
    n_rows = len(sorted_data)
    if n_rows == 0:
        return set()

    margin_l = Cm(0.3)
    margin_t = Cm(0.3)
    margin_r = Cm(0.3)
    img_col_w = Cm(2.0)
    gap       = Cm(0.2)

    table_left  = margin_l + img_col_w + gap
    table_top   = margin_t
    table_width = SLIDE_WIDTH - table_left - margin_r

    header_h   = Cm(0.9)
    content_h  = SLIDE_HEIGHT - margin_t - Cm(0.3) - header_h
    row_h      = max(Cm(0.5), min(Cm(1.8), int(content_h / n_rows)))
    total_h    = header_h + row_h * n_rows

    col_order = list(COLUMN_RENAME.keys())
    n_cols    = len(col_order)

    tbl_shape = slide.shapes.add_table(
        n_rows + 1, n_cols,
        int(table_left), int(table_top),
        int(table_width), int(total_h),
    )
    tbl = tbl_shape.table

    col_width_fracs = {'rank': 0.10, 'change': 0.22, 'ros': 0.26, 'nd': 0.20, 'price': 0.22}
    for ci, key in enumerate(col_order):
        tbl.columns[ci].width = int(table_width * col_width_fracs[key])

    tbl.rows[0].height = int(header_h)
    for ci, key in enumerate(col_order):
        _set_cell_text(tbl.cell(0, ci), COLUMN_RENAME[key], font_size=Pt(9), bold=True)

    def category_index(rank):
        for i, threshold in enumerate(thresholds):
            if rank <= threshold:
                return i
        return len(thresholds) if thresholds else -1

    not_found = set()

    for ri, rec in enumerate(sorted_data):
        row_idx = ri + 1
        tbl.rows[row_idx].height = int(row_h)

        cat = category_index(rec['rank'])
        color = CATEGORY_COLORS[cat] if 0 <= cat < len(CATEGORY_COLORS) else None

        for ci, key in enumerate(col_order):
            cell = tbl.cell(row_idx, ci)
            _set_cell_text(cell, rec.get(key, ''))
            if color:
                cell.fill.solid()
                cell.fill.fore_color.rgb = color

        sku      = rec['sku']
        img_top  = int(table_top + header_h + ri * row_h)
        img_left = int(margin_l)
        img_path = sku_images_map.get(sku)

        if img_path:
            try:
                slide.shapes.add_picture(img_path, img_left, img_top, height=int(row_h))
                continue
            except Exception:
                pass
        not_found.add(sku)
        add_cross(slide, img_left, img_top, int(img_col_w), int(row_h))

    return not_found


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
def main():
    try:
        if not os.path.exists(XLSX_PATH):
            write_error(
                "Файл velocity.xlsx не найден.\n"
                f"Пожалуйста, положите файл velocity.xlsx в папку example_folder "
                f"рядом с исполняемым файлом программы.\n"
                f"Ожидаемый путь: {XLSX_PATH}"
            )
            return

        try:
            sku_data = read_sheet1()
        except (KeyError, ValueError) as e:
            write_error(str(e))
            return
        except Exception as e:
            write_error(
                "Не удалось прочитать данные из файла velocity.xlsx.\n"
                "Убедитесь, что:\n"
                "  1. Файл не открыт в Excel прямо сейчас.\n"
                "  2. Лист Sheet1 содержит корректную таблицу с заголовками.",
                technical_detail=traceback.format_exc(),
            )
            return

        if not sku_data:
            write_error(
                "Таблица на листе Sheet1 пустая — нет строк с данными. "
                "Добавьте данные под строкой заголовков."
            )
            return

        graph_w, graph_h, thresholds = read_sheet2()
        sku_images_map = get_sku_images_map()

        prs = Presentation()
        prs.slide_width  = SLIDE_WIDTH
        prs.slide_height = SLIDE_HEIGHT

        build_slide_one(prs, sku_data)
        not_found_2 = build_slide_two(prs, sku_data, graph_w, graph_h, sku_images_map)
        not_found_3 = build_slide_three(prs, sku_data, thresholds, sku_images_map)

        out_path = uniquify(OUTPUT_PATH)
        prs.save(out_path)
        print(f"Презентация сохранена: {out_path}")

        all_missing = not_found_2 | not_found_3
        if all_missing:
            ext_list = ', '.join(IMAGE_EXTENSIONS)
            lines = [
                "Некоторые изображения не найдены в папке example_folder.\n"
                "Для каждого SKU ниже добавьте файл изображения в папку example_folder "
                f"рядом с программой (поддерживаемые форматы: {ext_list}):\n\n"
            ]
            for sku in sorted(all_missing):
                lines.append(f"  - Изображение для SKU '{sku}': файл должен называться '{sku}.jpg' "
                              f"(или другой поддерживаемый формат).\n")
            msg = "".join(lines)
            txt_path = uniquify(out_path.replace('.pptx', '.txt'))
            with open(txt_path, 'w', encoding='utf-8') as f:
                f.write(msg)
            print(msg)

    except Exception:
        write_error(
            "Произошла непредвиденная ошибка при создании презентации.\n"
            "Убедитесь, что файл velocity.xlsx корректен и не открыт в Excel.",
            technical_detail=traceback.format_exc(),
        )


if __name__ == "__main__":
    main()
