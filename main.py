import os

from pptx import Presentation
from pptx.chart.data import XyChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_CONNECTOR
from pptx.util import Cm

FILE_NAME = 'velocity.pptx'
BASE_FOLDER = './example_folder'
csv_path = './example_folder/velocity.csv'
xlsx_path = "./example_folder/velocity.xlsx"
sheet_name = "Sheet1"


def uniquify(file_name):
    """
    Uniquify file name. Return filename but if file with this name exists add number
    :param file_name: base file_name
    :return: unique file name
    """
    filename, extension = os.path.splitext(file_name)
    counter = 1

    while os.path.exists(file_name):
        file_name = filename + "_" + str(counter) + extension
        counter += 1

    return file_name


def percent_to_float(x) -> float:
    """
    Parse percent value to float
    :param x: percent value
    :return: float value representation
    """
    return float(x.strip('%'))  # /100


def get_sku_images_list():
    """
    Form list of sku images names
    :return: List[str] of sku image names
    """
    sku_list_images = []
    for address, dirs, files in os.walk(BASE_FOLDER):
        for filename in files:
            # if (filename.split('.')[-1] not in ('xlsx', 'csv', 'xls')):
            if filename.lower().endswith(('.jpg', '.jpeg', '.gif', '.png', '.bmp')):
                sku_list_images.append(filename)
    return sku_list_images


def generate_square(slide_one, height, width):
    """
    Add rectangle to slide with its width and height
    :param slide_one: slide to add image
    :param height: height of square
    :param width: width of square
    :return: None
    """
    line1 = slide_one.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, 0, width, 0)
    line2 = slide_one.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, width, 0, width, height)
    line3 = slide_one.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, width, height, 0, height)
    line4 = slide_one.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, 0, height, 0, 0)


def add_image_to_slide(slide_one, sku_name, context):
    """
    Add sku images to slide_one
    :param slide_one: slide to add image
    :param sku_name: sku name to add
    :param context: context dict
    :return: None
    """
    for product in context['sku_list_images']:
        if sku_name == product.split('.')[0]:
            # print(context['velocity_graph'])
            if context['velocity_graph']:
                margin_left = context['cur_val_ros'] / context['max_val_ros'] * context['graph_width']
                margin_top = context['graph_height'] - context['cur_val_nd'] / context['max_val_nd'] * context['graph_height']
            else:
                height_img_count = 15
                margin_left = (context['index'] // height_img_count) * Cm(5)
                margin_top = context['index'] * Cm(1) - (context['index'] // height_img_count) * Cm(height_img_count)

            image_item_path = os.path.join(BASE_FOLDER, product)
            # print(image_item_path, margin_left, margin_top, context['image_height'], context['cur_val_nd'], context['cur_val_ros'])
            pic = slide_one.shapes.add_picture(image_item_path, margin_left, margin_top, height=context['image_height'])
            return
    context['not_found_sku'].add(sku_name)


def main() -> None:
    """
    Entrypoint method, builds pptx presentation
    :return: None
    """

    prs = Presentation()
    title_slide_layout = prs.slide_layouts[0]
    slide_one = prs.slides.add_slide(title_slide_layout)

    for shape in slide_one.shapes:
        if shape.has_text_frame:
            shape.left = Cm(25)
            
    title = slide_one.shapes.title
    subtitle = slide_one.placeholders[1]

    title.margin_left = Cm(15)
    subtitle.text = ""

    height = Cm(15)
    width = Cm(20)

    sku_list_excel = None
    import pandas as pd
    try:
        sku_list_excel = pd.read_excel(xlsx_path, sheet_name=sheet_name)
    except Exception as err_xlsx:
        print(f'Error reading excel file: {err_xlsx}')
        try:
            sku_list_excel = pd.read_csv(csv_path, sep=';')
        except Exception as err_csv:
            print(f'Error reading csv file: {err_csv}')

    if sku_list_excel is not None:
        sku_list_excel[['sku']] = sku_list_excel[['sku']].astype(str)
        sku_list_excel['nd'] = sku_list_excel.apply(
            lambda prod_row: percent_to_float(prod_row['nd']) if isinstance(prod_row['nd'], str) else prod_row['nd'], axis=1)

        chart_data = XyChartData()

        for index, row in sku_list_excel.iterrows():
            series_product = chart_data.add_series(f"sku: {row['sku']}")
            series_product.add_data_point(row['ros'], row['nd'])

        x, y, cx, cy = 0, 0, width, height

        chart = slide_one.shapes.add_chart(
            XL_CHART_TYPE.XY_SCATTER, x, y, cx, cy, chart_data
        ).chart

        next_slide_layout = prs.slide_layouts[2]
        slide_two = prs.slides.add_slide(next_slide_layout)

        for shape in slide_two.shapes:
            if shape.has_text_frame:
                shape.left = Cm(25)

        generate_square(slide_two, height, width)

        sku_list_images = get_sku_images_list()

        context = {
            "sku_list_images": sku_list_images,
            "max_val_nd": sku_list_excel.nd.max(),
            "max_val_ros": sku_list_excel.ros.max(),
            "cur_val_nd": None,
            "cur_val_ros": None,
            "image_height": Cm(2),
            "graph_height": height,
            "graph_width": width,
            "velocity_graph": True,
            "not_found_sku": set()
        }
        for index, row in sku_list_excel.iterrows():
            sku_name = row['sku']
            context['cur_val_nd'] = row['nd']
            context['cur_val_ros'] = row['ros']
            add_image_to_slide(slide_two, sku_name, context)

        # logger = logging.getLogger()
        # logger.addHandler(logging.StreamHandler(sys.stdout))
        if context['not_found_sku']:
            with open(f"{uniquify(FILE_NAME).split('.')[0]}.txt", "w") as file:
                for sku_item in context['not_found_sku']:
                    file.write(f'SKU image not found in folder: {sku_item}. (Or maybe it have different extension, not from list: .jpg/.jpeg/.gif/.png/.bmp\n')
                    # logger.info(f'SKU image not found in folder: {sku_item}. (Or maybe it have different extension, not from list: .jpg/.jpeg/.gif/.png/.bmp')
                    print(f'SKU image not found in folder: {sku_item}. (Or maybe it have different extension, not from list: .jpg/.jpeg/.gif/.png/.bmp')

        one_more_slide_layout = prs.slide_layouts[3]
        slide_three = prs.slides.add_slide(one_more_slide_layout)

        context['velocity_graph'] = False

        for shape in slide_three.shapes:
            if shape.has_text_frame:
                shape.left = Cm(25)

        # for shape in slide_three.shapes:
            # if shape.is_placeholder and shape.name == 'Content Placeholder 2':
            #     shape.height = Cm(15)
            #     shape.width = Cm(4)
            #     shape.left = Cm(20)

        for index, row in sku_list_excel.sort_values('ros', ascending=False, ignore_index=True).iterrows():
            context['index'] = index
            context['image_height'] = Cm(1)
            sku_name = row['sku']
            add_image_to_slide(slide_three, sku_name, context)

                    # text_frame = shape.text_frame
                    # p = text_frame.paragraphs[0]
                    # from pptx.enum.text import PP_ALIGN
                    # p.text = p.text + str(index+1) + ' - ' + str(sku_name) + '\n'
                    # p.alignment = PP_ALIGN.LEFT
                    # print(index)
                    #
                    # font = p.font
                    # font.name = 'Calibri'
                    # from pptx.util import Pt
                    # from pptx.dml.color import RGBColor
                    # from pptx.enum.dml import MSO_THEME_COLOR
                    # font.size = Pt(8)
                    # font.bold = False
                    # font.italic = None  # cause value to be inherited from theme
                    # # font.color.theme_color = MSO_THEME_COLOR.ACCENT_1




        # try:
        #     os.remove(FILE_NAME)
        # except OSError:
        #     pass

        prs.save(uniquify(FILE_NAME))
        # prs.save(FILE_NAME)

    else:
        print('No sku list found in xlsx/csv file')


if __name__ == "__main__":
    main()
    # os.system("pause")
