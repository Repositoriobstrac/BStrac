from driver_work_report_generator.settings import REPORT_BASE, CSV_FILE, LOGO_IMAGE_PATH
from driver_work_report_generator.excel import save
from driver_work_report_generator.tabula_extractor import pdf_to_csv_with_tabula
from driver_work_report_generator.utils import get_data_from_csv_file


def report_generator(pdf_file, driver):
    csvfile = CSV_FILE
    pdf_to_csv_with_tabula(pdf_file, csvfile)
    data_dict = get_data_from_csv_file(csvfile)
    logo_image_path = LOGO_IMAGE_PATH
    output_file = save(data_dict, driver, REPORT_BASE, logo_image_path)
    return output_file

