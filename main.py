import sys

from driver_work_report_generator import report_generator


if __name__ == '__main__':
    pdf_file = sys.argv[1]
    driver = sys.argv[2]
    output_file = report_generator(pdf_file, driver)
