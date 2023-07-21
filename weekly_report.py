import shutil
from datetime import datetime, timedelta
import easygui
from openpyxl import load_workbook

FIRST_DATE = None
SEVENTH_DATE = None
DEST_DIR = None
PATH_CHOICE = None  # Move PATH_CHOICE to the global scope


def get_next_seven_dates():
    today = datetime.today()
    date_range = [today + timedelta(days=i) for i in range(1, 8)]
    date_strings = [date.strftime("%d.%m.%Y") for date in date_range]
    return date_strings


def copy_templates(src_dir_templates, destination):
    shutil.copytree(src_dir_templates, destination)


def process_report(source_file, destination_file, column_index, date_strings):
    dest_file = rf'{DEST_DIR}\{destination_file}'
    book_dest = load_workbook(filename=dest_file)
    sheet_dest = book_dest.active

    for i, date_str in enumerate(date_strings, start=1):
        src_file = rf'{PATH_CHOICE}\{date_str}\{source_file}'
        book_src = load_workbook(filename=src_file)
        sheet_src = book_src.active

        for row in range(3, 28):
            cell_src = sheet_src.cell(row=row, column=3)
            sheet_dest.cell(row=row + 2, column=column_index + i, value=cell_src.value)

    book_dest.save(filename=dest_file)


def main():
    global FIRST_DATE, SEVENTH_DATE, DEST_DIR, PATH_CHOICE

    easygui.msgbox(
        'Вас вітає програма для автоматизованого створення щонедільного звіту для НЕК. \nОберіть необхідну теку (місяць) де зберігаються файли НЕК'
    )
    PATH_CHOICE = easygui.diropenbox()

    FIRST_DATE = (datetime.today() + timedelta(1)).strftime("%d.%m.%Y")
    SEVENTH_DATE = (datetime.today() + timedelta(7)).strftime("%d.%m.%Y")
    DEST_DIR = rf'{PATH_CHOICE}\Недельный на {FIRST_DATE}-{SEVENTH_DATE}'
    print(f'path_choice: {PATH_CHOICE}')
    print(f'path_choice: {DEST_DIR}')

    date_strings = get_next_seven_dates()

    src_dir_templates = r"C:\Users\ihoraryku\Downloads\weekly_report\templates"
    copy_templates(src_dir_templates, DEST_DIR)

    create_bolgrad_solar(date_strings)
    create_voshod_volar(date_strings)
    create_dn_1(date_strings)
    create_dn_2(date_strings)
    create_eds(date_strings)
    create_le_1(date_strings)
    create_le_2(date_strings)
    create_neptun_solar(date_strings)
    create_pluton_solar(date_strings)
    create_pr_1(date_strings)
    create_pr_2(date_strings)
    create_fp(date_strings)
    create_fs(date_strings)
    create_evda(date_strings)

    easygui.msgbox('Звіт успішно створено!')


def create_bolgrad_solar(date_strings):
    source_file = 'Bolgrad_Solar.xlsx'
    destination_file = 'Болград Солар.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


def create_voshod_volar(date_strings):
    source_file = 'Voshod_Solar.xlsx'
    destination_file = 'Восход Солар.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


def create_dn_1(date_strings):
    source_file = 'Dunaiskaya_SES1.xlsx'
    destination_file = 'Дунайская СЭС-1.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


def create_dn_2(date_strings):
    source_file = 'Dunaiskaya_SES2.xlsx'
    destination_file = 'Дунайская СЭС-2.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


def create_eds(date_strings):
    source_file = 'EDS-SMART.xlsx'
    destination_file = 'ЕДС-Смарт Энерджи.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


def create_le_1(date_strings):
    source_file = 'Limanskaya_1.xlsx'
    destination_file = 'Лиманская Энерджи 1.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


def create_le_2(date_strings):
    source_file = 'Limanskaya_2.xlsx'
    destination_file = 'Лиманская Энерджи 2.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


def create_neptun_solar(date_strings):
    source_file = 'Neptun_Solar.xlsx'
    destination_file = 'Нептун Солар.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


def create_pluton_solar(date_strings):
    source_file = 'Pluton_Solar.xlsx'
    destination_file = 'Плутон Солар.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


def create_pr_1(date_strings):
    source_file = 'Priozernoe_1.xlsx'
    destination_file = 'Приозерное 1.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


def create_pr_2(date_strings):
    source_file = 'Priozernoe_2.xlsx'
    destination_file = 'Приозерное 2.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


def create_fp(date_strings):
    source_file = 'Franko_Pivi.xlsx'
    destination_file = 'Франко Пиви.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


def create_fs(date_strings):
    source_file = 'Franko_Solar.xlsx'
    destination_file = 'Франко Солар.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


def create_evda(date_strings):
    source_file = 'Evda_Energo.xlsx'
    destination_file = 'Эвда Энерго.xlsx'
    process_report(source_file, destination_file, 2, date_strings)


if __name__ == "__main__":
    main()
