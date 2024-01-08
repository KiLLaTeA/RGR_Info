import json
import locale
import os
import time

from datetime import datetime as dt
from docxtpl import DocxTemplate

from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


def fill_doc():
    locale.setlocale(locale.LC_ALL, '')
    doc = DocxTemplate("template.docx")
    user = json.load(open(os.path.join(os.getcwd(), 'based.json'), encoding='utf8'))
    if not os.path.isdir(os.path.join(os.getcwd(), 'Преподаватели')):
        os.mkdir(os.path.join(os.getcwd(), 'Преподаватели'))
    for usr in user:
        print(f'\r[~] Идёт заполнение документов на сотрудника: {usr["surname"]}', end='')
        time.sleep(1)
        data = {
            'number_of_order': usr["number_of_order"],
            'date_of_order': usr["date_of_order"],
            'surname': usr["surname"],
            'name': usr["name"],
            'middlename': usr["middlename"],
            'date_of_birth': usr["date_of_birth"],
            'pension': usr["pension"],
            'INN': usr["INN"],
            'number_of_attestation_VAK': usr["number_of_attestation_VAK"],
            'number_of_diplome_VAK': usr["number_of_diplome_VAK"],
            'workplace_and_job': usr["workplace_and_job"],
            'index_of_address': usr["index_of_address"],
            'passport': usr["passport"],
            'who_give_it': usr["who_give_it"],
            'date_begin': usr["date_begin"],
            'date_end': usr["date_end"],
            'discipline': usr["discipline"],
            'department': usr["department"],
            'date_exercise': usr["date_exercise"],
            'hours_from_to': usr["hours_from_to"],
            'type': usr["type"],
            'group_and_form_education': usr["group_and_form_education"],
            'hours_count': usr["hours_count"],
            'lection_pay': usr["lection_pay"],
            'practice_pay': usr["practice_pay"],
            'lab_pay': usr["lab_pay"],
            'course_pay': usr["course_pay"],
            'RGR_pay': usr["RGR_pay"],
            'review_pay': usr["review_pay"],
            'diplome_pay': usr["diplome_pay"],
            'zachet_pay': usr["zachet_pay"],
            'exam_pay': usr["exam_pay"],
            'practice_study_pay': usr["practice_study_pay"],
            'practice_work_pay': usr["practice_work_pay"],
            'practice_before_diplome_pay': usr["practice_before_diplome_pay"],
            'hours_all': usr["hours_all"],
            'date_today': usr["date_today"],
            'sign_why_not': usr["sign_why_not"],
            'head_of_department_name': usr["head_of_department_name"],
            'sign_hod': usr["sign_hod"],
            'decan_name': usr["decan_name"],
            'summ_hour': usr["summ_hour"],
            'summ_all': usr["summ_all"],
            'sign_decan': usr["sign_decan"],
            'head_edu': usr["head_edu"],
            'vice_rector': usr["vice_rector"]
        }
        doc.render(data)
        doc.save(os.path.join(os.getcwd(), 'Преподаватели',
                              f'{usr["surname"]} {usr["name"]} {usr["middlename"]}.docx'))


def main():
    fill_doc()
    print('\n[!] Все данные обработаны!')


if __name__ == "__main__":
    main()
