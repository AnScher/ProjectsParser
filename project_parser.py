# -*- coding: utf-8 -*-
import pandas as pd
import json
import logging

DEFAULT_PROJECT = "default_project.json"

logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
file_handler = logging.FileHandler("projects_parser.log")
file_handler.setLevel(logging.DEBUG)
file_formatter = logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)


class Button:
    def __init__(self, service_id, description, value):
        self.service_id = self._get_service_ids(
            service_id)  # Делаем список номеров(позиций) кнопок в проекте. В каждом экземпляре они должны начинаться с 1
        self.description = self._check_description(description)  # Удаляем пробелы по краям
        self.value = list(service_id)

    def _get_service_ids(self, service_id):
        return [number for number, value in enumerate(service_id,
                                                      start=1)]

    def _check_description(self, description):
        return [el.strip() for el in description]


class Controller:
    def __init__(self, excel_file_name, project_ids, sheet_names, is_default_project_required):
        self.excel_file_name = excel_file_name
        self.project_ids = project_ids
        self.sheet_names = sheet_names
        self.if_need_use_default_project = is_default_project_required

    def run(self):
        logger.debug("Запуск скрипта")
        logger.debug("Загрузка Excel-документа")
        button_info = self._load_excel_book(self.excel_file_name, self.sheet_names)
        logger.debug("Формирование информации о кнопках")
        json_projects = self._load_project_from_pc(button_info, self.if_need_use_default_project)
        logger.debug("Наполнение проектов")
        self._dump_projects(json_projects)
        logger.debug("Сохранение проектов")

    def _get_button_obj_from_pc(self, pd_data_frame):
        try:
            return Button(pd_data_frame["ID"], pd_data_frame["SERVICE"], pd_data_frame["PRICE"])
        except Exception as e:
            logger.debug(pd_data_frame)
            logger.error(e)

    def _load_excel_book(self, excel_file_name, sheet_names):
        try:
            xls_book = pd.ExcelFile(excel_file_name)
            button_info_list = []

            with pd.ExcelFile(xls_book) as xls:
                for sheet in sheet_names:
                    sheet_obj = pd.read_excel(xls, "{}".format(sheet))
                    button_info_list_tmp = []  # Создал новый временный список, потому что объект кнопки не итерабельный
                    button_info_list_tmp.append(self._get_button_obj_from_pc(sheet_obj))
                    button_info_list.append(self._generate_button_info_from_file(button_info_list_tmp))

            return zip(self.project_ids, button_info_list)

        except FileNotFoundError as ex:
            logger.error("Не найден файл excel-книги!")
            print("Не найден файл excel-книги!")

    def _generate_button_info_from_file(self, button_info):
        my_dict = {"Buttons": []}
        my_dict_buttons = []
        for info in button_info:
            for info in zip(info.service_id, info.description, info.value):
                my_dict_buttons.append({"id": info[0],
                                        "text": info[1],
                                        "value": info[2]})
        my_dict["Buttons"] = my_dict_buttons
        return my_dict

    def _load_project_from_pc(self, button_info, is_default_project_required):
        correct_projects = []
        for proj, info in button_info:
            if not is_default_project_required:
                with open(proj, encoding='UTF-8') as project:
                    old_project = json.load(project)
                    old_project['Steps'][0]['Buttons'] = info['Buttons']
                    correct_projects.append((proj, old_project))
            else:
                with open(DEFAULT_PROJECT, encoding='UTF-8') as project:
                    old_project = json.load(project)
                    old_project['Steps'][0]['Buttons'] = info['Buttons']
                    correct_projects.append((proj, old_project))
        return correct_projects

    def _dump_projects(self, json_projects):
        for proj, info in json_projects:
            with open(proj, 'w') as project:
                json.dump(info, project, ensure_ascii=False)  # , indent=4


if __name__ == '__main__':
    excel_file_name = "2.xlsx"
    project_ids = ["2000441.json", "2000442.json", "2000443.json"]  # Название новых/старых проектов по услугам
    sheet_names = ["5495", "5496", "5497"]  # Название листов в Excek-документе
    controller = Controller("2.xlsx", project_ids, sheet_names,
                            is_default_project_required=True)  # is_default_project_required = True  # Если нужно использовать дефолтный проект
    controller.run()
