# -*- coding: utf-8 -*-
import pandas as pd
import json

""" 1. Загружаем книгу, каждый лист кладем в отдельную переменную (циклом)
    2. Нужна проверка содержимого на наличие лидинг пробелов, запрещенных символов, дублей в id, проверка макс.длины поля
    3. Наполнение json_проекта содержимым
    4. Сохранение на комп в нужной кодировке
    5. Генерация БД, добавление header
    6. Сохранение БД на комп в нужной кодировке
"""

ProjectIdBySheetName = {"4891": "2000328.json",
                        "4893": "2000329.json",
                        "4894": "2000330.json",
                        "4895": "2000331.json",
                        "4896": "2000332.json",
                        "4897": "2000333.json"
                        }


class Button:
    def __init__(self, service_id, description, value):
        self.service_id = [number for number, value in enumerate(service_id, start=1)]  # Делаем список номеров(позиций) кнопок в проекте. В каждом экземпляре они должны начинаться с 1
        self.description = list(description)
        self.value = list(service_id)
        self.denied_chars = []

    def check_description(self, description):
        print(description)

    def check_value(self, value):
        pass


class Controller:
    def __init__(self, excel_file_name):
        self.excel_file_name = excel_file_name
        self.project_ids = []

    def run(self):
        button_info = self._load_excel_book(self.excel_file_name)
        json_projects = self._load_project_from_pc(button_info)
        self._dump_projects(json_projects)

    def _get_button_obj_from_pc(self, pd_data_frame):
        """
        Получает на вход датафрейм и возвращает экземплят класса Button
        :param pd_data_frame: Экземпляр класса Pd.DataFrame
        :return: Возвращает экземплят класса Button
        """
        return Button(pd_data_frame["ID"], pd_data_frame["SERVICE"], pd_data_frame["PRICE"])

    def _load_excel_book(self, excel_file_name):
        """

        :param excel_file_name:  Название excel-документа
        :return: Список кортежей Id-проекта/список словарей для замены
        """
        try:
            xls_book = pd.ExcelFile(excel_file_name)
            sheet_names = xls_book.sheet_names  # Получаем список названий листов в книге

            button_info_list = []

            with pd.ExcelFile(xls_book) as xls:
                for sheet in sheet_names:
                    sheet_obj = pd.read_excel(xls, "{}".format(sheet))
                    self.project_ids.append(ProjectIdBySheetName[
                                                sheet])  # Добавляем в список project_id id проекта, который соответствует данной конкрентой услуге
                    button_info_list_tmp = []  # Создал новый временный список, потому что объект кнопки не итерабельный
                    button_info_list_tmp.append(self._get_button_obj_from_pc(sheet_obj))
                    button_info_list.append(self._generate_button_info_from_file(button_info_list_tmp))

            return zip(self.project_ids, button_info_list)

        except FileNotFoundError as ex:
            print("Не найден файл excel-книги!")

    def _generate_button_info_from_file(self, button_info):
        """

        :param button_info: Экземпляр класса Button
        :return: Словарь, в котором ключу Buttons соответствует список словарей
        """
        my_dict = {"Buttons": []}
        my_dict_buttons = []
        for info in button_info:
            for info in zip(info.service_id, info.description, info.value):
                my_dict_buttons.append({"id": info[0],
                                        "text": info[1],
                                        "value": info[2]})
        my_dict["Buttons"] = my_dict_buttons
        return my_dict

    def _load_project_from_pc(self, button_info):
        """

        :param button_info: zip-объект, Id-проекта/список словарей для замены
        :return: Json-проект с ключом Buttons, переписанным новыми данными
        """
        correct_projects = []
        for proj, info in button_info:
            with open(proj, encoding='UTF-8') as project:
                old_project = json.load(project)
                old_project['Steps'][0]['Buttons'] = info['Buttons']
                correct_projects.append((proj, old_project))
        return correct_projects

    def _dump_projects(self, json_projects):
        for proj, info in json_projects:
            with open(proj, 'w') as project:
                json.dump(info, project, ensure_ascii=False)  # , indent=4


if __name__ == '__main__':
    controller = Controller("1.xlsx")
    controller.run()
