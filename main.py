import os
import wx
import wx.adv
import time
import math
import threading
import serial
import re
import wx.grid
import sqlite3
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException
import traceback
import logging


class ModelSelectorPanel(wx.Panel):
    """Панель для выбора модели электродвигателя из базы данных."""

    def __init__(self, parent, db_path="baseReda.db"):
        """Инициализация панели выбора модели.

        Args:
            parent: Родительский wx-объект (окно или другая панель)
            db_path (str): Путь к файлу базы данных SQLite (по умолчанию "baseReda.db")
        """
        # Вызов конструктора родительского класса wx.Panel
        super().__init__(parent)

        # Установка цвета фона панели (светло-голубой)
        self.SetBackgroundColour(wx.Colour(250, 252, 255))

        # Инициализация переменных для работы с БД
        self.db_path = db_path  # Сохраняем путь к БД
        self.conn = None  # Будущее соединение с БД
        self.cursor = None  # Будущий курсор для работы с БД
        self.all_models = []  # Список для хранения всех моделей
        self.current_model_id = None  # ID текущей выбранной модели

        # Создание вертикального контейнера для элементов управления
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Создание заголовка панели
        title = wx.StaticText(self, label="Выбор модели ЭД")
        # Настройка шрифта заголовка (размер 14, жирный)
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        # Установка цвета текста (темно-синий)
        title.SetForegroundColour(wx.Colour(0, 50, 100))
        # Добавление заголовка в контейнер с отступами 10 пикселей со всех сторон
        sizer.Add(title, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        # Создание группы "Поиск модели ЭД"
        search_box = wx.StaticBox(self, label="Поиск модели ЭД")
        # Создание контейнера для элементов поиска
        search_sizer = wx.StaticBoxSizer(search_box, wx.VERTICAL)

        # Горизонтальный контейнер для поля ввода и кнопки поиска
        hbox = wx.BoxSizer(wx.HORIZONTAL)

        # Создание текстового поля для ввода поискового запроса
        self.search_text = wx.TextCtrl(self, style=wx.TE_PROCESS_ENTER)
        # Установка подсказки в поле ввода
        self.search_text.SetHint("Введите модель ЭД...")
        # Добавление поля ввода в горизонтальный контейнер с растягиванием
        hbox.Add(self.search_text, 1, wx.EXPAND | wx.RIGHT, 5)

        # Создание кнопки "Поиск"
        self.search_btn = wx.Button(self, label="Поиск")
        # Добавление кнопки в горизонтальный контейнер
        hbox.Add(self.search_btn, 0, wx.EXPAND)

        # Добавление горизонтального контейнера в контейнер поиска
        search_sizer.Add(hbox, 0, wx.EXPAND | wx.ALL, 5)
        # Добавление контейнера поиска в основной контейнер
        sizer.Add(search_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # Создание группы "Доступные модели"
        models_box = wx.StaticBox(self, label="Доступные модели")
        # Контейнер для списка моделей
        models_sizer = wx.StaticBoxSizer(models_box, wx.VERTICAL)

        # Создание списка моделей (ListBox) высотой 300 пикселей
        self.models_list = wx.ListBox(self, style=wx.LB_SINGLE, size=wx.Size(-1, 300))
        # Добавление списка в контейнер с отступами 5 пикселей
        models_sizer.Add(self.models_list, 1, wx.EXPAND | wx.ALL, 5)
        # Добавление контейнера списка в основной контейнер
        sizer.Add(models_sizer, 1, wx.EXPAND | wx.ALL, 5)

        # Создание кнопки "Выбрать модель"
        self.select_btn = wx.Button(self, label="Выбрать модель")
        # Добавление кнопки в основной контейнер с отступами сверху и снизу
        sizer.Add(self.select_btn, 0, wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, 10)

        # Установка основного контейнера для панели
        self.SetSizer(sizer)

        # Привязка событий:
        # - Изменение текста в поле поиска
        self.search_text.Bind(wx.EVT_TEXT, self.on_search_text)
        # - Нажатие Enter в поле поиска
        self.search_text.Bind(wx.EVT_TEXT_ENTER, self.on_search)
        # - Нажатие кнопки поиска
        self.search_btn.Bind(wx.EVT_BUTTON, self.on_search)
        # - Выбор элемента в списке
        self.models_list.Bind(wx.EVT_LISTBOX, self.on_model_select)
        # - Двойной клик по элементу списка
        self.models_list.Bind(wx.EVT_LISTBOX_DCLICK, self.on_double_click)
        # - Нажатие кнопки выбора
        self.select_btn.Bind(wx.EVT_BUTTON, self.on_select)

        # Подключение к базе данных
        self.connect_db()
        # Загрузка всех моделей из БД
        self.load_all_models()

        # Начальное состояние кнопки выбора (отключена)
        self.select_btn.Disable()

    def connect_db(self):
        """Устанавливает соединение с базой данных SQLite."""
        try:
            # Создание соединения с БД
            self.conn = sqlite3.connect(self.db_path)
            # Создание курсора для выполнения запросов
            self.cursor = self.conn.cursor()
        except sqlite3.Error as e:
            # Вывод сообщения об ошибке при проблемах с подключением
            wx.MessageBox(f"Ошибка подключения к базе данных: {str(e)}", "Ошибка", wx.OK | wx.ICON_ERROR)

    def load_all_models(self):
        """Загружает все модели электродвигателей из базы данных."""
        # Если курсор не создан, выходим из метода
        if not self.cursor:
            return

        try:
            # Выполнение SQL-запроса для получения всех уникальных моделей
            self.cursor.execute("SELECT DISTINCT Model FROM Base ORDER BY Model")
            # Сохранение результатов в список (только непустые значения)
            self.all_models = [row[0] for row in self.cursor.fetchall() if row[0]]
            # Заполнение списка моделей в интерфейсе
            self.models_list.Set(self.all_models)
        except sqlite3.Error as e:
            # Вывод сообщения об ошибке при проблемах с загрузкой
            wx.MessageBox(f"Ошибка загрузки моделей: {str(e)}", "Ошибка", wx.OK | wx.ICON_ERROR)

    def on_search_text(self, event):
        """Фильтрация списка моделей при вводе текста в поле поиска.

        Args:
            event: Событие wxPython (не используется)
        """
        # Получение текста из поля поиска (в нижнем регистре)
        search_str = self.search_text.GetValue().lower()

        if search_str:
            # Фильтрация списка моделей по введенному тексту
            filtered = [m for m in self.all_models if search_str in m.lower()]
            # Обновление списка в интерфейсе
            self.models_list.Set(filtered)
        else:
            # Если поле поиска пустое, показываем все модели
            self.models_list.Set(self.all_models)

    def on_search(self, event):
        """Обработка поиска модели (по нажатию Enter или кнопки).

        Args:
            event: Событие wxPython (не используется)
        """
        # Получение текста из поля поиска
        model = self.search_text.GetValue()
        if model:
            # Поиск точного совпадения (без учета регистра)
            for i, m in enumerate(self.models_list.GetItems()):
                if model.lower() == m.lower():
                    # Выделение найденной модели в списке
                    self.models_list.SetSelection(i)
                    # Вызов обработчика выбора модели
                    self.on_model_select(None)
                    return
            # Если модель не найдена, показываем сообщение
            wx.MessageBox(f"Модель '{model}' не найдена", "Внимание", wx.OK | wx.ICON_INFORMATION)

    def on_model_select(self, event):
        """Обработка выбора модели из списка.

        Args:
            event: Событие wxPython (не используется)
        """
        # Получение индекса выбранного элемента
        selection = self.models_list.GetSelection()
        if selection != wx.NOT_FOUND:
            # Активация кнопки выбора, если элемент выбран
            self.select_btn.Enable()

    def on_double_click(self, event):
        """Обработка двойного клика по модели в списке.

        Args:
            event: Событие wxPython
        """
        # Вызов обработчика выбора модели
        self.on_select(event)

    def on_select(self, event):
        """Обработка выбора модели (кнопка или двойной клик).

        Args:
            event: Событие wxPython (не используется)
        """
        # Получение индекса выбранного элемента
        selection = self.models_list.GetSelection()
        if selection != wx.NOT_FOUND:
            # Получение названия выбранной модели
            model = self.models_list.GetString(selection)
            try:
                # Запрос всех параметров выбранной модели из БД
                self.cursor.execute("SELECT * FROM Base WHERE Model = ?", (model,))
                row = self.cursor.fetchone()
                if row:
                    # Создание словаря параметров модели
                    params = {}
                    # Получение названий колонок из описания курсора
                    columns = [col[0] for col in self.cursor.description]

                    # Заполнение словаря параметров
                    for i, value in enumerate(row):
                        col_name = columns[i]
                        params[col_name] = value

                    # Получение родительского элемента (предполагается, что это вкладка)
                    parent_tab = self.GetParent()
                    # Если у родителя есть метод set_selected_model, вызываем его
                    if hasattr(parent_tab, 'set_selected_model'):
                        parent_tab.set_selected_model(params)
                else:
                    # Если параметры не найдены, показываем сообщение
                    wx.MessageBox(f"Параметры для модели '{model}' не найдены",
                                  "Ошибка", wx.OK | wx.ICON_WARNING)
            except sqlite3.Error as e:
                # Обработка ошибок при работе с БД
                wx.MessageBox(f"Ошибка загрузки параметров: {str(e)}",
                              "Ошибка", wx.OK | wx.ICON_ERROR)

    def __del__(self):
        """Деструктор - закрывает соединение с базой данных при удалении объекта."""
        if self.conn:
            self.conn.close()


class EDParametersPanel(wx.Panel):
    """Панель для отображения и редактирования параметров электродвигателя."""

    def __init__(self, parent, db_path="baseReda.db"):
        """Инициализация панели параметров.

        Args:
            parent: Родительский wx-объект
            db_path (str): Путь к файлу базы данных
        """
        super().__init__(parent)
        # Установка цвета фона панели
        self.SetBackgroundColour(wx.Colour(250, 252, 255))

        # Инициализация переменных для работы с БД
        self.db_path = db_path
        self.conn = None
        self.cursor = None
        self.current_model_id = None  # ID текущей модели
        self.param_controls = {}  # Словарь для хранения элементов управления параметрами

        # Создание основного вертикального контейнера
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Создание заголовка панели
        title = wx.StaticText(self, label="Номинальные параметры ЭД")
        # Настройка шрифта заголовка
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        # Установка цвета текста заголовка
        title.SetForegroundColour(wx.Colour(0, 50, 100))
        # Добавление заголовка в контейнер
        main_sizer.Add(title, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        # Создание прокручиваемой панели для параметров
        self.scrolled_panel = wx.ScrolledWindow(self, style=wx.VSCROLL)
        # Установка скорости прокрутки
        self.scrolled_panel.SetScrollRate(0, 20)
        # Создание табличного контейнера для параметров (2 колонки)
        self.params_sizer = wx.FlexGridSizer(cols=2, vgap=8, hgap=10)
        # Разрешение растягивания второй колонки
        self.params_sizer.AddGrowableCol(1)

        # Список параметров с их отображаемыми названиями
        self.param_names = [
            ("ID", "ID"),  # Скрытый параметр
            ("Model", "Модель"),
            ("Power_nom", "Номинальная мощность, кВт"),
            # ... остальные параметры ...
        ]

        # Создание элементов управления для каждого параметра
        for db_name, ru_name in self.param_names:
            if db_name == "ID":  # Пропускаем поле ID
                continue

            # Создание метки с названием параметра
            label = wx.StaticText(self.scrolled_panel, label=f"{ru_name}:")
            label.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL))

            # Создание поля ввода для параметра
            if db_name == "Model":
                # Поле "Модель" доступно только для чтения
                ctrl = wx.TextCtrl(self.scrolled_panel, style=wx.TE_READONLY)
            else:
                # Остальные поля доступны для редактирования
                ctrl = wx.TextCtrl(self.scrolled_panel)

            # Добавление элементов в контейнер
            self.params_sizer.Add(label, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 5)
            self.params_sizer.Add(ctrl, 1, wx.EXPAND | wx.RIGHT, 5)
            # Сохранение элемента управления в словаре
            self.param_controls[db_name] = ctrl

        # Установка контейнера для прокручиваемой панели
        self.scrolled_panel.SetSizer(self.params_sizer)
        # Добавление прокручиваемой панели в основной контейнер
        main_sizer.Add(self.scrolled_panel, 1, wx.EXPAND | wx.ALL, 5)

        # Создание кнопки сохранения изменений
        self.save_btn = wx.Button(self, label="Сохранить изменения")
        # Добавление кнопки в основной контейнер
        main_sizer.Add(self.save_btn, 0, wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, 10)

        # Установка основного контейнера для панели
        self.SetSizer(main_sizer)

        # Привязка события нажатия кнопки сохранения
        self.save_btn.Bind(wx.EVT_BUTTON, self.on_save)

        # Подключение к базе данных
        self.connect_db()

        # Начальное состояние кнопки сохранения (отключена)
        self.save_btn.Disable()

    def connect_db(self):
        """Устанавливает соединение с базой данных SQLite."""
        try:
            self.conn = sqlite3.connect(self.db_path)
            self.cursor = self.conn.cursor()
        except sqlite3.Error as e:
            wx.MessageBox(f"Ошибка подключения к базе данных: {str(e)}", "Ошибка", wx.OK | wx.ICON_ERROR)

    def set_parameters(self, params):
        """Заполняет поля параметров значениями из словаря.

        Args:
            params (dict): Словарь параметров модели
        """
        # Сохранение ID текущей модели
        self.current_model_id = params.get("ID", None)

        # Заполнение полей ввода значениями из словаря
        for db_name, ctrl in self.param_controls.items():
            value = params.get(db_name, "")
            ctrl.SetValue(str(value))

        # Активация кнопки сохранения
        self.save_btn.Enable()

    def on_save(self, event):
        """Обработчик сохранения изменений параметров.

        Args:
            event: Событие wxPython (не используется)
        """
        # Проверка, что модель выбрана
        if not self.current_model_id:
            wx.MessageBox("Не выбрана модель для сохранения", "Ошибка", wx.OK | wx.ICON_ERROR)
            return

        # Сбор измененных данных
        update_data = {}
        for db_name, ctrl in self.param_controls.items():
            if db_name != "Model":  # Поле "Модель" не обновляем
                update_data[db_name] = ctrl.GetValue()

        # Формирование SQL-запроса для обновления
        set_clause = ", ".join([f"{k} = ?" for k in update_data.keys()])
        values = list(update_data.values())
        values.append(self.current_model_id)  # Добавляем ID для условия WHERE

        try:
            # Выполнение SQL-запроса
            query = f"UPDATE Base SET {set_clause} WHERE ID = ?"
            self.cursor.execute(query, values)
            self.conn.commit()  # Подтверждение изменений
            wx.MessageBox("Изменения успешно сохранены", "Сохранено", wx.OK | wx.ICON_INFORMATION)
        except sqlite3.Error as e:
            # Обработка ошибок при сохранении
            wx.MessageBox(f"Ошибка сохранения данных: {str(e)}", "Ошибка", wx.OK | wx.ICON_ERROR)
            self.conn.rollback()  # Откат изменений при ошибке

    def __del__(self):
        """Деструктор - закрывает соединение с базой данных при удалении объекта."""
        if self.conn:
            self.conn.close()


class DatabaseTab(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent)
        self.SetBackgroundColour(wx.Colour(240, 245, 250))  # Основной фон

        self.db_path = "baseReda.db"
        self.conn = None
        self.cursor = None

        # Создаем соединение с базой данных
        self.connect_db()

        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Создаем таблицу для отображения данных
        self.grid = wx.grid.Grid(self, -1)
        self.grid.CreateGrid(0, 22)  # Создаем пустую таблицу
        self.grid.SetColLabelValue(0, "ID")
        self.grid.SetColLabelValue(1, "Model")
        self.grid.SetColLabelValue(2, "Power_nom")
        self.grid.SetColLabelValue(3, "U_nom")
        self.grid.SetColLabelValue(4, "I_nom")
        self.grid.SetColLabelValue(5, "Turning")
        self.grid.SetColLabelValue(6, "R_ColdWinding")
        self.grid.SetColLabelValue(7, "R_Insul")
        self.grid.SetColLabelValue(8, "U_accel")
        self.grid.SetColLabelValue(9, "BoringMoment")
        self.grid.SetColLabelValue(10, "U_k_z")
        self.grid.SetColLabelValue(11, "I_k_z")
        self.grid.SetColLabelValue(12, "P_h_h")
        self.grid.SetColLabelValue(13, "I_Iding")
        self.grid.SetColLabelValue(14, "U_Iding")
        self.grid.SetColLabelValue(15, "P_k_z")
        self.grid.SetColLabelValue(16, "Time_RunDown")
        self.grid.SetColLabelValue(17, "Vibri_evel")
        self.grid.SetColLabelValue(18, "TurningMoment")
        self.grid.SetColLabelValue(19, "P_HeatedWaste")
        self.grid.SetColLabelValue(20, "U_MinInsulWinding")
        self.grid.SetColLabelValue(21, "U_InsulWinding")

        main_sizer.Add(self.grid, 1, wx.EXPAND | wx.ALL, 5)

        # Панель с кнопками
        btn_panel = wx.Panel(self)
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.btn_add = wx.Button(btn_panel, label="Добавить запись")
        self.btn_delete = wx.Button(btn_panel, label="Удалить запись")
        self.btn_save = wx.Button(btn_panel, label="Сохранить изменения")
        self.btn_refresh = wx.Button(btn_panel, label="Обновить данные")

        # Стилизуем кнопки
        for btn in [self.btn_add, self.btn_delete, self.btn_save, self.btn_refresh]:
            btn.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
            btn.SetBackgroundColour(wx.Colour(70, 130, 180))  # Приятный синий
            btn.SetForegroundColour(wx.WHITE)

        btn_sizer.Add(self.btn_add, 0, wx.ALL, 5)
        btn_sizer.Add(self.btn_delete, 0, wx.ALL, 5)
        btn_sizer.Add(self.btn_save, 0, wx.ALL, 5)
        btn_sizer.Add(self.btn_refresh, 0, wx.ALL, 5)

        btn_panel.SetSizer(btn_sizer)
        main_sizer.Add(btn_panel, 0, wx.ALIGN_CENTER | wx.TOP, 10)

        # Привязка событий
        self.btn_add.Bind(wx.EVT_BUTTON, self.on_add)
        self.btn_delete.Bind(wx.EVT_BUTTON, self.on_delete)
        self.btn_save.Bind(wx.EVT_BUTTON, self.on_save)
        self.btn_refresh.Bind(wx.EVT_BUTTON, self.on_refresh)

        self.SetSizer(main_sizer)

        # Загрузка данных из базы
        self.load_data()

    def connect_db(self):
        """Установка соединения с базой данных"""
        try:
            self.conn = sqlite3.connect(self.db_path)
            self.cursor = self.conn.cursor()
            self.ensure_table_exists()
        except sqlite3.Error as e:
            wx.MessageBox(f"Ошибка подключения к базе данных: {str(e)}", "Ошибка", wx.OK | wx.ICON_ERROR)

    def ensure_table_exists(self):
        """Создание таблицы, если она не существует"""
        create_table_query = """
        CREATE TABLE IF NOT EXISTS Base (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            Model TEXT,
            Power_nom INTEGER,
            U_nom INTEGER,
            I_nom REAL,
            Turning INTEGER,
            R_ColdWinding REAL,
            R_Insul INTEGER,
            U_accel INTEGER,
            BoringMoment REAL,
            U_k_z INTEGER,
            I_k_z INTEGER,
            P_h_h REAL,
            I_Iding REAL,
            U_Iding REAL,
            P_k_z INTEGER,
            Time_RunDown REAL,
            Vibri_evel REAL,
            TurningMoment REAL,
            P_HeatedWaste REAL,
            U_MinInsulWinding INTEGER,
            U_InsulWinding INTEGER
        )
        """
        self.cursor.execute(create_table_query)
        self.conn.commit()

    def load_data(self):
        """Загрузка данных из базы в таблицу"""
        if not self.cursor:
            return

        try:
            self.cursor.execute("SELECT * FROM Base")
            rows = self.cursor.fetchall()

            # Очистка таблицы
            if self.grid.GetNumberRows() > 0:
                self.grid.DeleteRows(0, self.grid.GetNumberRows())

            # Добавление строк
            for row in rows:
                self.grid.AppendRows(1)
                row_index = self.grid.GetNumberRows() - 1
                for col_index, value in enumerate(row):
                    if value is None:
                        value = ""
                    self.grid.SetCellValue(row_index, col_index, str(value))
        except sqlite3.Error as e:
            wx.MessageBox(f"Ошибка загрузки данных: {str(e)}", "Ошибка", wx.OK | wx.ICON_ERROR)

    def on_add(self, event):  # noqa: unused-argument
        """Добавление новой пустой строки"""
        self.grid.AppendRows(1)

    def on_delete(self, event):  # noqa: unused-argument
        """Удаление выбранных строк"""
        selected_rows = self.grid.GetSelectedRows()
        if not selected_rows:
            wx.MessageBox("Выберите строки для удаления", "Внимание", wx.OK | wx.ICON_INFORMATION)
            return

        # Удаляем из базы данных
        for row in reversed(sorted(selected_rows)):
            try:
                row_id = int(self.grid.GetCellValue(row, 0))
                self.cursor.execute("DELETE FROM Base WHERE ID = ?", (row_id,))
                self.conn.commit()
            except (ValueError, sqlite3.Error) as e:
                wx.MessageBox(f"Ошибка удаления записи: {str(e)}", "Ошибка", wx.OK | wx.ICON_ERROR)

        # Обновляем таблицу
        self.load_data()

    def on_save(self, event):  # noqa: unused-argument
        """Сохранение изменений в базе данных"""
        if not self.cursor:
            return

        try:
            for row in range(self.grid.GetNumberRows()):
                row_id = self.grid.GetCellValue(row, 0)

                # Проверяем, существует ли запись в базе
                if row_id:
                    self.cursor.execute("SELECT * FROM Base WHERE ID = ?", (row_id,))
                    exists = self.cursor.fetchone()
                else:
                    exists = False

                # Подготавливаем данные
                data = []
                for col in range(1, 22):  # Пропускаем ID (col 0)
                    value = self.grid.GetCellValue(row, col)
                    # Преобразуем пустые строки в None
                    if value == "":
                        data.append(None)
                    else:
                        # Пытаемся преобразовать в число
                        try:
                            if col in [2, 3, 5, 7, 8, 10, 11, 15, 20, 21]:  # Целочисленные поля
                                data.append(int(value))
                            elif col in [4, 6, 9, 12, 13, 14, 16, 17, 18, 19]:  # Вещественные поля
                                data.append(float(value))
                            else:
                                data.append(value)
                        except ValueError:
                            data.append(value)

                # Обновление или вставка (параметры ЭД?)
                if exists:
                    update_query = """
                    UPDATE Base SET
                        Model = ?,
                        Power_nom = ?,
                        U_nom = ?,
                        I_nom = ?,
                        Turning = ?,
                        R_ColdWinding = ?,
                        R_Insul = ?,
                        U_accel = ?,
                        BoringMoment = ?,
                        U_k_z = ?,
                        I_k_z = ?,
                        P_h_h = ?,
                        I_Iding = ?,
                        U_Iding = ?,
                        P_k_z = ?,
                        Time_RunDown = ?,
                        Vibri_evel = ?,
                        TurningMoment = ?,
                        P_HeatedWaste = ?,
                        U_MinInsulWinding = ?,
                        U_InsulWinding = ?
                    WHERE ID = ?
                    """
                    self.cursor.execute(update_query, tuple(data + [row_id]))
                else:
                    insert_query = """
                    INSERT INTO Base (
                        Model, Power_nom, U_nom, I_nom, Turning, R_ColdWinding, 
                        R_Insul, U_accel, BoringMoment, U_k_z, I_k_z, P_h_h, 
                        I_Iding, U_Iding, P_k_z, Time_RunDown, Vibri_evel, 
                        TurningMoment, P_HeatedWaste, U_MinInsulWinding, U_InsulWinding
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """
                    self.cursor.execute(insert_query, tuple(data))
                    # Получаем ID новой записи
                    new_id = self.cursor.lastrowid
                    self.grid.SetCellValue(row, 0, str(new_id))

            self.conn.commit()
            wx.MessageBox("Изменения успешно сохранены", "Сохранено", wx.OK | wx.ICON_INFORMATION)

        except sqlite3.Error as e:
            wx.MessageBox(f"Ошибка сохранения данных: {str(e)}", "Ошибка", wx.OK | wx.ICON_ERROR)
            self.conn.rollback()

    def on_refresh(self, event):  # noqa: unused-argument
        """Обновление данных из базы"""
        self.load_data()

    def __del__(self):
        """Закрытие соединения при уничтожении объекта"""
        if self.conn:
            self.conn.close()


class ColdInputResistanceDialog(wx.Dialog):
    def __init__(self, parent, title, description):
        super().__init__(parent, title=title, size=wx.Size(600, 500))
        self.parent = parent
        self.ser = None
        self.measurement_thread = None
        self.running = False
        self.measurement_stop = False

        sizer = wx.BoxSizer(wx.VERTICAL)

        # Заголовок теста
        title_label = wx.StaticText(self, label=title)
        title_label.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sizer.Add(title_label, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        # Описание теста
        desc_label = wx.StaticText(self, label=description)
        desc_label.Wrap(550)
        sizer.Add(desc_label, 0, wx.ALL | wx.EXPAND, 10)

        # Параметры измерений
        params_box = wx.StaticBox(self, label="Параметры измерений")
        params_sizer = wx.StaticBoxSizer(params_box, wx.VERTICAL)
        grid = wx.FlexGridSizer(cols=3, vgap=10, hgap=15)
        grid.AddGrowableCol(1, 1)
        grid.AddGrowableCol(2, 1)

        # Заголовки таблицы
        grid.Add(wx.StaticText(self, label="Параметр"), 0, wx.ALIGN_CENTER)
        grid.Add(wx.StaticText(self, label="Измеренное значение"), 0, wx.ALIGN_CENTER)
        grid.Add(wx.StaticText(self, label="Допустимое значение"), 0, wx.ALIGN_CENTER)

        # Напряжение
        grid.Add(wx.StaticText(self, label="Напряжение, В:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.voltage_value = wx.StaticText(self, label="0.00")
        grid.Add(self.voltage_value, 0, wx.ALIGN_CENTER)
        grid.Add(wx.StaticText(self, label="≥ 5000 В"), 0, wx.ALIGN_CENTER)

        # Сопротивление
        grid.Add(wx.StaticText(self, label="Сопротивление, МОм:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.resistance_value = wx.StaticText(self, label="0.00")
        grid.Add(self.resistance_value, 0, wx.ALIGN_CENTER)
        grid.Add(wx.StaticText(self, label="≥ 100 МОм"), 0, wx.ALIGN_CENTER)

        # Индекс поляризации
        grid.Add(wx.StaticText(self, label="Индекс поляризации:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.polarization_value = wx.StaticText(self, label="0.00")
        grid.Add(self.polarization_value, 0, wx.ALIGN_CENTER)
        grid.Add(wx.StaticText(self, label="≥ 2.0"), 0, wx.ALIGN_CENTER)

        params_sizer.Add(grid, 0, wx.EXPAND | wx.ALL, 10)
        sizer.Add(params_sizer, 0, wx.EXPAND | wx.ALL, 10)

        # Чекбоксы обкатки
        run_box = wx.StaticBox(self, label="Тип обкатки")
        run_sizer = wx.StaticBoxSizer(run_box, wx.HORIZONTAL)

        self.first_run = wx.CheckBox(self, label="Первая обкатка")
        self.second_run = wx.CheckBox(self, label="Вторая обкатка")

        run_sizer.Add(self.first_run, 0, wx.ALL, 10)
        run_sizer.Add(self.second_run, 0, wx.ALL, 10)

        sizer.Add(run_sizer, 0, wx.EXPAND | wx.ALL, 10)

        # Привязка событий для чекбоксов
        self.first_run.Bind(wx.EVT_CHECKBOX, self.on_checkbox)
        self.second_run.Bind(wx.EVT_CHECKBOX, self.on_checkbox)

        # Кнопки управления
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.btn_start = wx.Button(self, label="Начать измерения")
        self.btn_stop = wx.Button(self, label="Стоп")
        self.btn_save = wx.Button(self, label="Сохранить")

        btn_sizer.Add(self.btn_start, 0, wx.ALL, 5)
        btn_sizer.Add(self.btn_stop, 0, wx.ALL, 5)
        btn_sizer.AddStretchSpacer()
        btn_sizer.Add(self.btn_save, 0, wx.ALL, 5)

        sizer.Add(btn_sizer, 0, wx.EXPAND | wx.ALL, 10)

        # Журнал событий
        log_box = wx.StaticBox(self, label="Ход выполнения")
        log_sizer = wx.StaticBoxSizer(log_box, wx.VERTICAL)
        self.log_text = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.TE_READONLY)
        log_sizer.Add(self.log_text, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(log_sizer, 1, wx.EXPAND | wx.ALL, 10)

        self.SetSizer(sizer)
        self.Centre()

        # Привязка событий
        self.btn_start.Bind(wx.EVT_BUTTON, self.on_start)
        self.btn_stop.Bind(wx.EVT_BUTTON, self.on_stop)
        self.btn_save.Bind(wx.EVT_BUTTON, self.on_save)
        self.Bind(wx.EVT_CLOSE, self.on_close)

        # Настройка состояния кнопок
        self.btn_stop.Disable()

        # Стартовое сообщение
        self.log_message(f"Инициализация теста: {title}")

    def on_checkbox(self, event):
        # Гарантируем, что выбрана только одна обкатка
        if event.GetEventObject() == self.first_run and self.first_run.GetValue():
            self.second_run.SetValue(False)
        elif event.GetEventObject() == self.second_run and self.second_run.GetValue():
            self.first_run.SetValue(False)

    def log_message(self, message):
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        self.log_text.AppendText(f"{timestamp} - {message}\n")

    def on_start(self, event):  # noqa: unused-argument
        if not (self.first_run.GetValue() or self.second_run.GetValue()):
            wx.MessageBox("Выберите тип обкатки!", "Ошибка", wx.OK | wx.ICON_WARNING)
            return

        self.log_message("Начинаем измерения...")
        self.btn_start.Disable()
        self.btn_stop.Enable()
        self.running = True
        self.measurement_stop = False

        # Запускаем измерения в отдельном потоке
        self.measurement_thread = threading.Thread(target=self.run_measurements)
        self.measurement_thread.daemon = True
        self.measurement_thread.start()

    def run_measurements(self):
        try:
            # Инициализация COM-порта (указать правильный порт)
            self.ser = serial.Serial(
                port='COM7',  # Замените на актуальный порт
                baudrate=9600,  # Скорость обмена
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                timeout=1  # Таймаут чтения (секунды)
            )
            self.log_message("COM-порт открыт")
        except Exception as e:
            self.log_message(f"Ошибка открытия COM-порта: {str(e)}")
            wx.CallAfter(self.on_measurement_error, f"Ошибка открытия COM-порта: {str(e)}")
            return

        # Шаг 1: Удаленное управление
        if not self.send_command("Rn", "OK"):
            self.cleanup()
            return

        # Шаг 2: Включение прибора
        if not self.send_command("Bd", "OK"):
            self.cleanup()
            return

        # Шаг 3: Ожидание 1 минута
        self.log_message("Ожидание 1 минута...")
        for i in range(10): # 60
            if self.measurement_stop:
                break
            time.sleep(1)
            wx.CallAfter(self.log_message, f"Осталось: {20 - i} секунд") # 59 s

        # Шаг 4: Отключение прибора
        if not self.send_command("Bu", "OK"):
            self.cleanup()
            return

        if not self.send_command("Bu", "OK"):
            self.cleanup()
            return

        # Шаг 5: Настройка прибора (Df0040)
        if not self.send_command("Df0040", "OK"):
            self.cleanup()
            return

        # Шаг 6: Запрос показаний (Resistance Isolation)
        resistance = self.get_measurement("Dg")
        if resistance is None:
            self.cleanup()
            return

        wx.CallAfter(self.resistance_value.SetLabel, f"{resistance:.2f}")
        self.log_message(f"Сопротивление изоляции: {resistance:.2f} МОм")

        # Шаг 7: Настройка прибора (Df0100)
        if not self.send_command("Df0100", "OK"):
            self.cleanup()
            return

        # Шаг 8: Запрос показаний (Voltage)
        voltage = self.get_measurement("Dg")
        if voltage is None:
            self.cleanup()
            return

        wx.CallAfter(self.voltage_value.SetLabel, f"{voltage:.2f}")
        self.log_message(f"Напряжение: {voltage:.2f} В")

        # Расчет индекса поляризации (примерная логика)
        polarization_index = resistance / (voltage / 1000) if voltage > 0 else 0
        wx.CallAfter(self.polarization_value.SetLabel, f"{polarization_index:.2f}")
        self.log_message(f"Индекс поляризации: {polarization_index:.2f}")

        # Проверка результатов
        self.check_results(resistance, voltage, polarization_index)

        self.log_message("Измерения завершены")
        self.cleanup()
        wx.CallAfter(self.btn_stop.Disable)
        wx.CallAfter(self.btn_start.Enable)

    def send_command(self, command, expected_response, max_retries=5):
        """Отправка команды прибору с проверкой ответа"""
        for attempt in range(max_retries):
            if self.measurement_stop:
                return False

            try:
                self.ser.write(f"{command}\r\n".encode('ascii'))
                self.log_message(f"Отправлено: {command}")
                print(self.ser.write(f"{command}\r\n".encode('ascii')))
                print(f"Отправлено: {command}")

                # Чтение ответа
                response = b''
                start_time = time.time()
                while time.time() - start_time < 5:  # Таймаут 2 секунды
                    if self.ser.in_waiting > 0:
                        response += self.ser.read(self.ser.in_waiting)
                        if response.endswith(b'\r\n'):
                            break
                    time.sleep(0.1)

                response_str = response.decode('ascii', errors='ignore').strip()
                self.log_message(f"Получено: {response_str}")

                if expected_response in response_str:
                    return True

            except Exception as e:
                self.log_message(f"Ошибка при отправке команды: {str(e)}")

            self.log_message(f"Повторная попытка ({attempt + 1}/{max_retries})...")
            time.sleep(0.5)

        wx.CallAfter(self.on_measurement_error, f"Не получен ожидаемый ответ на команду: {command}")
        return False

    def get_measurement(self, command, max_retries=3):
        """Получение числового значения измерения от прибора"""
        for attempt in range(max_retries):
            if self.measurement_stop:
                return None

            try:
                self.ser.write(f"{command}\r\n".encode('ascii'))
                self.log_message(f"Отправлено: {command}")

                # Чтение ответа
                response = b''
                start_time = time.time()
                while time.time() - start_time < 2:  # Таймаут 2 секунды
                    if self.ser.in_waiting > 0:
                        response += self.ser.read(self.ser.in_waiting)
                        if response.endswith(b'\r'):
                            break
                    time.sleep(0.1)

                response_str = response.decode('ascii', errors='ignore').strip()
                self.log_message(f"Получено: {response_str}")

                # Поиск числового значения в ответе
                match = re.search(r'[-+]?\d*\.\d+|\d+', response_str)
                if match:
                    return float(match.group())

            except Exception as e:
                self.log_message(f"Ошибка при получении измерения: {str(e)}")

            self.log_message(f"Повторная попытка ({attempt + 1}/{max_retries})...")
            time.sleep(0.5)

        wx.CallAfter(self.on_measurement_error, "Не удалось получить значение измерения")
        return None

    @staticmethod
    def check_results(resistance, voltage, polarization_index):
        """Проверка результатов на соответствие нормативам"""
        messages = []
        if resistance < 100:
            messages.append(f"Сопротивление изоляции {resistance:.2f} МОм < 100 МОм")
        if voltage < 5000:
            messages.append(f"Напряжение {voltage:.2f} В < 5000 В")
        if polarization_index < 2.0:
            messages.append(f"Индекс поляризации {polarization_index:.2f} < 2.0")

        if messages:
            message = "Обнаружены проблемы:\n" + "\n".join(messages)
            wx.CallAfter(lambda: (wx.MessageBox, message, "Результаты проверки", wx.OK | wx.ICON_WARNING))
        else:
            wx.CallAfter(lambda: wx.MessageBox("Все параметры в норме!", "Результаты проверки",
                                               wx.OK | wx.ICON_INFORMATION))

    def on_measurement_error(self, message):
        self.log_message(f"Ошибка: {message}")
        wx.MessageBox(message, "Ошибка измерения", wx.OK | wx.ICON_ERROR)
        self.btn_stop.Disable()
        self.btn_start.Enable()

    def cleanup(self):
        """Очистка ресурсов после измерений"""
        if self.ser and self.ser.is_open:
            try:
                self.ser.close()
                self.log_message("COM-порт закрыт")
            except Exception as e:
                logging.error(f"Необработанное исключение: {e}")
        self.running = False

    def on_stop(self, event):  # noqa: unused-argument
        if not self.running:
            return

        self.measurement_stop = True
        self.log_message("Остановка измерений...")

        # Попытка отправить команду отключения
        if self.ser and self.ser.is_open:
            try:
                self.ser.write(b"Bu\r\n")
                self.log_message("Отправлена команда отключения")
            except Exception as e:
                logging.error(f"Необработанное исключение: {e}")

        wx.MessageBox("Остановлено пользователем!", "Внимание", wx.OK | wx.ICON_WARNING)
        self.btn_stop.Disable()
        self.btn_start.Enable()

    def on_save(self, event):  # noqa: unused-argument
        # Здесь должна быть логика сохранения результатов
        wx.MessageBox("Результаты измерений сохранены", "Сохранение", wx.OK | wx.ICON_INFORMATION)

    def on_close(self, event):  # noqa: unused-argument
        if self.running:
            self.on_stop(None)
        self.cleanup()
        if self.parent:
            self.parent.Enable(True)
        self.Destroy()


class TestDialog(wx.Dialog):
    def __init__(self, parent, title, description):
        super().__init__(parent, title=title, size=wx.Size(600, 400))
        self.SetBackgroundColour(wx.Colour(245, 248, 252))  # Светлый фон
        self.parent = parent

        sizer = wx.BoxSizer(wx.VERTICAL)

        # Заголовок теста
        title_label = wx.StaticText(self, label=title)
        title_label.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        title_label.SetForegroundColour(wx.Colour(0, 50, 100))  # Темно-синий цвет
        sizer.Add(title_label, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        # Описание теста
        desc_label = wx.StaticText(self, label=description)
        desc_label.Wrap(500)
        sizer.Add(desc_label, 0, wx.ALL | wx.EXPAND, 10)

        # Параметры теста
        params_box = wx.StaticBox(self, label="Параметры тестирования")
        params_sizer = wx.StaticBoxSizer(params_box, wx.VERTICAL)
        grid = wx.FlexGridSizer(cols=2, vgap=5, hgap=10)
        grid.AddGrowableCol(1)

        # Пример параметров (можно настроить для каждого теста отдельно)
        params = [
            ("Напряжение, В:", "380"),
            ("Длительность, сек:", "60"),
            ("Порог срабатывания:", "5.0"),
            ("Режим работы:", "Автоматический")
        ]

        self.controls = []
        for label, default in params:
            grid.Add(wx.StaticText(self, label=label), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
            ctrl = wx.TextCtrl(self, value=default)
            grid.Add(ctrl, 1, wx.EXPAND)
            self.controls.append(ctrl)

        params_sizer.Add(grid, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(params_sizer, 0, wx.EXPAND | wx.ALL, 10)

        # Кнопки управления
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.btn_start = wx.Button(self, label="Начать тест")
        self.btn_stop = wx.Button(self, label="Остановить")
        self.btn_save = wx.Button(self, label="Сохранить результаты")

        # Стилизуем кнопки
        for btn in [self.btn_start, self.btn_stop, self.btn_save]:
            btn.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
            btn.SetBackgroundColour(wx.Colour(70, 130, 180))  # Приятный синий
            btn.SetForegroundColour(wx.WHITE)

        btn_sizer.Add(self.btn_start, 0, wx.ALL, 5)
        btn_sizer.Add(self.btn_stop, 0, wx.ALL, 5)
        btn_sizer.AddStretchSpacer()
        btn_sizer.Add(self.btn_save, 0, wx.ALL, 5)

        sizer.Add(btn_sizer, 0, wx.EXPAND | wx.ALL, 10)

        # Журнал событий
        log_box = wx.StaticBox(self, label="Ход выполнения")
        log_sizer = wx.StaticBoxSizer(log_box, wx.VERTICAL)
        self.log_text = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.TE_READONLY)
        log_sizer.Add(self.log_text, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(log_sizer, 1, wx.EXPAND | wx.ALL, 10)

        self.SetSizer(sizer)
        self.Centre()

        # Привязка событий
        self.btn_start.Bind(wx.EVT_BUTTON, self.on_start)
        self.btn_stop.Bind(wx.EVT_BUTTON, self.on_stop)
        self.btn_save.Bind(wx.EVT_BUTTON, self.on_save)
        self.Bind(wx.EVT_CLOSE, self.on_close)

        # Стартовое сообщение
        self.log_message(f"Инициализация теста: {title}")

    def log_message(self, message):
        timestamp = time.strftime("%H:%M:%S", time.localtime())
        self.log_text.AppendText(f"{timestamp} - {message}\n")

    def on_start(self, event):  # noqa: unused-argument
        self.log_message("Тестирование запущено")
        self.btn_start.Disable()
        self.btn_stop.Enable()

    def on_stop(self, event):  # noqa: unused-argument
        self.log_message("Тестирование остановлено")
        self.btn_start.Enable()
        self.btn_stop.Disable()

    def on_save(self, event):  # noqa: unused-argument
        self.log_message("Результаты сохранены")
        wx.MessageBox("Результаты теста успешно сохранены", "Сохранение", wx.OK | wx.ICON_INFORMATION)

    def on_close(self, event):  # noqa: unused-argument
        # Разблокируем основное окно при закрытии диалога
        if self.parent:
            self.parent.Enable(True)
        self.Destroy()


class SplashScreen(wx.Frame):
    def __init__(self, on_close_callback=None):
        # Создаем полноэкранное окно с черным фоном
        style = wx.STAY_ON_TOP | wx.FRAME_NO_TASKBAR | wx.BORDER_NONE
        super().__init__(None, style=style)

        # Получаем размеры экрана
        screen_width, screen_height = wx.GetDisplaySize()
        self.SetSize(wx.Size(screen_width, screen_height))
        self.SetBackgroundColour(wx.BLACK)
        self.Centre()

        # Создаем панель для надписи
        self.panel = wx.Panel(self)
        self.panel.SetBackgroundColour(wx.BLACK)

        # Изменяем цвет текста заставки
        self.text = wx.StaticText(self.panel, label='ООО "РИНПО"\nСлужба\nАвтоматизации\nПроизводственных\nПроцессов')
        self.text.SetForegroundColour(wx.Colour(220, 230, 255))  # Светло-голубой
        self.text.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))

        # Центрируем надпись
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.AddStretchSpacer()
        sizer.Add(self.text, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        sizer.AddStretchSpacer()
        self.panel.SetSizer(sizer)

        # Переменные для анимации
        self.start_time = time.time()
        self.animation_duration = 5.0  # 5 секунд
        self.min_font_size = 1
        self.max_font_size = 72

        # Таймер для анимации
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.on_timer)
        self.timer.Start(90)  # Обновление каждые 30 мс

        self.ShowFullScreen(True)
        self.Show()

        self.on_close_callback = on_close_callback

    def on_timer(self, event):  # noqa: unused-argument
        current_time = time.time()
        elapsed = current_time - self.start_time

        if elapsed >= self.animation_duration:
            self.timer.Stop()
            if self.on_close_callback:
                wx.CallAfter(self.on_close_callback)
            self.Destroy()
            return

        # Рассчитываем прогресс (0.0 - 1.0)
        progress = elapsed / self.animation_duration

        # Плавное увеличение размера шрифта
        font_size = self.min_font_size + (self.max_font_size - self.min_font_size) * self.ease_out(progress)
        self.text.SetFont(wx.Font(int(font_size), wx.FONTFAMILY_DEFAULT,
                                  wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))

        # Центрируем текст
        self.panel.Layout()

    @staticmethod
    def ease_out(t):
        """Функция плавности для более естественной анимации"""
        return 1 - math.pow(1 - t, 3)


class CombinedTab(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent)
        self.SetBackgroundColour(wx.Colour(240, 245, 250))  # Основной фон
        sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Левая панель: Выбор модели ЭД
        self.left_panel = ModelSelectorPanel(self)
        sizer.Add(self.left_panel, 1, wx.EXPAND | wx.ALL, 5)

        # Центральная панель: Параметры выбранного ЭД
        self.center_panel = EDParametersPanel(self)
        sizer.Add(self.center_panel, 1, wx.EXPAND | wx.ALL, 5)

        # Правая панель: Параметры протокола
        self.right_panel = wx.Panel(self)
        right_sizer = wx.BoxSizer(wx.VERTICAL)

        # Изменяем стиль заголовка в правой панели
        title = wx.StaticText(self.right_panel, label="Параметры протокола")
        title.SetFont(wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        title.SetForegroundColour(wx.Colour(0, 50, 100))  # Темно-синий цвет
        right_sizer.Add(title, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        # Форма протокола
        form_box = wx.StaticBox(self.right_panel, label="Исходные данные")
        form_sizer = wx.StaticBoxSizer(form_box, wx.VERTICAL)

        # Создаем панель с прокруткой для формы
        form_scrolled = wx.ScrolledWindow(self.right_panel, style=wx.VSCROLL)
        form_scrolled.SetScrollRate(0, 20)
        form_grid = wx.FlexGridSizer(cols=2, vgap=8, hgap=10)
        form_grid.AddGrowableCol(1)

        # Поля формы
        fields = [
            "Номер протокола:", "Группа исполнения:", "Номер ЭД:",
            "Категория исполнения:", "Тип масла:", "Тип муфты:",
            "Дата и время испытаний:", "Проверка маркировки:",
            "Радиальное биение шлицевого\nконца относительно\nцентрирующей поверхности\n= 16мм:",
            "Вылет вала:", "Тип РТИ:", "Диаметр вала, мм:", "Тип ТМС / конденсатора:",
            "Длительность обкатки:", "Сочленение шлицевых соединений:",
            "Тип блока ТМС:", "Версия прошивки:", "Номер вала:",
            "Предел текучести:", "Оператор:"
        ]

        # Создаем элементы управления
        self.search_controls = []
        for field in fields:
            if field == "Дата и время испытаний:":
                ctrl = wx.adv.DatePickerCtrl(form_scrolled)
            else:
                ctrl = wx.TextCtrl(form_scrolled)
            self.search_controls.append(ctrl)

        # Добавляем в сетку
        for i, field in enumerate(fields):
            label = wx.StaticText(form_scrolled, label=field)
            form_grid.Add(label, 0, wx.ALIGN_CENTER_VERTICAL | wx.LEFT, 5)
            form_grid.Add(self.search_controls[i], 1, wx.EXPAND | wx.RIGHT, 5)

        form_scrolled.SetSizer(form_grid)
        form_sizer.Add(form_scrolled, 1, wx.EXPAND | wx.ALL, 5)
        right_sizer.Add(form_sizer, 1, wx.EXPAND | wx.ALL, 5)

        # Кнопки
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.btn_search = wx.Button(self.right_panel, label="Редактировать")
        self.btn_clear = wx.Button(self.right_panel, label="Очистить форму")
        self.btn_export = wx.Button(self.right_panel, label="Сохранить")

        # Стилизуем кнопки
        for btn in [self.btn_search, self.btn_clear, self.btn_export]:
            btn.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
            btn.SetBackgroundColour(wx.Colour(70, 130, 180))  # Приятный синий
            btn.SetForegroundColour(wx.WHITE)

        btn_sizer.Add(self.btn_search, 0, wx.ALL, 5)
        btn_sizer.Add(self.btn_clear, 0, wx.ALL, 5)
        btn_sizer.Add(self.btn_export, 0, wx.ALL, 5)

        right_sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.TOP, 10)

        self.right_panel.SetSizer(right_sizer)
        sizer.Add(self.right_panel, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(sizer)

        # Привязка событий
        self.btn_search.Bind(wx.EVT_BUTTON, self.on_search_protocols)
        self.btn_clear.Bind(wx.EVT_BUTTON, self.on_clear_search)
        self.btn_export.Bind(wx.EVT_BUTTON, self.on_export)

    def set_selected_model(self, params):
        """Обработчик выбора модели из левой панели"""
        # Передаем параметры в центральную панель
        self.center_panel.set_parameters(params)

    def on_search_protocols(self, event):  # noqa: unused-argument
        """Поиск протоколов по критериям"""
        wx.MessageBox("Поиск протоколов выполнен", "Результат", wx.OK | wx.ICON_INFORMATION)

    def on_clear_search(self, event):  # noqa: unused-argument
        """Очистка формы поиска"""
        for ctrl in self.search_controls:
            if isinstance(ctrl, wx.adv.DatePickerCtrl):
                ctrl.SetValue(wx.DateTime.Now())
            else:
                ctrl.SetValue("")
        wx.MessageBox("Форма очищена", "Очистка", wx.OK | wx.ICON_INFORMATION)

    def on_export(self, event):  # noqa: unused-argument
        """Экспорт результатов поиска в Excel"""
        try:
            # Собираем данные из всех полей
            field_values = []
            field_names = []

            # Собираем данные из правой панели (форма поиска)
            fields_right = [
                "Номер протокола", "Группа исполнения", "Номер ЭД",
                "Категория исполнения", "Тип масла", "Тип муфты",
                "Дата и время испытаний", "Проект 3,3/1",
                "Реализация боевых шпилеков", "Балет вала", "Тип РТИ",
                "Данные вала, мм", "Тип ТИС / конденсатора",
                "Длительность объекта", "Сочленение шпилеков соединений",
                "Тип блока ТИС", "Версия прошивки", "Номер вала",
                "Предел текучести", "Оператор"
            ]

            for i, field in enumerate(fields_right):
                ctrl = self.search_controls[i]
                if isinstance(ctrl, wx.adv.DatePickerCtrl):
                    value = ctrl.GetValue().FormatISODate() if ctrl.GetValue().IsValid() else ""
                else:
                    value = ctrl.GetValue().strip()
                field_values.append(value)
                field_names.append(field)

            # Собираем данные из центральной панели (параметры ЭД)
            model_value = self.center_panel.param_controls["Model"].GetValue()
            field_values.append(model_value)
            field_names.append("Модель")

            # Проверяем заполненность ВСЕХ полей
            missing_fields = []
            for i, value in enumerate(field_values):
                if not value:
                    missing_fields.append(field_names[i])

            if missing_fields:
                wx.MessageBox(
                    f"Следующие поля не заполнены:\n{', '.join(missing_fields)}\n\n"
                    "Пожалуйста, заполните все поля перед экспортом.",
                    "Незаполненные поля", wx.OK | wx.ICON_WARNING
                )
                return

            # Извлекаем ключевые значения
            protocol_number = field_values[0]  # Номер протокола
            execution_group = field_values[1]  # Группа исполнения
            ed_number = field_values[2]  # Номер ЭД
            model = field_values[-1]  # Модель

            # Проверяем допустимость символов в имени файла
            invalid_chars = r'[\\/*?:"<>|]'
            if re.search(invalid_chars, protocol_number) or re.search(invalid_chars, ed_number) or re.search(
                    invalid_chars, model):
                wx.MessageBox(
                    "Номер протокола, номер ЭД и модель не должны содержать следующие символы:\n"
                    "\\ / : * ? \" < > |",
                    "Недопустимые символы", wx.OK | wx.ICON_WARNING
                )
                return

            # Формируем имя файла
            file_name = f"{protocol_number}{ed_number}{model}.xlsx"
            template_path = r"C:\pattern\pattern.xlsx"
            export_path = os.path.join(r"C:\dumpProtocols", file_name)

            # Проверяем существование шаблона
            if not os.path.exists(template_path):
                wx.MessageBox(
                    f"Шаблон протокола не найден по пути:\n{template_path}",
                    "Ошибка экспорта", wx.OK | wx.ICON_ERROR
                )
                return

            # Создаем директорию для экспорта, если нужно
            try:
                os.makedirs(os.path.dirname(export_path), exist_ok=True)
            except OSError as e:
                wx.MessageBox(
                    f"Ошибка создания директории для экспорта:\n{str(e)}",
                    "Ошибка экспорта", wx.OK | wx.ICON_ERROR
                )
                return

            # Копируем и заполняем шаблон
            try:
                # Загружаем шаблон
                wb = openpyxl.load_workbook(template_path)
                sheet = wb.active

                # Записываем данные в ячейки
                sheet['L4'] = protocol_number
                sheet['F7'] = execution_group

                # Сохраняем файл
                wb.save(export_path)

                # Проверяем, что файл действительно создан
                if not os.path.exists(export_path):
                    raise RuntimeError(f"Файл не был создан: {export_path}")

                # Деактивируем все поля ввода
                self.disable_all_controls()

                # Выводим сообщение об успехе
                dlg = wx.MessageDialog(
                    self,
                    f"Протокол успешно создан:\n{file_name}\n\n"
                    f"Путь: {export_path}",
                    "Экспорт завершен",
                    wx.OK | wx.ICON_INFORMATION
                )
                dlg.ShowModal()
                dlg.Destroy()

            except (InvalidFileException, PermissionError, OSError, RuntimeError) as e:
                wx.MessageBox(
                    f"Ошибка при создании файла протокола:\n{str(e)}",
                    "Ошибка экспорта", wx.OK | wx.ICON_ERROR
                )
            except Exception as e:
                wx.MessageBox(
                    f"Неизвестная ошибка при создании протокола:\n{str(e)}\n\n"
                    f"Детали:\n{traceback.format_exc()}",
                    "Ошибка", wx.OK | wx.ICON_ERROR
                )

        except Exception as e:
            wx.MessageBox(
                f"Критическая ошибка при экспорте:\n{str(e)}\n\n"
                f"Детали:\n{traceback.format_exc()}",
                "Ошибка", wx.OK | wx.ICON_ERROR
            )

    def disable_all_controls(self):
        """Деактивирует все элементы управления на вкладке"""
        # Деактивируем элементы в левой панели
        self.left_panel.search_text.Disable()
        self.left_panel.search_btn.Disable()
        self.left_panel.select_btn.Disable()
        self.left_panel.models_list.Disable()

        # Деактивируем элементы в центральной панели
        for ctrl in self.center_panel.param_controls.values():
            ctrl.Disable()
        self.center_panel.save_btn.Disable()

        # Деактивируем элементы в правой панели
        for ctrl in self.search_controls:
            ctrl.Disable()

        self.btn_search.Disable()
        self.btn_clear.Disable()
        self.btn_export.Disable()


class PEDTestingApp(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, id=wx.ID_ANY, title=title, size=wx.Size(1500, 1000))

        # Устанавливаем основной цвет фона
        self.SetBackgroundColour(wx.Colour(240, 245, 250))

        self.test_buttons = None
        self.btn_export = None
        self.btn_clear = None
        self.btn_search = None
        self.left_panel = None
        self.center_panel = None
        self.search_controls = None
        # Создаем панель и основной сайзер
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Создаем Notebook (вкладки)
        self.notebook = wx.Notebook(panel)
        main_sizer.Add(self.notebook, 1, wx.EXPAND | wx.ALL, 5)

        # Создаем вкладки
        self.setup_combined_tab()
        self.setup_connection_tab()
        self.setup_testing_tab()
        self.setup_database_tab()

        panel.SetSizer(main_sizer)
        self.Centre()

    def setup_database_tab(self):
        """Вкладка работы с базой данных"""
        tab = DatabaseTab(self.notebook)
        self.notebook.AddPage(tab, "База данных ПЭД")

    def setup_connection_tab(self):
        """Вкладка настройки подключений"""
        tab = wx.Panel(self.notebook)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Группа настроек подключения
        group = wx.StaticBox(tab, label="Настройка подключений")
        group_sizer = wx.StaticBoxSizer(group, wx.VERTICAL)
        grid = wx.FlexGridSizer(cols=2, vgap=5, hgap=5)
        grid.AddGrowableCol(1)

        # Поля ввода
        fields = [
            ("Тип подключения:", wx.TextCtrl(tab)),
            ("IP-адрес:", wx.TextCtrl(tab)),
            ("Порт:", wx.TextCtrl(tab)),
            ("Таймаут, сек:", wx.TextCtrl(tab)),
            ("Пользователь:", wx.TextCtrl(tab)),
            ("Пароль:", wx.TextCtrl(tab, style=wx.TE_PASSWORD))
        ]

        for label, ctrl in fields:
            grid.Add(wx.StaticText(tab, label=label), 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 5)
            grid.Add(ctrl, 1, wx.EXPAND | wx.ALL, 5)

        group_sizer.Add(grid, 1, wx.EXPAND | wx.ALL, 5)
        sizer.Add(group_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # Кнопки
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_connect = wx.Button(tab, label="Подключиться")
        btn_disconnect = wx.Button(tab, label="Отключиться")
        btn_test = wx.Button(tab, label="Проверить соединение")

        btn_sizer.Add(btn_connect, 0, wx.ALL, 5)
        btn_sizer.Add(btn_disconnect, 0, wx.ALL, 5)
        btn_sizer.Add(btn_test, 0, wx.ALL, 5)

        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.TOP, 10)
        sizer.AddStretchSpacer()

        tab.SetSizer(sizer)
        self.notebook.AddPage(tab, "Настройка подключений")

    def setup_testing_tab(self):
        """Вкладка тестирования ПЭД"""
        tab = wx.Panel(self.notebook)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Список типов ПЭД с короткими названиями для кнопок
        ped_tests = [
            ("1. Пробой масла (хол.)", "Намерение напряжения пробоя масла холодного ПЭД"),
            ("2. Сопр. обмоток (хол.)", "Намерение сопротивления обмоток холодного ПЭД"),
            ("3. Сопр. вводов (хол.)", "Намерение сопротивления вводящим холодного ПЭД"),
            ("4. Подержание вала", "Намерение попечата продерживания вала"),
            ("5. КЗ трансформатора", "Опыт короткого замыкания"),
            ("6. Напряжение травма", "Определение напряжения травма"),
            ("7. Обязательная проверка", "Обязательная ПЭД"),
            ("8. Выбег ПЭД", "Измерение выбега ПЭД"),
            ("9. Сопр. обмоток (гор.)", "Намерение сопротивления обмоток горячего ПЭД"),
            ("10. Изоляция (гор.)", "Измерение сопротивления изоляции горячего ПЭД"),
            ("11. Пробой масла (гор.)", "Измерение напряжения пробоя масла горячего ПЭД"),
            ("12. Обмотки на эл.проч.", "Ипы также нежелте, но они обмотки на эл. прочих (4,9/10)"),
            ("13. Обмотки-корпус", "Ипы также кое-где обмотки на эл. прочих отл. корпуса"),
            ("14. Герметичность", "Проект 3,3/1 на герметичность маслом"),
            ("15. Показания ТИС", "Показания ТИС")
        ]

        # Создаем кнопки для каждого теста в три ряда
        self.test_buttons = []
        grid_sizer = wx.GridSizer(5, 3, 100, 100)  # 3 ряда, 5 столбцов, отступы 5 пикселей

        for short_name, full_name in ped_tests:
            btn = wx.Button(tab, label=short_name, size=wx.Size(200, 350))
            btn.SetToolTip(full_name)  # Полное название в подсказке

            # Настройка внешнего вида кнопки
            btn.SetBackgroundColour(wx.BLACK)
            btn.SetForegroundColour(wx.WHITE)
            font = wx.Font(12, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
            btn.SetFont(font)

            # Обработка наведения курсора
            btn.Bind(wx.EVT_ENTER_WINDOW, self.on_enter_button)
            btn.Bind(wx.EVT_LEAVE_WINDOW, self.on_leave_button)

            # Для кнопки 3 используем специальный диалог
            if short_name.startswith("3."):
                btn.Bind(wx.EVT_BUTTON, lambda event, sn=short_name,
                         fn=full_name: self.open_cold_input_dialog(event, sn, fn))
            else:
                btn.Bind(wx.EVT_BUTTON, lambda event, sn=short_name,
                         fn=full_name: self.open_test_dialog(event, sn, fn))

            self.test_buttons.append(btn)
            grid_sizer.Add(btn, 0, wx.EXPAND)

        # Добавляем grid_sizer в основной sizer
        sizer.Add(grid_sizer, 0, wx.EXPAND | wx.ALL, 5)

        tab.SetSizer(sizer)
        self.notebook.AddPage(tab, "Тестирование ПЭД")

    @staticmethod
    def on_enter_button(event):
        """Обработчик наведения курсора на кнопку"""
        button = event.GetEventObject()
        button.SetBackgroundColour(wx.Colour(50, 50, 50))  # Темно-серый цвет при наведении
        button.Refresh()

    @staticmethod
    def on_leave_button(event):
        """Обработчик ухода курсора с кнопки"""
        button = event.GetEventObject()
        button.SetBackgroundColour(wx.BLACK)  # Возвращаем черный цвет
        button.Refresh()

    def open_cold_input_dialog(self, event, title, description):  # noqa: unused-argument
        """Открывает специальный диалог для измерения сопротивления вводов"""
        self.Enable(False)
        dlg = ColdInputResistanceDialog(self, title, description)
        dlg.Show()

    def open_test_dialog(self, event, title, description):  # noqa: unused-argument
        """Открывает диалоговое окно для конкретного теста"""
        self.Enable(False)
        dlg = TestDialog(self, title, description)
        dlg.Show()

    def on_export(self, event):  # noqa: unused-argument
        """Экспорт результатов поиска в Excel"""
        try:
            # Собираем данные из всех полей
            field_values = []
            field_names = []

            # Собираем данные из правой панели (форма поиска)
            fields_right = [
                "Номер протокола", "Группа исполнения", "Номер ЭД",
                "Категория исполнения", "Тип масла", "Тип муфты",
                "Дата и время испытаний", "Проект 3,3/1",
                "Реализация боевых шпилеков", "Балет вала", "Тип РТИ",
                "Данные вала, мм", "Тип ТИС / конденсатора",
                "Длительность объекта", "Сочленение шпилеков соединений",
                "Тип блока ТИС", "Версия прошивки", "Номер вала",
                "Предел текучести", "Оператор"
            ]

            for i, field in enumerate(fields_right):
                ctrl = self.search_controls[i]
                if isinstance(ctrl, wx.adv.DatePickerCtrl):
                    value = ctrl.GetValue().FormatISODate() if ctrl.GetValue().IsValid() else ""
                else:
                    value = ctrl.GetValue().strip()
                field_values.append(value)
                field_names.append(field)

            # Собираем данные из центральной панели (параметры ЭД)
            model_value = self.center_panel.param_controls["Model"].GetValue()
            field_values.append(model_value)
            field_names.append("Модель")

            # Проверяем заполненность ВСЕХ полей
            missing_fields = []
            for i, value in enumerate(field_values):
                if not value:
                    missing_fields.append(field_names[i])

            if missing_fields:
                wx.MessageBox(
                    f"Следующие поля не заполнены:\n{', '.join(missing_fields)}\n\n"
                    "Пожалуйста, заполните все поля перед экспортом.",
                    "Незаполненные поля", wx.OK | wx.ICON_WARNING
                )
                return

            # Извлекаем ключевые значения
            protocol_number = field_values[0]  # Номер протокола
            execution_group = field_values[1]  # Группа исполнения
            ed_number = field_values[2]  # Номер ЭД
            model = field_values[-1]  # Модель

            # Проверяем допустимость символов в имени файла
            invalid_chars = r'[\\/*?:"<>|]'
            if re.search(invalid_chars, protocol_number) or re.search(invalid_chars, ed_number) or re.search(
                    invalid_chars, model):
                wx.MessageBox(
                    "Номер протокола, номер ЭД и модель не должны содержать следующие символы:\n"
                    "\\ / : * ? \" < > |",
                    "Недопустимые символы", wx.OK | wx.ICON_WARNING
                )
                return

            # Формируем имя файла
            file_name = f"{protocol_number}{ed_number}{model}.xlsx"
            template_path = r"C:\pattern\pattern.xlsx"
            export_path = os.path.join(r"C:\dumpProtocols", file_name)

            # Проверяем существование шаблона
            if not os.path.exists(template_path):
                wx.MessageBox(
                    f"Шаблон протокола не найден по пути:\n{template_path}",
                    "Ошибка экспорта", wx.OK | wx.ICON_ERROR
                )
                return

            # Создаем директорию для экспорта, если нужно
            try:
                os.makedirs(os.path.dirname(export_path), exist_ok=True)
            except OSError as e:
                wx.MessageBox(
                    f"Ошибка создания директории для экспорта:\n{str(e)}",
                    "Ошибка экспорта", wx.OK | wx.ICON_ERROR
                )
                return

            # Копируем и заполняем шаблон
            try:
                # Загружаем шаблон
                wb = openpyxl.load_workbook(template_path)
                sheet = wb.active

                # Записываем данные в ячейки
                sheet['L4'] = protocol_number
                sheet['F7'] = execution_group

                # Сохраняем файл
                wb.save(export_path)

                # Проверяем, что файл действительно создан
                if not os.path.exists(export_path):
                    raise RuntimeError(f"Файл не был создан: {export_path}")

                # Деактивируем все поля ввода
                self.disable_all_controls()

                # Выводим сообщение об успехе
                dlg = wx.MessageDialog(
                    self,
                    f"Протокол успешно создан:\n{file_name}\n\n"
                    f"Путь: {export_path}",
                    "Экспорт завершен",
                    wx.OK | wx.ICON_INFORMATION
                )
                dlg.ShowModal()
                dlg.Destroy()

            except (InvalidFileException, PermissionError, OSError, RuntimeError) as e:
                wx.MessageBox(
                    f"Ошибка при создании файла протокола:\n{str(e)}",
                    "Ошибка экспорта", wx.OK | wx.ICON_ERROR
                )
            except Exception as e:
                wx.MessageBox(
                    f"Неизвестная ошибка при создании протокола:\n{str(e)}\n\n"
                    f"Детали:\n{traceback.format_exc()}",
                    "Ошибка", wx.OK | wx.ICON_ERROR
                )

        except Exception as e:
            wx.MessageBox(
                f"Критическая ошибка при экспорте:\n{str(e)}\n\n"
                f"Детали:\n{traceback.format_exc()}",
                "Ошибка", wx.OK | wx.ICON_ERROR
            )

    def disable_all_controls(self):
        """Деактивирует все элементы управления на вкладке"""
        # Деактивируем элементы в левой панели
        self.left_panel.search_text.Disable()
        self.left_panel.search_btn.Disable()
        self.left_panel.select_btn.Disable()
        self.left_panel.models_list.Disable()

        # Деактивируем элементы в центральной панели
        for ctrl in self.center_panel.param_controls.values():
            ctrl.Disable()
        self.center_panel.save_btn.Disable()

        # Деактивируем элементы в правой панели
        for ctrl in self.search_controls:
            ctrl.Disable()

        self.btn_search.Disable()
        self.btn_clear.Disable()
        self.btn_export.Disable()

    def setup_combined_tab(self):
        """Параметры ПЭД"""
        tab = CombinedTab(self.notebook)  # Используем новый класс
        self.notebook.AddPage(tab, "Параметры ПЭД")

    def set_selected_model(self, params):
        """Обработчик выбора модели из левой панели"""
        # Передаем параметры в центральную панель
        self.center_panel.set_parameters(params)

        # Автоматическое заполнение поля "Номер ЭД" в форме протокола
        model = params.get("Model", "")
        if model:
            # Поле "Номер ЭД" - третий элемент в search_controls
            if len(self.search_controls) > 2:
                self.search_controls[2].SetValue(model)

    def on_search_protocols(self, event):  # noqa: unused-argument
        """Поиск протоколов по критериям"""
        # Здесь будет реализация поиска
        wx.MessageBox("Поиск протоколов выполнен", "Результат", wx.OK | wx.ICON_INFORMATION)

    def on_clear_search(self, event):  # noqa: unused-argument
        """Очистка формы поиска"""
        for ctrl in self.search_controls:
            if isinstance(ctrl, wx.adv.DatePickerCtrl):
                ctrl.SetValue(wx.DateTime.Now())
            else:
                ctrl.SetValue("")
        wx.MessageBox("Форма очищена", "Очистка", wx.OK | wx.ICON_INFORMATION)


if __name__ == "__main__":
    app = wx.App()

    # Создаем главное окно, но пока скрыто
    main_frame = PEDTestingApp(None, "Программный комплекс тестирования ПЭД")
    main_frame.Hide()


    def after_splash():
        # Максимизируем и показываем главное окно
        main_frame.Maximize(True)
        main_frame.Show()


    # Создаем заставку, передаем колбэк
    splash = SplashScreen(after_splash)

    app.MainLoop()
