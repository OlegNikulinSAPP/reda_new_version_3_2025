import wx
import wx.lib.scrolledpanel as scrolled
import os


class MyFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="Приложение с параметрами ПЭД",
                         style=wx.DEFAULT_FRAME_STYLE, size=(1600, 900))

        # Устанавливаем минимальный размер окна
        self.SetMinSize((1400, 700))

        # Создаем Notebook (блокнот с вкладками)
        self.notebook = wx.Notebook(self)

        # Создаем четыре вкладки
        self.tab1 = wx.Panel(self.notebook)
        self.tab2 = wx.Panel(self.notebook)
        self.tab3 = wx.Panel(self.notebook)
        self.tab4 = wx.Panel(self.notebook)

        # Добавляем панели в блокнот с названиями вкладок
        self.notebook.AddPage(self.tab1, "Основная")
        self.notebook.AddPage(self.tab2, "Вторая")
        self.notebook.AddPage(self.tab3, "Третья")
        self.notebook.AddPage(self.tab4, "Четвертая")

        # На первой вкладке создаем четыре колонки
        self.create_four_columns_on_tab1()

        # На остальных вкладках просто добавляем заглушки
        self.add_stub_to_tab(self.tab2, "Содержимое второй вкладки")
        self.add_stub_to_tab(self.tab3, "Содержимое третьей вкладки")
        self.add_stub_to_tab(self.tab4, "Содержимое четвертой вкладки")

        # Центрируем и показываем окно
        self.Centre()
        self.Show()

    def create_four_columns_on_tab1(self):
        """Создает четыре колонки на первой вкладке"""
        # Создаем SplitterWindow для разделения на левую и правую части
        splitter = wx.SplitterWindow(self.tab1)

        # Левая панель (кнопки и элементы управления)
        left_panel = wx.Panel(splitter)
        left_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # 1. Левая колонка - кнопки
        buttons_panel = wx.Panel(left_panel)
        self.create_buttons_column(buttons_panel)

        # 2. Средняя колонка - TextCtrl и ListBox
        controls_panel = wx.Panel(left_panel)
        self.create_controls_column(controls_panel)

        left_sizer.Add(buttons_panel, 0, wx.EXPAND | wx.ALL, 3)
        left_sizer.Add(controls_panel, 0, wx.EXPAND | wx.ALL, 3)
        left_panel.SetSizer(left_sizer)

        # Правая панель (параметры и поиск протокола) с прокруткой
        right_panel = scrolled.ScrolledPanel(splitter)
        right_panel.SetupScrolling()
        right_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # 3. Правая колонка - таблица параметров
        params_panel = wx.Panel(right_panel)
        self.create_parameters_column(params_panel)

        # 4. Четвертая колонка - форма поиска протокола
        search_panel = wx.Panel(right_panel)
        self.create_search_protocol_column(search_panel)

        right_sizer.Add(params_panel, 1, wx.EXPAND | wx.ALL, 3)
        right_sizer.Add(search_panel, 0, wx.EXPAND | wx.ALL, 3)
        right_panel.SetSizer(right_sizer)

        # Разделяем окно
        splitter.SplitVertically(left_panel, right_panel, sashPosition=400)
        splitter.SetMinimumPaneSize(400)

        # Устанавливаем основной sizer для первой вкладки
        tab_sizer = wx.BoxSizer(wx.VERTICAL)
        tab_sizer.Add(splitter, 1, wx.EXPAND)
        self.tab1.SetSizer(tab_sizer)

    def add_stub_to_tab(self, tab, text):
        """Добавляет заглушку на вкладку"""
        sizer = wx.BoxSizer(wx.VERTICAL)
        label = wx.StaticText(tab, label=text)
        sizer.Add(label, 0, wx.ALL, 20)
        tab.SetSizer(sizer)

    def create_buttons_column(self, parent):
        """Создает колонку с кнопками, которые заполняют всю высоту"""
        # Используем BoxSizer, который растягивается на всю высоту
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Список названий кнопок с иконками
        button_data = [
            ("Загрузить", wx.ART_FILE_OPEN),
            ("Сохранить", wx.ART_FILE_SAVE),
            ("Рассчитать", wx.ART_EXECUTABLE_FILE),
            ("Очистить", wx.ART_DELETE),
            ("Экспорт", wx.ART_GO_DIR_UP),
            ("Импорт", wx.ART_GO_DOWN),
            ("Печать", wx.ART_PRINT),
            ("Настройки", wx.ART_LIST_VIEW),
            ("Справка", wx.ART_HELP),
            ("О программе", wx.ART_INFORMATION),
            ("Инструмент 1", wx.ART_EXECUTABLE_FILE),
            ("Инструмент 2", wx.ART_NORMAL_FILE),
            ("Измерение", wx.ART_FIND),
            ("Калибровка", wx.ART_EXECUTABLE_FILE),
            ("Поиск", wx.ART_FIND_AND_REPLACE)
        ]

        # Создаем кнопки с иконками и текстом
        self.buttons = []
        for text, art_id in button_data:
            # Создаем кнопку с текстом
            button = wx.Button(parent, label=text, size=(150, 40))

            # Добавляем иконку, если она доступна
            if art_id:
                bmp = wx.ArtProvider.GetBitmap(art_id, wx.ART_BUTTON, (16, 16))
                if bmp.IsOk():
                    button.SetBitmap(bmp, wx.LEFT)

            sizer.Add(button, 0, wx.EXPAND | wx.ALL, 2)
            self.buttons.append(button)

        # Добавляем растягивающийся спейсер, чтобы кнопки занимали всю высоту
        sizer.AddStretchSpacer(1)

        parent.SetSizer(sizer)

    def create_controls_column(self, parent):
        """Создает колонку с TextCtrl и ListBox, которые заполняют всю высоту"""
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Добавляем TextCtrl
        self.text_ctrl = wx.TextCtrl(parent, size=(180, 25))
        sizer.Add(self.text_ctrl, 0, wx.EXPAND | wx.ALL, 3)

        # Добавляем ListBox, который растягивается на всю доступную высоту
        self.list_box = wx.ListBox(parent, size=(180, -1))
        sizer.Add(self.list_box, 1, wx.EXPAND | wx.ALL, 3)

        parent.SetSizer(sizer)

    def create_parameters_column(self, parent):
        """Создает колонку с таблицей параметров, которая заполняет всю высоту"""
        # Список параметров
        parameters = [
            ("Напряжение трансформатора, В", ""),
            ("Номинальные параметры ПЭД", "header"),
            ("Тип", ""),
            ("Мощность, кВт", ""),
            ("Напряжение, В", ""),
            ("Ток, А", ""),
            ("Частота вращения", ""),
            ("Сопр. обмотки (Ом)", ""),
            ("Сопр. изоляции, МОм", ""),
            ("Напр. разгона, В", ""),
            ("Момент проворачивания", ""),
            ("Напр. к.з., В", ""),
            ("Ток к.з., А", ""),
            ("Потери к.з., кВт", ""),
            ("Ток х.х., А", ""),
            ("Напр. х.х., В", ""),
            ("Потери х.х., кВт", ""),
            ("Выбег, с", ""),
            ("Вибрация, мм/с", ""),
            ("Крутящий момент", ""),
            ("Потери в нагр. сост.", ""),
            ("Испыт. напр. изол., В", ""),
            ("Испыт. изол. обм., В", ""),
            ("Напр. опыта КЗ", "")
        ]

        # Создаем сетку для параметров
        grid = wx.FlexGridSizer(len(parameters), 2, 5, 5)
        grid.AddGrowableCol(1, 1)  # Разрешаем растягивание второй колонки

        # Создаем шрифт для заголовка
        bold_font = wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)

        # Добавляем параметры в сетку
        self.param_fields = {}
        for param, param_type in parameters:
            if param_type == "header":
                # Заголовок - занимает всю ширину
                header = wx.StaticText(parent, label=param)
                header.SetFont(bold_font)
                grid.Add(header, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 3)
                grid.Add((0, 0), 0)  # Пустая ячейка
            else:
                # Обычный параметр
                label = wx.StaticText(parent, label=param)
                text_ctrl = wx.TextCtrl(parent, size=(150, -1))
                self.param_fields[param] = text_ctrl

                grid.Add(label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 3)
                grid.Add(text_ctrl, 0, wx.EXPAND | wx.ALL, 3)

        # Устанавливаем сетку на панель
        parent.SetSizer(grid)

    def create_search_protocol_column(self, parent):
        """Создает колонку с формой поиска протокола, которая заполняет всю высоту"""
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Заголовок
        title = wx.StaticText(parent, label="Найти протокол:")
        title.SetFont(wx.Font(11, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sizer.Add(title, 0, wx.ALL, 5)

        # Таблица параметров протокола
        protocol_params = [
            ("Новый тест", "header"),
            ("Номер протокола", ""),
            ("Группа исполнения", ""),
            ("Номер ЭД", ""),
            ("Категория исполнения", ""),
            ("Тип масла", ""),
            ("Тип муфты", ""),
            ("Дата и время", ""),
            ("Проверка маркировки", ""),
            ("Радиальное биение", ""),
            ("Тип РТИ", ""),
            ("Диаметр вала, мм", ""),
            ("Тип ТМС/ конденсатора", ""),
            ("Длительность обжатки", ""),
            ("Вылет вала", ""),
            ("Сочленение шлицевых соединений", ""),
            ("Тип блока ТМС", ""),
            ("Версия прошивки", ""),
            ("Оператор", "")
        ]

        # Создаем сетку для параметров протокола
        grid = wx.FlexGridSizer(len(protocol_params), 2, 5, 5)
        grid.AddGrowableCol(1, 1)  # Разрешаем растягивание второй колонки

        # Создаем шрифт для заголовка
        bold_font = wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)

        # Добавляем параметры в сетку (поля ввода вместо чекбоксов)
        self.protocol_fields = {}
        for param, param_type in protocol_params:
            if param_type == "header":
                # Заголовок - занимает всю ширину
                header = wx.StaticText(parent, label=param)
                header.SetFont(bold_font)
                grid.Add(header, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 3)
                grid.Add((0, 0), 0)  # Пустая ячейка
            else:
                # Обычный параметр с полем ввода
                label = wx.StaticText(parent, label=param)
                text_ctrl = wx.TextCtrl(parent, size=(150, -1))
                self.protocol_fields[param] = text_ctrl

                grid.Add(label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 3)
                grid.Add(text_ctrl, 0, wx.EXPAND | wx.ALL, 3)

        sizer.Add(grid, 0, wx.EXPAND | wx.ALL, 5)

        # Кнопки Сохранить и Редактировать
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Создаем кнопки
        btn_save = wx.Button(parent, label="Сохранить", size=(120, 40))
        btn_save.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_FILE_SAVE, wx.ART_BUTTON, (16, 16)), wx.LEFT)

        btn_edit = wx.Button(parent, label="Редактировать", size=(120, 40))
        btn_edit.SetBitmap(wx.ArtProvider.GetBitmap(wx.ART_EDIT, wx.ART_BUTTON, (16, 16)), wx.LEFT)

        button_sizer.Add(btn_save, 0, wx.ALL, 5)
        button_sizer.Add(btn_edit, 0, wx.ALL, 5)

        sizer.Add(button_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 5)

        # Добавляем растягивающийся спейсер, чтобы форма занимала всю высоту
        sizer.AddStretchSpacer(1)

        parent.SetSizer(sizer)


if __name__ == "__main__":
    app = wx.App()
    frame = MyFrame()
    app.MainLoop()
