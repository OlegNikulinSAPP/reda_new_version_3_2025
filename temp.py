import wx
import wx.lib.scrolledpanel as scrolled
import wx.lib.newevent
import win32com.client as win32
from pathlib import Path

# Создаем пользовательское событие для сигнала о завершении опыта
ExperienceOneEndedEvent, EVT_EXPERIENCE_ONE_ENDED = wx.lib.newevent.NewEvent()


class ExperienceOneDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title="1. Измерение напряжения пробоя масла холодного ПЭД", size=(602, 294))

        self.file_protocol = None
        self.init_ui()
        self.Centre()

    def init_ui(self):
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # Инструкция
        instruction_text = ("Возьмите пробу масла. Произведите измерение напряжения пробоя масла. Введите измеренное и "
                            "паспортное значение:")
        instruction = wx.StaticText(panel, label=instruction_text)
        instruction.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        instruction.SetForegroundColour(wx.Colour(85, 0, 255))
        main_sizer.Add(instruction, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        # Создаем горизонтальный сайзер для основного содержимого
        content_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Левая часть с формой
        form_panel = wx.Panel(panel)
        form_sizer = wx.GridBagSizer(5, 5)

        lbl_voltage = wx.StaticText(form_panel, label="Напряжение пробоя масла, кВ")
        lbl_voltage.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        lbl_voltage.SetForegroundColour(wx.Colour(85, 0, 255))
        self.txt_voltage = wx.TextCtrl(form_panel, style=wx.TE_PROCESS_ENTER)
        self.txt_voltage.SetValidator(NumberValidator())

        lbl_nominal = wx.StaticText(form_panel, label="Паспортное значение, кВ")
        lbl_nominal.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        lbl_nominal.SetForegroundColour(wx.Colour(85, 0, 255))
        self.txt_nominal = wx.TextCtrl(form_panel, style=wx.TE_PROCESS_ENTER)
        self.txt_nominal.SetValidator(NumberValidator())

        lbl_result = wx.StaticText(form_panel, label="Результат")
        lbl_result.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        lbl_result.SetForegroundColour(wx.Colour(85, 0, 255))
        self.txt_result = wx.TextCtrl(form_panel, style=wx.TE_READONLY)
        self.txt_result.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))

        form_sizer.Add(lbl_voltage, (0, 0), flag=wx.ALIGN_CENTER_VERTICAL)
        form_sizer.Add(self.txt_voltage, (0, 1))
        form_sizer.Add(lbl_nominal, (1, 0), flag=wx.ALIGN_CENTER_VERTICAL)
        form_sizer.Add(self.txt_nominal, (1, 1))
        form_sizer.Add(lbl_result, (2, 0), flag=wx.ALIGN_CENTER_VERTICAL)
        form_sizer.Add(self.txt_result, (2, 1))

        # Чекбоксы
        self.check_first = wx.CheckBox(form_panel, label="Первая обкатка")
        self.check_first.SetValue(True)
        self.check_first.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))

        self.check_second = wx.CheckBox(form_panel, label="Вторая обкатка")
        self.check_second.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))

        form_sizer.Add(self.check_first, (3, 0), flag=wx.TOP, border=10)
        form_sizer.Add(self.check_second, (4, 0), flag=wx.TOP, border=5)

        form_panel.SetSizer(form_sizer)
        content_sizer.Add(form_panel, 0, wx.ALL, 10)

        # Правая часть с пояснением
        right_text = "Для успешного испытания необходимо чтобы напряжение пробоя масла было больше паспортного значения"
        right_label = wx.StaticText(panel, label=right_text)
        right_label.Wrap(200)  # Перенос текста
        content_sizer.Add(right_label, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        main_sizer.Add(content_sizer, 0, wx.EXPAND)

        # Кнопка сохранения
        self.btn_save = wx.Button(panel, label="Сохранить")
        self.btn_save.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        self.btn_save.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.btn_save.SetBackgroundColour(wx.Colour(255, 255, 127))
        self.btn_save.Bind(wx.EVT_BUTTON, self.on_save)

        main_sizer.Add(self.btn_save, 0, wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, 10)

        # Бинды для чекбоксов
        self.check_first.Bind(wx.EVT_CHECKBOX, self.on_checkbox)
        self.check_second.Bind(wx.EVT_CHECKBOX, self.on_checkbox)

        panel.SetSizer(main_sizer)

        # Устанавливаем сайзер для диалога
        dialog_sizer = wx.BoxSizer(wx.VERTICAL)
        dialog_sizer.Add(panel, 1, wx.EXPAND)
        self.SetSizer(dialog_sizer)
        self.Layout()

    def on_checkbox(self, event):
        # Обработка взаимного исключения чекбоксов
        checkbox = event.GetEventObject()
        if checkbox == self.check_first and checkbox.IsChecked():
            self.check_second.SetValue(False)
        elif checkbox == self.check_second and checkbox.IsChecked():
            self.check_first.SetValue(False)

    def on_save(self, event):
        # Проверка заполнения полей
        if not self.txt_voltage.GetValue() or not self.txt_nominal.GetValue():
            wx.MessageBox("Не заполнено одно из полей!", "Предупреждение!", wx.OK | wx.ICON_WARNING)
            return

        # Сравнение значений
        nominal = int(self.txt_nominal.GetValue())
        voltage = int(self.txt_voltage.GetValue())

        if nominal > voltage:
            self.txt_result.SetValue("Не годно!")
        else:
            self.txt_result.SetValue("Годно")

        # Изменение вида кнопки
        self.btn_save.Enable(False)
        self.btn_save.SetBackgroundColour(wx.Colour(51, 255, 153))
        self.Refresh()

        # Сохранение в Excel ну и чо это?
        try:
            excel = win32.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False

            workbook = excel.Workbooks.Open(self.file_protocol)
            sheet = workbook.Worksheets(1)

            sheet.Cells(22, 19).Value = self.txt_result.GetValue()
            sheet.Cells(22, 10).Value = nominal

            if self.check_first.GetValue():
                sheet.Cells(22, 13).Value = voltage
            elif self.check_second.GetValue():
                sheet.Cells(22, 16).Value = voltage
            else:
                wx.MessageBox("Выберите номер обкатки!", "Предупреждение!", wx.OK | wx.ICON_WARNING)
                workbook.Close(SaveChanges=False)
                excel.Quit()
                return

            workbook.Close(SaveChanges=True)
            excel.Quit()

        except Exception as e:
            wx.MessageBox(f"Ошибка при работе с Excel: {str(e)}", "Ошибка", wx.OK | wx.ICON_ERROR)
            return

        # Закрытие окна и отправка события
        self.EndModal(wx.ID_OK)
        wx.PostEvent(self, ExperienceOneEndedEvent())

    def set_file_protocol(self, file_path):
        self.file_protocol = str(Path(file_path).resolve())

    def reset_state(self):
        # Сброс состояния при закрытии
        self.txt_result.SetValue("")
        self.txt_nominal.SetValue("")
        self.txt_voltage.SetValue("")
        self.btn_save.Enable(True)
        self.btn_save.SetBackgroundColour(wx.Colour(255, 255, 127))
        self.check_first.SetValue(True)
        self.check_second.SetValue(False)

    def on_close(self, event):
        self.reset_state()
        event.Skip()



class NumberValidator(wx.Validator):
    def __init__(self):
        super().__init__()
        self.Bind(wx.EVT_CHAR, self.on_char)

    def on_char(self, event):
        key = event.GetKeyCode()
        if key < wx.WXK_SPACE or key == wx.WXK_DELETE or key > 255:
            event.Skip()
            return

        if chr(key).isdigit():
            event.Skip()
            return

        if not wx.Validator.IsSilent():
            wx.Bell()

    def Clone(self):
        return NumberValidator()

    def Validate(self, parent):
        return True

    def TransferToWindow(self):
        return True

    def TransferFromWindow(self):
        return True


class MyFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title="Приложение с параметрами ПЭД",
                         style=wx.DEFAULT_FRAME_STYLE, size=(1600, 900))

        self.SetMinSize((1400, 700))

        # Создаем Notebook с вкладками
        self.notebook = wx.Notebook(self)

        # Создаем четыре вкладки
        tabs = []
        for i in range(4):
            tabs.append(wx.Panel(self.notebook))
            self.notebook.AddPage(tabs[i], ["Основная", "Вторая", "Третья", "Четвертая"][i])

        # Настраиваем первую вкладку
        self.setup_main_tab(tabs[0])

        # Остальные вкладки - заглушки
        for i in range(1, 4):
            self.setup_stub_tab(tabs[i], f"Содержимое {['второй', 'третьей', 'четвертой'][i - 1]} вкладки")

        self.Centre()
        self.Show()

    def setup_stub_tab(self, tab, text):
        sizer = wx.BoxSizer(wx.VERTICAL)
        label = wx.StaticText(tab, label=text)
        sizer.Add(label, 0, wx.ALL, 20)
        tab.SetSizer(sizer)

    def setup_main_tab(self, tab):
        splitter = wx.SplitterWindow(tab)

        # Левая панель с кнопками и контролами
        left_panel = wx.Panel(splitter)
        left_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Колонка с кнопками
        buttons_panel = wx.Panel(left_panel)
        self.setup_buttons_column(buttons_panel)

        # Колонка с контролами
        controls_panel = wx.Panel(left_panel)
        self.setup_controls_column(controls_panel)

        left_sizer.Add(buttons_panel, 0, wx.EXPAND | wx.ALL, 3)
        left_sizer.Add(controls_panel, 0, wx.EXPAND | wx.ALL, 3)
        left_panel.SetSizer(left_sizer)

        # Правая панель с параметрами и поиском
        right_panel = scrolled.ScrolledPanel(splitter)
        right_panel.SetupScrolling()
        right_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # Колонка с параметрами
        params_panel = wx.Panel(right_panel)
        self.setup_parameters_column(params_panel)

        # Колонка с поиском протокола
        search_panel = wx.Panel(right_panel)
        self.setup_search_protocol_column(search_panel)

        right_sizer.Add(params_panel, 1, wx.EXPAND | wx.ALL, 3)
        right_sizer.Add(search_panel, 0, wx.EXPAND | wx.ALL, 3)
        right_panel.SetSizer(right_sizer)

        # Настройка разделителя
        splitter.SplitVertically(left_panel, right_panel, sashPosition=400)
        splitter.SetMinimumPaneSize(400)

        # Основной sizer для вкладки
        tab_sizer = wx.BoxSizer(wx.VERTICAL)
        tab_sizer.Add(splitter, 1, wx.EXPAND)
        tab.SetSizer(tab_sizer)

    def setup_buttons_column(self, parent):
        sizer = wx.BoxSizer(wx.VERTICAL)

        # Список кнопок
        buttons = [
            "📥 Загрузить", "💾 Сохранить", "🧮 Рассчитать", "🧹 Очистить",
            "📊 Экспорт", "📁 Импорт", "🖨️ Печать", "⚙️ Настройки",
            "❓ Справка", "ℹ️ О программе", "🔧 Инструмент 1", "🔨 Инструмент 2",
            "📏 Измерение", "📐 Калибровка", "🔍 Поиск"
        ]

        # Создаем кнопки и привязываем обработчики
        for text in buttons:
            btn = wx.Button(parent, label=text, size=(150, 40))

            # Привязываем обработчик для кнопки "Загрузить"
            if text.startswith("📥 Загрузить"):
                btn.Bind(wx.EVT_BUTTON, self.on_load_button)

            sizer.Add(btn, 0, wx.EXPAND | wx.ALL, 2)

        sizer.AddStretchSpacer(1)
        parent.SetSizer(sizer)

    def on_load_button(self, event):
        """Обработчик нажатия кнопки Загрузить"""
        # Создаем и показываем модальное диалоговое окно ExperienceOneDialog
        dlg = ExperienceOneDialog(self)
        # Установите file_protocol если нужно
        # dlg.set_file_protocol("путь/к/файлу.xlsx")
        dlg.ShowModal()
        dlg.Destroy()  # Уничтожаем диалог после закрытия

    def setup_controls_column(self, parent):
        sizer = wx.BoxSizer(wx.VERTICAL)

        self.text_ctrl = wx.TextCtrl(parent, size=(180, 25))
        self.list_box = wx.ListBox(parent, size=(180, -1))

        sizer.Add(self.text_ctrl, 0, wx.EXPAND | wx.ALL, 3)
        sizer.Add(self.list_box, 1, wx.EXPAND | wx.ALL, 3)

        parent.SetSizer(sizer)

    def setup_parameters_column(self, parent):
        # Список параметров
        parameters = [
            ("⚡ Напряжение трансформатора, В", ""),
            ("📋 Номинальные параметры ПЭД", "header"),
            ("🔧 Тип", ""),
            ("💪 Мощность, кВт", ""),
            ("⚡ Напряжение, В", ""),
            ("🔌 Ток, А", ""),
            ("🔄 Частота вращения", ""),
            ("📏 Сопр. обмотки (Ом)", ""),
            ("📊 Сопр. изоляции, МОм", ""),
            ("🚀 Напр. разгона, В", ""),
            ("⚙️ Момент проворачивания", ""),
            ("⚡ Напр. к.з., В", ""),
            ("🔌 Ток к.з., А", ""),
            ("💥 Потери к.з., кВт", ""),
            ("🔌 Ток х.х., А", ""),
            ("⚡ Напр. х.х., В", ""),
            ("💥 Потери х.х., кВт", ""),
            ("⏱️ Выбег, с", ""),
            ("📳 Вибрация, мм/с", ""),
            ("⚙️ Крутящий момент", ""),
            ("💥 Потери в нагр. сост.", ""),
            ("⚡ Испыт. напр. изол., В", ""),
            ("🔌 Испыт. изол. обм., В", ""),
            ("⚡ Напр. опыта КЗ", "")
        ]

        grid = wx.FlexGridSizer(len(parameters), 2, 5, 5)
        grid.AddGrowableCol(1, 1)

        self.param_fields = {}
        for param, param_type in parameters:
            if param_type == "header":
                header = wx.StaticText(parent, label=param)
                header.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
                grid.Add(header, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 3)
                grid.Add((0, 0), 0)
            else:
                label = wx.StaticText(parent, label=param)
                text_ctrl = wx.TextCtrl(parent, size=(150, -1))
                self.param_fields[param] = text_ctrl

                grid.Add(label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 3)
                grid.Add(text_ctrl, 0, wx.EXPAND | wx.ALL, 3)

        parent.SetSizer(grid)

    def setup_search_protocol_column(self, parent):
        sizer = wx.BoxSizer(wx.VERTICAL)

        title = wx.StaticText(parent, label="Найти протокол:")
        title.SetFont(wx.Font(11, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sizer.Add(title, 0, wx.ALL, 5)

        # Параметры протокола
        protocol_params = [
            ("🧪 Новый теста", "header"),
            ("📋 Номер протокола", ""),
            ("👥 Группа исполнения", ""),
            ("🔢 Номер ЭД", ""),
            ("🏷️ Категория исполнения", ""),
            ("🛢️ Тип масла", ""),
            ("🔩 Тип муфты", ""),
            ("📅 Дата и время", ""),
            ("🏷️ Проверка маркировки", ""),
            ("📏 Радиальное биение", ""),
            ("🔧 Тип РТИ", ""),
            ("📐 Диаметр вала, мм", ""),
            ("⚡ Тип ТМС/ конденсатора", ""),
            ("⏱️ Длительность обжатки", ""),
            ("📏 Вылет вала", ""),
            ("🔗 Сочленение шлицевых соединений", ""),
            ("🔌 Тип блока ТМС", ""),
            ("💾 Версия прошивки", ""),
            ("👤 Оператор", "")
        ]

        grid = wx.FlexGridSizer(len(protocol_params), 2, 5, 5)
        grid.AddGrowableCol(1, 1)

        self.protocol_fields = {}
        for param, param_type in protocol_params:
            if param_type == "header":
                header = wx.StaticText(parent, label=param)
                header.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
                grid.Add(header, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 3)
                grid.Add((0, 0), 0)
            else:
                label = wx.StaticText(parent, label=param)
                text_ctrl = wx.TextCtrl(parent, size=(150, -1))
                self.protocol_fields[param] = text_ctrl

                grid.Add(label, 0, wx.ALIGN_CENTER_VERTICAL | wx.ALL, 3)
                grid.Add(text_ctrl, 0, wx.EXPAND | wx.ALL, 3)

        sizer.Add(grid, 0, wx.EXPAND | wx.ALL, 5)

        # Кнопки внизу
        btn_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_save = wx.Button(parent, label="💾 Сохранить", size=(120, 40))
        btn_edit = wx.Button(parent, label="✏️ Редактировать", size=(120, 40))

        btn_sizer.Add(btn_save, 0, wx.ALL, 5)
        btn_sizer.Add(btn_edit, 0, wx.ALL, 5)

        sizer.Add(btn_sizer, 0, wx.ALIGN_CENTER | wx.ALL, 5)
        sizer.AddStretchSpacer(1)

        parent.SetSizer(sizer)


if __name__ == "__main__":
    app = wx.App()
    frame = MyFrame()
    app.MainLoop()
