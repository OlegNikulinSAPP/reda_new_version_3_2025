import wx
import wx.lib.scrolledpanel as scrolled
from events import ExperienceOneEndedEvent, EVT_EXPERIENCE_ONE_ENDED
import win32com.client as win32
from pathlib import Path
from validators import NumberValidator


class ExperienceOneDialog(wx.Dialog):
    def __init__(self, parent):
        super().__init__(parent, title="1. Измерение напряжения пробоя масла холодного ПЭД", size=(602, 294))
        self.file_protocol = None
        self.init_ui()
        self.Centre()

    def init_ui(self):
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        instruction_text = ("Возьмите пробу масла. Произведите измерение напряжения пробоя масла. "
                            "Введите измеренное и паспортное значения:")
        instruction = wx.StaticText(panel, label=instruction_text)
        instruction.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        instruction.SetForegroundColour(wx.Colour(85, 0, 255))
        main_sizer.Add(instruction, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        content_sizer = wx.BoxSizer(wx.HORIZONTAL)
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

        self.check_first = wx.CheckBox(form_panel, label="Первая обкатка")
        self.check_first.SetValue(True)
        self.check_first.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))

        self.check_second = wx.CheckBox(form_panel, label="Вторая обкатка")
        self.check_second.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))

        form_sizer.Add(self.check_first, (3, 0), flag=wx.TOP, border=10)
        form_sizer.Add(self.check_second, (4, 0), flag=wx.TOP, border=5)

        form_panel.SetSizer(form_sizer)
        content_sizer.Add(form_panel, 0, wx.ALL, 10)

        right_text = "Для успешного испытания необходимо чтобы напряжение пробоя масла было больше паспортного значения"
        right_label = wx.StaticText(panel, label=right_text)
        right_label.Wrap(200)  # Перенос текста

        content_sizer.Add(right_label, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        main_sizer.Add(content_sizer, 0, wx.EXPAND)

        self.btn_save = wx.Button(panel, label="Сохранить")
        self.btn_save.SetFont(wx.Font(10, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        self.btn_save.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.btn_save.SetBackgroundColour(wx.Colour(255, 255, 127))

        self.btn_save.Bind(wx.EVT_BUTTON, self.on_save)

        main_sizer.Add(self.btn_save, 0, wx.ALIGN_RIGHT | wx.RIGHT | wx.BOTTOM, 10)

        self.check_first.Bind(wx.EVT_CHECKBOX, self.on_checkbox)
        self.check_second.Bind(wx.EVT_CHECKBOX, self.on_checkbox)

        panel.SetSizer(main_sizer)

        dialog_sizer = wx.BoxSizer(wx.VERTICAL)
        dialog_sizer.Add(panel, 1, wx.EXPAND)
        self.SetSizer(dialog_sizer)

        self.Layout()

    def on_checkbox(self, event):
        checkbox = event.GetEventObject()
        if checkbox == self.check_first and checkbox.IsChecked():
            self.check_second.SetValue(False)
        elif checkbox == self.check_second and checkbox.IsChecked():
            self.check_first.SetValue(False)

    def on_save(self, event):
        if not self.txt_voltage.GetValue() or not self.txt_nominal.GetValue():
            wx.MessageBox("Не заполнено одно из полей!", "Предупреждение!", wx.OK | wx.ICON_WARNING)
            return

        nominal = int(self.txt_nominal.GetValue())
        voltage = int(self.txt_voltage.GetValue())

        if nominal > voltage:
            self.txt_result.SetValue("Не годно!")
        else:
            self.txt_result.SetValue("Годно")

        self.btn_save.Enable(False)
        self.btn_save.SetBackgroundColour(wx.Colour(51, 255, 153))
        self.Refresh()

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

        self.EndModal(wx.ID_OK)

        wx.PostEvent(self, ExperienceOneEndedEvent())
