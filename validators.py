import wx


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
