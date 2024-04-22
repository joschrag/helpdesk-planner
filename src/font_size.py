import ctypes

def get_text_dims(text, points, font):
    class SIZE(ctypes.Structure):
            _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]
    size = SIZE(0, 0)
    if text:
        hdc = ctypes.windll.user32.GetDC(0)
        hfont = ctypes.windll.gdi32.CreateFontA(points, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, font)
        hfont_old = ctypes.windll.gdi32.SelectObject(hdc, hfont)
        ctypes.windll.gdi32.GetTextExtentPoint32A(hdc, text, len(text), ctypes.byref(size))
        ctypes.windll.gdi32.SelectObject(hdc, hfont_old)
        ctypes.windll.gdi32.DeleteObject(hfont)

    return (size.cx, size.cy)