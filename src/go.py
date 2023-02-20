"""
_summary_.
"""
import pandas as pd
import xlwings as xw
import pathlib
import openpyxl as xl
from openpyxl.styles import NamedStyle, Font, Border, Side, PatternFill 
import datetime as dt

class Termin():
    instances = set()
    def __init__(self,wt,start,finish,tutor,sp,ort):
        self.wt = wt
        self.tutor = tutor
        self.start = start
        self.finish = finish
        self.sp = sp
        self.ort = ort
        self.deu = []
        self.__class__.instances.add(self)
        self.col = self.day_to_col()
        self.rows = self.time_to_row()
        self.hasmoved = False
        self.isLeft = False
        

    def check_overlap(self):
        for termin in self.instances:
            if not self == termin and termin.wt == self.wt:
                if (termin.start < self.start < termin.finish) or (termin.start < self.finish < termin.finish) or (self.start < termin.start < self.finish) or (self.start < termin.finish < self.finish):
                    self.deu.append(termin)


    def correct_cols(self):
        for ter in self.deu:
            if ter.hasmoved:
                self.isLeft = not ter.isLeft
                self.hasmoved = True
            else:
                self.isLeft = True
                self.hasmoved = True

        if self.isLeft:
            self.col = (self.col[0],self.col[0])
        else:
            self.col = (self.col[1],self.col[1])

    def day_to_col(self) -> tuple:
        m_dict = {"Montag":(4,5),"Dienstag":(6,7),"Mittwoch":(8,9),"Donnerstag":(10,11),"Freitag":(12,13)}
        return m_dict[self.wt]

    def time_to_row(self) -> tuple:
        finish_dt = dt.datetime.combine(dt.datetime.now().date(),self.finish)
        start_dt = dt.datetime.combine(dt.datetime.now().date(),self.start)
        first_row = dt.datetime.combine(dt.datetime.now().date(),dt.time(8,0))
        lower_diff = start_dt-first_row
        upper_diff = finish_dt-first_row

        start_row = lower_diff.seconds // 3600 * 2 + (lower_diff.seconds % 3600) // 1800 + 5
        end_row = upper_diff.seconds // 3600 * 2 + (upper_diff.seconds % 3600) // 1800 + 4

        return(start_row,end_row)

    def fill_colors(self,ws):
        cols = {"Rüsselsheim":"E2EFDA","WBS":"DDEBF7","online":"FFF2CC"}
        color = cols[self.ort]
        for rows in ws.iter_rows(min_row=self.rows[0], max_row=self.rows[1], min_col=self.col[0], max_col=self.col[1]): 
            for cell in rows:
                cell.fill = PatternFill(start_color=color, end_color=color,fill_type = "solid")


def load_template() -> pathlib.Path:
    path1 = pathlib.Path(__file__).parents[1] / "Wochenplanersteller.xlsm"
    path2 = pathlib.Path(__file__).parents[1] / "Ergebnis.xlsx"

    wb1 = xw.Book(path1)
    wb2 = xw.Book()

    ws1 = wb1.sheets(1)
    ws1.copy(before=wb2.sheets[0])
    wb2.save(path2)
    wb2.app.quit()

    wb= xl.load_workbook(path2)
    sheet_names = wb.sheetnames
    for sheet in sheet_names[1:]:
        wb.remove(wb[sheet])
    ws = wb[sheet_names[0]]
    ws.title = 'Wochenplan'
    wb.save(path2)
    return wb, ws, path2

def prep_styles() -> list:
    highlight = NamedStyle(name="highlight")
    bd = Side(style='thick', color="000000")
    highlight.border= Border(left=bd, top=bd, right=bd, bottom=bd)



if __name__ == "__main__":

    df = pd.read_excel("Liste.xlsx")
    df = df.dropna(subset="Anfangszeit")
    for index, row in df.iterrows():
        Termin(*row)
    for t in Termin.instances:
        t.check_overlap()

    # for t in Termin.instances:
    #     print(t,t.wt,t.start,t.finish,t.deu)

    wb, ws ,path = load_template()
    for t in Termin.instances:
        if len(t.deu) > 0:
            t.correct_cols()
    for t in Termin.instances:
        t.fill_colors(ws)
    img = xl.drawing.image.Image(img_name)
    img.anchor = 'D2' # Or whatever cell location you want to use.
    ws.add_image(img)
    wb.save(path)