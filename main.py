import kivy
from kivy.app import App
from kivy.uix.gridlayout import GridLayout
from kivy.properties import ObjectProperty
import os, sys
from kivy.resources import resource_add_path
from kivy.lang import Builder
import openpyxl
from openpyxl import Workbook, load_workbook


class MyLayout(GridLayout):

    gray = (0.4,0.4,0.4,1)
    red = (245/255.0, 115/255.0, 115/255.0,1)
    blue = (115/255.0, 182/255.0, 245/255.0,1)
    green = (164/255.0, 250/255.0, 132/255.0,1)
    purple = (212/255.0, 174/255.0, 242/255.0,1)
    yellow = (240/255.0, 235/255.0, 144/255.0,1)

    xl = load_workbook('checklist.xlsx')
    sheet = xl['Sheet1']

    cells = ['A1', 'A2', 'A3', 'A4', 'A5']
    
    for cell in cells:
        if sheet[cell].value is None:
            sheet[cell].value = " "

    is_checked = [False, False, False, False, False]
    
    xl.save('checklist.xlsx')

    
    if sheet['B1'].value == "True":
        is_checked[0] = True
    if sheet['B2'].value == "True":
        is_checked[1] = True
    if sheet['B3'].value == "True":
        is_checked[2] = True
    if sheet['B4'].value == "True":
        is_checked[3] = True
    if sheet['B5'].value == "True":
        is_checked[4] = True

    def checked(self, instance, value, cell, which_box, which_color):
        if value:
            cell.value = "True"
            which_box.background_color = self.gray
        else:
            cell.value = "False"
            which_box.background_color = which_color
        self.xl.save('checklist.xlsx')
        
    def save_data(self):
        self.sheet['A1'] = self.ids.task1.text
        self.sheet['A2'] = self.ids.task2.text
        self.sheet['A3'] = self.ids.task3.text
        self.sheet['A4'] = self.ids.task4.text
        self.sheet['A5'] = self.ids.task5.text
        self.xl.save('checklist.xlsx')

    
class ChecklistApp(App):
    def build(self):
        #self.icon = self.resource_path('Calculator_icon.png')
        return MyLayout()
        #return Builder.load_file(self.resource_path('calculator.kv'))
    def start_app(self):
        self.ids.task1.text = self.sheet['A1']
        
    @staticmethod
    def resource_path(relative_path):

        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath('.')
        return os.path.join(base_path, relative_path)



if __name__ == '__main__':

    if hasattr(sys, '_MEIPASS'):
        resource_add_path((os.path.join(sys._MEIPASS)))
    
    ChecklistApp().run()


