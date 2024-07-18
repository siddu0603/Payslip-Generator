import os
import webbrowser
from kivy.app import App
from kivy.uix.gridlayout import GridLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.popup import Popup
import pandas as pd
from docxtpl import DocxTemplate
import datetime

class SalaryStatementGenerator(App):
    def build(self):
        self.title = 'Salary Statement Generator'
        
        main_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        
        self.grid = GridLayout(cols=2, spacing=10, size_hint_y=None)
        self.grid.bind(minimum_height=self.grid.setter('height'))

        self.grid.add_widget(Label(text='Employee ID:'))
        self.id_input = TextInput()
        self.grid.add_widget(self.id_input)

        main_layout.add_widget(self.grid)

        search_button = Button(text='Search Document', size_hint_y=None, height=50)
        search_button.bind(on_press=self.search_document)
        main_layout.add_widget(search_button)

        self.file_label = Label(text="No file selected")
        main_layout.add_widget(self.file_label)

        file_chooser = FileChooserListView()
        file_chooser.filters = ['.xls', '.xlsx']
        main_layout.add_widget(file_chooser)

        upload_button = Button(text="Generate Statements", size_hint=(None, None), size=(200, 50))
        upload_button.bind(on_press=lambda x: self.generate_salary_statements(file_chooser.path, file_chooser.selection))
        main_layout.add_widget(upload_button)
        
        return main_layout

    def search_document(self, instance):
        emp_id = self.id_input.text
        doc_filename = os.path.join('E:\\python\\Employee', f'salary_statement_{emp_id}.docx')

        if os.path.exists(doc_filename):
            webbrowser.open(doc_filename)
        else:
            popup = Popup(title='Document Not Found', content=Label(text=f'Document for Employee ID {emp_id} not found.'), size_hint=(None, None), size=(400, 200))
            popup.open()

    def generate_salary_statements(self, path, filename):
        if filename:
            selected_file = filename[0]
            self.file_label.text = f"Selected file: {selected_file}"
            try:
                df = pd.read_excel(selected_file)
                for index, row in df.iterrows():
                    today_date = datetime.datetime.today().strftime('%B %d %Y')
                    emp_id = row['EMPLOYEE_CODE']
                    name = row['EMPLOYEE_NAME']
                    des = row['DESIGNATION']
                    doj = row['DATE_OF_JOINING']
                    es = row['EMPLOYEE_STATUS']
                    sftm = row['STATEMENT_FOR_THE_MONTH']
                    bp = row['BASIC_PAY']
                    hra = row['HOUSE_RENT_ALLOWANCE']
                    cca = row["City_compensation_allowance"]
                    tra = row['TRAVEL_ALLOWANCE']
                    fa = row['FOOD_ALLOWANCE']
                    pi = row['PERFORMANCE_INCENTIVES']
                    pt = row['PROFESSIONAL_TAX']
                    income_tax = row['INCOME_TAX']
                    pf = row['PROVIDENT_FUND']
                    esi = row['ESI']
                    llop = row['LEAVE_LOSS_OF_PAY']
                    others = row['OTHERS']
                    gross = row['GROSS_PAY']
                    ded = row['DEDUCTIONS']
                    np = row['NET_PAY']
                    auth = row['AUTHORISED']

                    self.generate_salary_statement(emp_id, name, des, doj, es, sftm, bp, hra, cca, today_date, tra, fa, pi, pt, income_tax, pf, esi, llop, others, np, ded, gross,auth)
                popup = Popup(title='Success', content=Label(text='Salary statements generated successfully.'), size_hint=(None, None), size=(400, 200))
                print("Successfully generated")
                popup.open()
            except Exception as e:
                popup = Popup(title='Error', content=Label(text=f"Error generating: {e}"), size_hint=(None, None), size=(400, 200))
                popup.open()
        else:
            self.file_label.text = "No file selected"

    def generate_salary_statement(self, emp_id, name, des, doj, es, sftm, bp, hra, cca, today_date, tra, fa, pi, pt, income_tax, pf, esi, llop, others, np, ded, gross,auth):
        folder_path = os.path.join(os.getcwd(), 'Employee')
        os.makedirs(folder_path, exist_ok=True)
        salary_filename = f'salary_statement_{emp_id}.docx'
        salary_path = os.path.join(folder_path, salary_filename)

        doj = doj.strftime("%B-%m-%y")
        context = {
            'today_date': today_date,
            'emp_id': emp_id,
            'des': des,
            'name': name,
            'doj': doj,
            'es': es,
            'sftm': sftm,
            'bp': bp,
            'hra': hra,
            'cca': cca,
            'tra': tra,
            'fa': fa,
            'pi': pi,
            'pt': pt,
            'income_tax': income_tax,
            'pf': pf,
            'esi': esi,
            'llop': llop,
            'others': others,
            'np': np,
            'gross': gross,
            'ded': ded,
            'auth': auth
        }

        template_path = "SYMBIOSYS TECHNOLOGIES_salary.docx"
        if not os.path.exists(template_path):
            popup = Popup(title='Error', content=Label(text='Template file not found.'), size_hint=(None, None), size=(400, 200))
            popup.open()
            return

        doc = DocxTemplate(template_path)
        doc.render(context)
        doc.save(salary_path)

        

if _name_ == '_main_':
    SalaryStatementGenerator().run()
