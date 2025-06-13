import tkinter as tk
from tkinter import filedialog
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from statsmodels.tsa.arima.model import ARIMA
from sklearn.preprocessing import StandardScaler
import torch
from torch import nn
from skorch import NeuralNetRegressor
import numpy as np


class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        x, y, _, cy = self.widget.bbox("insert") or (0, 0, 0, 0)
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + cy + 25
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.geometry(f"+{x}+{y}")
        label = tk.Label(
            tw, text=self.text, justify='left',
            background="#FFFFFF", relief='flat', borderwidth=1,
            font=("Bahnschrift", 14, "normal")
        )
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

class FileSelectorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Домашня сторінка")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg = "#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.22, width=800, height=400)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.2, width=800, height=400)

        self.label0 = tk.Label(self, text="СИСТЕМА ДЛЯ ОЦІНКИ Й ПЕРЕДБАЧЕННЯ ТЕНДЕНЦІЙ ЗМІН\nЕКОЛОГІЧНОЇ ТА РАДІАЦІЙНОЇ ОБСТАНОВКИ В ЗОНІ\nРОЗТАШУВАННЯ АТОМНИХ ЕЛЕКТРОСТАНЦІЙ", font=("Bahnschrift", 14))
        self.label0.configure(bg="#FFFFFF", fg="#000000", anchor="center")
        self.label0.place(relx=0.18, rely=0.34, width=700, height=100)

        self.h_button = tk.Button(self, text="⌂",
                                       font=("Bahnschrift", 14))
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                  font=("Bahnschrift", 14))
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                       font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Головна сторінка", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        self.file_Button = tk.Button(self, text="Про автора", font=("Bahnschrift", 14), command = self.show_author_popup)
        self.file_Button.configure(bg="#3C6E71", fg="#FFFFFF")
        self.file_Button.place(relx=0.29, rely=0.59, width=200, height=40)

        self.file_Button1 = tk.Button(self, text="Інструкції", font=("Bahnschrift", 14), command=self.show_instructions)
        self.file_Button1.configure(bg="#3C6E71", fg="#FFFFFF")
        self.file_Button1.place(relx=0.53, rely=0.59, width=200, height=40)

        self.browse_button = tk.Button(self, text="Для початку роботи оберіть файл для аналізу:", font=("Bahnschrift", 14), command=self.browse_file)
        self.browse_button.configure(bg="#3C6E71", fg="#FFFFFF")
        self.browse_button.place(relx=0.157, rely=0.69, width=692, height=77)

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.', justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)




    def browse_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.selected_file = file_path
            self.open_next_window()
        else:
            messagebox.showwarning("Увага", "Файл не обрано!")

    def show_author_popup(self):
        popup = tk.Toplevel(self)
        popup.title("Про автора")
        popup.geometry("250x150")  # Set size of popup
        popup.resizable(False, False)

        label = tk.Label(popup, text="Автор: Єлизавета Кайдан\nГрупа: ПМА-42",
                         font=("Bahnschrift", 12))
        label.pack(expand=True, padx=10, pady=20)

        btn_close = tk.Button(popup, text="Закрити", command=popup.destroy)
        btn_close.pack(pady=10)

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def open_next_window(self):
        NextWindow(self, self.selected_file)
        self.withdraw()
class NextWindow(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.master = master
        self.title("Оберіть дію")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")

        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.22, width=800, height=400)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.2, width=800, height=400)


        self.label = tk.Label(self, text="Оберіть спосіб аналізу даних:", font=("Bahnschrift", 28))
        self.label.configure(bg="#FFFFFF",anchor="center")
        self.label.place(relx=0.21, rely=0.25, width=600, height=50)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Оберіть дію", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)





        self.button1 = tk.Button(self, text="Аналіз на основі заданих значень",
                                       font=("Bahnschrift", 18), command=self.open_analysis_window1)
        self.button1.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.button1.place(relx=0.20, rely=0.44, width=601, height=44)

        self.button2 = tk.Button(self, text="Передбачення за допомогою TCN",
                                 font=("Bahnschrift", 18), command=self.open_analysis_window2)
        self.button2.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.button2.place(relx=0.20, rely=0.57, width=601, height=44)

        self.button3 = tk.Button(self, text="Передбачення за допомогою ARIMA",
                                       font=("Bahnschrift", 18), command=self.open_analysis_window3)
        self.button3.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.button3.place(relx=0.20, rely=0.70, width=601, height=44 )



    def on_close(self):
        self.destroy()
        self.master.deiconify()

    def open_analysis_window1(self):
        AnalysisWindow1(self.master, self.selected_file)
        self.withdraw()

    def open_analysis_window2(self):
        AnalysisWindow2(self.master, self.selected_file)
        self.withdraw()

    def open_analysis_window3(self):
        AnalysisWindow3(self.master, self.selected_file)
        self.withdraw()

    def go_back(self):
        FileSelectorApp()
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)
class AnalysisWindow1(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.title("Оберіть АЕС для аналізу")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Оберіть дію", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.17, width=800, height=450)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.15, width=800, height=450)

        self.but1 = tk.Button(self, text="Загальний аналіз", font=("Bahnschrift", 14), command=self.go_Analysis_ZAG)
        self.but1.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but1.place(relx=0.2, rely=0.22, width=600, height=50)

        self.but2 = tk.Button(self, text="Аналіз для ЗАЕС", font=("Bahnschrift", 14), command=self.go_Analysis_ZNPP)
        self.but2.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but2.place(relx=0.2, rely=0.35, width=600, height=50)

        self.but3 = tk.Button(self, text="Аналіз для РАЕС", font=("Bahnschrift", 14), command=self.go_Analysis_RNPP)
        self.but3.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but3.place(relx=0.2, rely=0.48, width=600, height=50)

        self.but4 = tk.Button(self, text="Аналіз для ПАЕС", font=("Bahnschrift", 14), command=self.go_Analysis_PNPP)
        self.but4.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but4.place(relx=0.2, rely=0.61, width=600, height=50)

        self.but5 = tk.Button(self, text="Аналіз для ХАЕС", font=("Bahnschrift", 14), command=self.go_Analysis_KhNPP)
        self.but5.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but5.place(relx=0.2, rely=0.74, width=600, height=50)

    def go_back(self):
        NextWindow(self.master.master, self.selected_file)
        self.withdraw()

    def go_Analysis_ZAG(self):
        Analysis_ZAG(self.master.master, self.selected_file)
        self.withdraw()

    def go_Analysis_RNPP(self):
        Analysis_RNPP(self.master.master, self.selected_file)
        self.withdraw()

    def go_Analysis_PNPP(self):
        Analysis_PNPP(self.master.master, self.selected_file)
        self.withdraw()

    def go_Analysis_ZNPP(self):
        Analysis_ZNPP(self.master.master, self.selected_file)
        self.withdraw()

    def go_Analysis_KhNPP(self):
        Analysis_KhNPP(self.master.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)

class AnalysisWindow2(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.title("Оберіть АЕС для передбачення")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Оберіть дію", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.17, width=800, height=450)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.15, width=800, height=450)

        self.but1 = tk.Button(self, text="Загальне передбачення", font=("Bahnschrift", 14), command = self.go_TCN_ZAG)
        self.but1.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but1.place(relx=0.2, rely=0.22, width=600, height=50)

        self.but2 = tk.Button(self, text="Передбачення для ЗАЕС", font=("Bahnschrift", 14), command = self.go_TCN_ZNPP)
        self.but2.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but2.place(relx=0.2, rely=0.35, width=600, height=50)

        self.but3 = tk.Button(self, text="Передбачення для РАЕС", font=("Bahnschrift", 14), command = self.go_TCN_RNPP)
        self.but3.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but3.place(relx=0.2, rely=0.48, width=600, height=50)

        self.but4 = tk.Button(self, text="Передбачення для ПАЕС", font=("Bahnschrift", 14), command = self.go_TCN_PNPP)
        self.but4.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but4.place(relx=0.2, rely=0.61, width=600, height=50)

        self.but5 = tk.Button(self, text="Передбачення для ХАЕС", font=("Bahnschrift", 14), command = self.go_TCN_KhNPP)
        self.but5.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but5.place(relx=0.2, rely=0.74, width=600, height=50)

    def go_TCN_ZAG(self):
        TCN_ZAG(self.master, self.selected_file)
        self.withdraw()

    def go_TCN_RNPP(self):
        TCN_RNPP(self.master, self.selected_file)
        self.withdraw()

    def go_TCN_PNPP(self):
        TCN_PNPP(self.master, self.selected_file)
        self.withdraw()

    def go_TCN_ZNPP(self):
        TCN_ZNPP(self.master, self.selected_file)
        self.withdraw()

    def go_TCN_KhNPP(self):
        TCN_KhNPP(self.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)

    def go_back(self):
        NextWindow(self.master.master, self.selected_file)
        self.withdraw()

class AnalysisWindow3(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.title("Оберіть АЕС для передбачення")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.cur_button = tk.Label(self, text="Оберіть дію", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.17, width=800, height=450)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.15, width=800, height=450)





        self.but1 = tk.Button(self, text="Загальне передбачення", font=("Bahnschrift", 14), command=self.go_ARIMA_ZAG)
        self.but1.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but1.place(relx=0.2, rely=0.22, width=600, height=50)

        self.but2 = tk.Button(self, text="Передбачення для ЗАЕС", font=("Bahnschrift", 14), command=self.go_Analysis_ZNPP)
        self.but2.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but2.place(relx=0.2, rely=0.35, width=600, height=50)

        self.but3 = tk.Button(self, text="Передбачення для РАЕС", font=("Bahnschrift", 14), command=self.go_ARIMA_RNPP)
        self.but3.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but3.place(relx=0.2, rely=0.48, width=600, height=50)

        self.but4 = tk.Button(self, text="Передбачення для ПАЕС", font=("Bahnschrift", 14), command=self.go_ARIMA_PNPP)
        self.but4.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but4.place(relx=0.2, rely=0.61, width=600, height=50)

        self.but5 = tk.Button(self, text="Передбачення для ХАЕС", font=("Bahnschrift", 14), command=self.go_ARIMA_KhNPP)
        self.but5.configure(bg="#3C6E71", fg="#FFFFFF", anchor="center")
        self.but5.place(relx=0.2, rely=0.74, width=600, height=50)

    def go_ARIMA_ZAG(self):
        ARIMA_ZAG(self.master, self.selected_file)
        self.withdraw()

    def go_ARIMA_RNPP(self):
        ARIMA_RNPP(self.master, self.selected_file)
        self.withdraw()

    def go_ARIMA_PNPP(self):
        ARIMA_PNPP(self.master, self.selected_file)
        self.withdraw()

    def go_Analysis_ZNPP(self):
        ARIMA_ZNPP(self.master, self.selected_file)
        self.withdraw()

    def go_ARIMA_KhNPP(self):
        ARIMA_KhNPP(self.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)

    def go_back(self):
        NextWindow(self.master.master, self.selected_file)
        self.withdraw()

class Analysis_ZAG(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")

        self.thresholds = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 100,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 50,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 200,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 75,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 300,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 60,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 150,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 120,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 130,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 110,
            "Загальний об'єм водних скидів, куб. м": 5000,
            "Сумарний показник радіоактивних викидів, %": 80,
            "Сумарний індекс скиду радіоактивних речовин, %": 90
        }


        self.title("Аналіз для всіх АЕС")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")

        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.18, width=800, height=440)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Аналіз для всіх АЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)


        options = ["Радіоактивність викидів інертних радіоактивних газів, ГБк/доба",
                   "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %",
                   "Радіоактивність викидів радіонуклідів йоду, кБк/доба",
                   "Відсоток від допустимого рівня викидів радіонуклідів йоду, %",
                   "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба",
                   "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць",
                   "Загальний об'єм водних скидів, куб. м",
                   "Сумарний показник радіоактивних викидів, %",
                   "Сумарний індекс скиду радіоактивних речовин, %"]
        self.selected_option = tk.StringVar(value=options[0])
        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)

        self.output1 = tk.Label(self, text="Мода:", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center", highlightbackground="black", highlightthickness=1)
        self.output1.place(relx=0.15, rely=0.38, width=320, height=60)

        self.output2 = tk.Label(self, text="Медіана:", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center", highlightbackground="black", highlightthickness=1)
        self.output2.place(relx=0.53, rely=0.38, width=320, height=60)

        self.output3 = tk.Label(self, text="Середнє арифметичне:", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center", highlightbackground="black", highlightthickness=1)
        self.output3.place(relx=0.15, rely=0.53, width=700, height=60)

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        self.load_and_compute_modes()
        self.update_outputs(self.selected_option.get())


    def load_and_compute_modes(self):
        self.mode_results = {}
        self.median_results = {}
        self.mean_results = {}

        try:
            df = pd.read_excel(self.selected_file, engine='openpyxl')

            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
            df = df.infer_objects(copy=False)
            df.interpolate(method='linear', limit_direction='both', inplace=True)

            options = {
                "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
                "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
                "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
                "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
                "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
                "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
                "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
                "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
                "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
                "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
                "Загальний об'єм водних скидів, куб. м": 10,
                "Сумарний показник радіоактивних викидів, %": 11,
                "Сумарний індекс скиду радіоактивних речовин, %": 12
            }

            for option in options:
                if option in df.columns:
                    series = df[option]
                    if series.dtype == object:
                        series = series.astype(str).str.replace(",", ".", regex=False)
                        series = series.str.replace(r"[^\d\.]+", "", regex=True)
                        series = pd.to_numeric(series, errors='coerce')

                    series = series.dropna()
                    self.mode_results[option] = series.mode().tolist() or ["—"]
                    self.median_results[option] = series.median() if not series.empty else "—"
                    self.mean_results[option] = series.mean() if not series.empty else "—"
                else:
                    self.mode_results[option] = ["—"]
                    self.median_results[option] = "—"
                    self.mean_results[option] = "—"

        except Exception as e:
            messagebox.showerror("Помилка завантаження даних", f"Не вдалося обробити файл:\n{e}")

    def update_outputs(self, selection):
        mode = self.mode_results.get(selection, ["—"])
        median = self.median_results.get(selection, "—")
        mean = self.mean_results.get(selection, "—")

        if hasattr(self, 'warning_label'):
            self.warning_label.destroy()

        threshold = self.thresholds.get(selection)
        mean_warning = ""
        mode_warning = ""
        median_warning = ""

        if isinstance(mean, (float, int)) and threshold is not None and mean > threshold:
            mean_warning = " ⚠"
        if isinstance(mode, (float, int)) and threshold is not None and mode > threshold:
            mode_warning = " ⚠"
        if isinstance(median, (float, int)) and threshold is not None and median > threshold:
            median_warning = " ⚠"

        self.output1.config(text=f"Мода: {', '.join(map(str, mode)) if mode else '—'}{mode_warning}" if isinstance(mode, (float, int)) else f"Мода: {', '.join(map(str, mode)) if mode else '—'}")
        self.output2.config(text=f"Медіана: {median}{median_warning}" if isinstance(median, (float, int)) else f"Медіана: {median}")
        self.output3.config(text=f"Середнє арифметичне: {mean:.3f}{mean_warning}" if isinstance(mean, (float, int)) else f"Середнє арифметичне: {mean}")

        if mean_warning:
            self.warning_label = tk.Label(self.output3.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.68, rely=0.54, width=50, height=50)
            ToolTip(self.warning_label, "Показник середнього арифметичного перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")
        if median_warning:
            self.warning_label = tk.Label(self.output2.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.78, rely=0.39, width=50, height=50)
            ToolTip(self.warning_label, "Показник медіани перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")
        if mode_warning:
            self.warning_label = tk.Label(self.output1.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.39, rely=0.39, width=50, height=50)
            ToolTip(self.warning_label, "Показник моди перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")





    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
            f"{self.output1.cget('text')}",
            f"{self.output2.cget('text')}",
            f"{self.output3.cget('text')}",
            f"Файл даних: {self.master.selected_file if hasattr(self.master, 'selected_file') else 'Невідомо'}",
            "",
            ""
        ]

        try:
            with open("результат_аналізу.txt", "a", encoding="utf-8") as f:
                f.write("\n".join(result_lines))
        except Exception as e:
            print(f"Error loading file: {e}")






    def go_back(self):
        AnalysisWindow1(self.master.master, self.selected_file)
        self.withdraw()
    def go_home(self):
        self.destroy()
        self.master.deiconify()
    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)
class Analysis_ZNPP(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")

        self.thresholds = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 100,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 50,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 200,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 75,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 300,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 60,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 150,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 120,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 130,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 110,
            "Загальний об'єм водних скидів, куб. м": 5000,
            "Сумарний показник радіоактивних викидів, %": 80,
            "Сумарний індекс скиду радіоактивних речовин, %": 90
        }


        self.title("Аналіз для ЗАЕС")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")

        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.18, width=800, height=440)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Аналіз для ЗАЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)


        options = ["Радіоактивність викидів інертних радіоактивних газів, ГБк/доба",
                   "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %",
                   "Радіоактивність викидів радіонуклідів йоду, кБк/доба",
                   "Відсоток від допустимого рівня викидів радіонуклідів йоду, %",
                   "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба",
                   "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць",
                   "Загальний об'єм водних скидів, куб. м",
                   "Сумарний показник радіоактивних викидів, %",
                   "Сумарний індекс скиду радіоактивних речовин, %"]
        self.selected_option = tk.StringVar(value=options[0])
        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)

        self.output1 = tk.Label(self, text="Мода:", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center", highlightbackground="black", highlightthickness=1)
        self.output1.place(relx=0.15, rely=0.38, width=320, height=60)

        self.output2 = tk.Label(self, text="Медіана:", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center", highlightbackground="black", highlightthickness=1)
        self.output2.place(relx=0.53, rely=0.38, width=320, height=60)

        self.output3 = tk.Label(self, text="Середнє арифметичне:", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center", highlightbackground="black", highlightthickness=1)
        self.output3.place(relx=0.15, rely=0.53, width=700, height=60)

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        self.load_and_compute_modes()
        self.update_outputs(self.selected_option.get())


    def load_and_compute_modes(self):
        self.mode_results = {}
        self.median_results = {}
        self.mean_results = {}



        try:
            df = pd.read_excel(self.selected_file, engine='openpyxl')
            df = df[df['Назва станції'].str.strip() == "ЗАЕС"]

            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
            df = df.infer_objects(copy=False)
            df.interpolate(method='linear', limit_direction='both', inplace=True)

            options = {
                "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
                "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
                "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
                "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
                "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
                "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
                "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
                "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
                "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
                "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
                "Загальний об'єм водних скидів, куб. м": 10,
                "Сумарний показник радіоактивних викидів, %": 11,
                "Сумарний індекс скиду радіоактивних речовин, %": 12
            }

            for option in options:
                if option in df.columns:
                    series = df[option]
                    if series.dtype == object:
                        series = series.astype(str).str.replace(",", ".", regex=False)
                        series = series.str.replace(r"[^\d\.]+", "", regex=True)
                        series = pd.to_numeric(series, errors='coerce')

                    series = series.dropna()
                    self.mode_results[option] = series.mode().tolist() or ["—"]
                    self.median_results[option] = series.median() if not series.empty else "—"
                    self.mean_results[option] = series.mean() if not series.empty else "—"
                else:
                    self.mode_results[option] = ["—"]
                    self.median_results[option] = "—"
                    self.mean_results[option] = "—"

        except Exception as e:
            messagebox.showerror("Помилка завантаження даних", f"Не вдалося обробити файл:\n{e}")

    def update_outputs(self, selection):
        mode = self.mode_results.get(selection, ["—"])
        median = self.median_results.get(selection, "—")
        mean = self.mean_results.get(selection, "—")

        if hasattr(self, 'warning_label'):
            self.warning_label.destroy()

        threshold = self.thresholds.get(selection)
        mean_warning = ""
        mode_warning = ""
        median_warning = ""

        if isinstance(mean, (float, int)) and threshold is not None and mean > threshold:
            mean_warning = " ⚠"
        if isinstance(mode, (float, int)) and threshold is not None and mode > threshold:
            mode_warning = " ⚠"
        if isinstance(median, (float, int)) and threshold is not None and median > threshold:
            median_warning = " ⚠"

        self.output1.config(text=f"Мода: {', '.join(map(str, mode)) if mode else '—'}{mode_warning}" if isinstance(mode, (float, int)) else f"Мода: {', '.join(map(str, mode)) if mode else '—'}")
        self.output2.config(text=f"Медіана: {median}{median_warning}" if isinstance(median, (float, int)) else f"Медіана: {median}")
        self.output3.config(text=f"Середнє арифметичне: {mean:.3f}{mean_warning}" if isinstance(mean, (float, int)) else f"Середнє арифметичне: {mean}")

        if mean_warning:
            self.warning_label = tk.Label(self.output3.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.68, rely=0.54, width=50, height=50)
            ToolTip(self.warning_label, "Показник середнього арифметичного перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")
        if median_warning:
            self.warning_label = tk.Label(self.output2.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.78, rely=0.39, width=50, height=50)
            ToolTip(self.warning_label, "Показник медіани перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")
        if mode_warning:
            self.warning_label = tk.Label(self.output1.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.39, rely=0.39, width=50, height=50)
            ToolTip(self.warning_label, "Показник моди перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")





    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
            f"{self.output1.cget('text')}",
            f"{self.output2.cget('text')}",
            f"{self.output3.cget('text')}",
            f"Файл даних: {self.master.selected_file if hasattr(self.master, 'selected_file') else 'Невідомо'}",
            "",
            ""
        ]

        try:
            with open("результат_аналізу.txt", "a", encoding="utf-8") as f:
                f.write("\n".join(result_lines))
        except Exception as e:
            print(f"Error loading file: {e}")






    def go_back(self):
        AnalysisWindow1(self.master.master, self.selected_file)
        self.withdraw()
    def go_home(self):
        self.destroy()
        self.master.deiconify()
    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)
class Analysis_PNPP(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")

        self.thresholds = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 100,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 50,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 200,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 75,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 300,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 60,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 150,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 120,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 130,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 110,
            "Загальний об'єм водних скидів, куб. м": 5000,
            "Сумарний показник радіоактивних викидів, %": 80,
            "Сумарний індекс скиду радіоактивних речовин, %": 90
        }

        self.title("Аналіз для ПАЕС")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")

        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.18, width=800, height=440)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Аналіз для ПАЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        options = ["Радіоактивність викидів інертних радіоактивних газів, ГБк/доба",
                   "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %",
                   "Радіоактивність викидів радіонуклідів йоду, кБк/доба",
                   "Відсоток від допустимого рівня викидів радіонуклідів йоду, %",
                   "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба",
                   "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць",
                   "Загальний об'єм водних скидів, куб. м",
                   "Сумарний показник радіоактивних викидів, %",
                   "Сумарний індекс скиду радіоактивних речовин, %"]
        self.selected_option = tk.StringVar(value=options[0])
        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)

        self.output1 = tk.Label(self, text="Мода:", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output1.place(relx=0.15, rely=0.38, width=320, height=60)

        self.output2 = tk.Label(self, text="Медіана:", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output2.place(relx=0.53, rely=0.38, width=320, height=60)

        self.output3 = tk.Label(self, text="Середнє арифметичне:", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output3.place(relx=0.15, rely=0.53, width=700, height=60)

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        self.load_and_compute_modes()
        self.update_outputs(self.selected_option.get())

    def load_and_compute_modes(self):
        self.mode_results = {}
        self.median_results = {}
        self.mean_results = {}

        try:
            df = pd.read_excel(self.selected_file, engine='openpyxl')
            df = df[df['Назва станції'].str.strip() == "ПАЕС"]

            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
            df = df.infer_objects(copy=False)
            df.interpolate(method='linear', limit_direction='both', inplace=True)

            options = {
                "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
                "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
                "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
                "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
                "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
                "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
                "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
                "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
                "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
                "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
                "Загальний об'єм водних скидів, куб. м": 10,
                "Сумарний показник радіоактивних викидів, %": 11,
                "Сумарний індекс скиду радіоактивних речовин, %": 12
            }

            for option in options:
                if option in df.columns:
                    series = df[option]
                    if series.dtype == object:
                        series = series.astype(str).str.replace(",", ".", regex=False)
                        series = series.str.replace(r"[^\d\.]+", "", regex=True)
                        series = pd.to_numeric(series, errors='coerce')

                    series = series.dropna()
                    self.mode_results[option] = series.mode().tolist() or ["—"]
                    self.median_results[option] = series.median() if not series.empty else "—"
                    self.mean_results[option] = series.mean() if not series.empty else "—"
                else:
                    self.mode_results[option] = ["—"]
                    self.median_results[option] = "—"
                    self.mean_results[option] = "—"

        except Exception as e:
            messagebox.showerror("Помилка завантаження даних", f"Не вдалося обробити файл:\n{e}")

    def update_outputs(self, selection):
        mode = self.mode_results.get(selection, ["—"])
        median = self.median_results.get(selection, "—")
        mean = self.mean_results.get(selection, "—")

        if hasattr(self, 'warning_label'):
            self.warning_label.destroy()

        threshold = self.thresholds.get(selection)
        mean_warning = ""
        mode_warning = ""
        median_warning = ""

        if isinstance(mean, (float, int)) and threshold is not None and mean > threshold:
            mean_warning = " ⚠"
        if isinstance(mode, (float, int)) and threshold is not None and mode > threshold:
            mode_warning = " ⚠"
        if isinstance(median, (float, int)) and threshold is not None and median > threshold:
            median_warning = " ⚠"

        self.output1.config(text=f"Мода: {', '.join(map(str, mode)) if mode else '—'}{mode_warning}" if isinstance(mode,
                                                                                                                   (
                                                                                                                   float,
                                                                                                                   int)) else f"Мода: {', '.join(map(str, mode)) if mode else '—'}")
        self.output2.config(
            text=f"Медіана: {median}{median_warning}" if isinstance(median, (float, int)) else f"Медіана: {median}")
        self.output3.config(text=f"Середнє арифметичне: {mean:.3f}{mean_warning}" if isinstance(mean, (
        float, int)) else f"Середнє арифметичне: {mean}")

        if mean_warning:
            self.warning_label = tk.Label(self.output3.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.68, rely=0.54, width=50, height=50)
            ToolTip(self.warning_label,
                    "Показник середнього арифметичного перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")
        if median_warning:
            self.warning_label = tk.Label(self.output2.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.78, rely=0.39, width=50, height=50)
            ToolTip(self.warning_label,
                    "Показник медіани перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")
        if mode_warning:
            self.warning_label = tk.Label(self.output1.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.39, rely=0.39, width=50, height=50)
            ToolTip(self.warning_label,
                    "Показник моди перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")

    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
            f"{self.output1.cget('text')}",
            f"{self.output2.cget('text')}",
            f"{self.output3.cget('text')}",
            f"Файл даних: {self.master.selected_file if hasattr(self.master, 'selected_file') else 'Невідомо'}",
            "",
            ""
        ]

        try:
            with open("результат_аналізу.txt", "a", encoding="utf-8") as f:
                f.write("\n".join(result_lines))
        except Exception as e:
            print(f"Error loading file: {e}")

    def go_back(self):
        AnalysisWindow1(self.master.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)
class Analysis_KhNPP(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")

        self.thresholds = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 100,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 50,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 200,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 75,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 300,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 60,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 150,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 120,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 130,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 110,
            "Загальний об'єм водних скидів, куб. м": 5000,
            "Сумарний показник радіоактивних викидів, %": 80,
            "Сумарний індекс скиду радіоактивних речовин, %": 90
        }

        self.title("Аналіз для ХАЕС")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")

        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.18, width=800, height=440)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Аналіз для ХАЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        options = ["Радіоактивність викидів інертних радіоактивних газів, ГБк/доба",
                   "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %",
                   "Радіоактивність викидів радіонуклідів йоду, кБк/доба",
                   "Відсоток від допустимого рівня викидів радіонуклідів йоду, %",
                   "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба",
                   "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць",
                   "Загальний об'єм водних скидів, куб. м",
                   "Сумарний показник радіоактивних викидів, %",
                   "Сумарний індекс скиду радіоактивних речовин, %"]
        self.selected_option = tk.StringVar(value=options[0])
        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)

        self.output1 = tk.Label(self, text="Мода:", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output1.place(relx=0.15, rely=0.38, width=320, height=60)

        self.output2 = tk.Label(self, text="Медіана:", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output2.place(relx=0.53, rely=0.38, width=320, height=60)

        self.output3 = tk.Label(self, text="Середнє арифметичне:", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output3.place(relx=0.15, rely=0.53, width=700, height=60)

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        self.load_and_compute_modes()
        self.update_outputs(self.selected_option.get())

    def load_and_compute_modes(self):
        self.mode_results = {}
        self.median_results = {}
        self.mean_results = {}

        try:
            df = pd.read_excel(self.selected_file, engine='openpyxl')
            df = df[df['Назва станції'].str.strip() == "ХАЕС"]

            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
            df = df.infer_objects(copy=False)
            df.interpolate(method='linear', limit_direction='both', inplace=True)

            options = {
                "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
                "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
                "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
                "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
                "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
                "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
                "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
                "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
                "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
                "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
                "Загальний об'єм водних скидів, куб. м": 10,
                "Сумарний показник радіоактивних викидів, %": 11,
                "Сумарний індекс скиду радіоактивних речовин, %": 12
            }

            for option in options:
                if option in df.columns:
                    series = df[option]
                    if series.dtype == object:
                        series = series.astype(str).str.replace(",", ".", regex=False)
                        series = series.str.replace(r"[^\d\.]+", "", regex=True)
                        series = pd.to_numeric(series, errors='coerce')

                    series = series.dropna()
                    self.mode_results[option] = series.mode().tolist() or ["—"]
                    self.median_results[option] = series.median() if not series.empty else "—"
                    self.mean_results[option] = series.mean() if not series.empty else "—"
                else:
                    self.mode_results[option] = ["—"]
                    self.median_results[option] = "—"
                    self.mean_results[option] = "—"

        except Exception as e:
            messagebox.showerror("Помилка завантаження даних", f"Не вдалося обробити файл:\n{e}")

    def update_outputs(self, selection):
        mode = self.mode_results.get(selection, ["—"])
        median = self.median_results.get(selection, "—")
        mean = self.mean_results.get(selection, "—")

        if hasattr(self, 'warning_label'):
            self.warning_label.destroy()

        threshold = self.thresholds.get(selection)
        mean_warning = ""
        mode_warning = ""
        median_warning = ""

        if isinstance(mean, (float, int)) and threshold is not None and mean > threshold:
            mean_warning = " ⚠"
        if isinstance(mode, (float, int)) and threshold is not None and mode > threshold:
            mode_warning = " ⚠"
        if isinstance(median, (float, int)) and threshold is not None and median > threshold:
            median_warning = " ⚠"

        self.output1.config(text=f"Мода: {', '.join(map(str, mode)) if mode else '—'}{mode_warning}" if isinstance(mode,
                                                                                                                   (
                                                                                                                   float,
                                                                                                                   int)) else f"Мода: {', '.join(map(str, mode)) if mode else '—'}")
        self.output2.config(
            text=f"Медіана: {median}{median_warning}" if isinstance(median, (float, int)) else f"Медіана: {median}")
        self.output3.config(text=f"Середнє арифметичне: {mean:.3f}{mean_warning}" if isinstance(mean, (
        float, int)) else f"Середнє арифметичне: {mean}")

        if mean_warning:
            self.warning_label = tk.Label(self.output3.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.68, rely=0.54, width=50, height=50)
            ToolTip(self.warning_label,
                    "Показник середнього арифметичного перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")
        if median_warning:
            self.warning_label = tk.Label(self.output2.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.78, rely=0.39, width=50, height=50)
            ToolTip(self.warning_label,
                    "Показник медіани перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")
        if mode_warning:
            self.warning_label = tk.Label(self.output1.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.39, rely=0.39, width=50, height=50)
            ToolTip(self.warning_label,
                    "Показник моди перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")

    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
            f"{self.output1.cget('text')}",
            f"{self.output2.cget('text')}",
            f"{self.output3.cget('text')}",
            f"Файл даних: {self.master.selected_file if hasattr(self.master, 'selected_file') else 'Невідомо'}",
            "",
            ""
        ]

        try:
            with open("результат_аналізу.txt", "a", encoding="utf-8") as f:
                f.write("\n".join(result_lines))
        except Exception as e:
            print(f"Error loading file: {e}")

    def go_back(self):
        AnalysisWindow1(self.master.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)
class Analysis_RNPP(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")

        self.thresholds = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 100,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 50,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 200,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 75,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 300,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 60,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 150,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 120,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 130,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 110,
            "Загальний об'єм водних скидів, куб. м": 5000,
            "Сумарний показник радіоактивних викидів, %": 80,
            "Сумарний індекс скиду радіоактивних речовин, %": 90
        }

        self.title("Аналіз для РАЕС")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")

        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.18, width=800, height=440)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Аналіз для РАЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        options = ["Радіоактивність викидів інертних радіоактивних газів, ГБк/доба",
                   "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %",
                   "Радіоактивність викидів радіонуклідів йоду, кБк/доба",
                   "Відсоток від допустимого рівня викидів радіонуклідів йоду, %",
                   "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба",
                   "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць",
                   "Загальний об'єм водних скидів, куб. м",
                   "Сумарний показник радіоактивних викидів, %",
                   "Сумарний індекс скиду радіоактивних речовин, %"]
        self.selected_option = tk.StringVar(value=options[0])
        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)

        self.output1 = tk.Label(self, text="Мода:", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output1.place(relx=0.15, rely=0.38, width=320, height=60)

        self.output2 = tk.Label(self, text="Медіана:", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output2.place(relx=0.53, rely=0.38, width=320, height=60)

        self.output3 = tk.Label(self, text="Середнє арифметичне:", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output3.place(relx=0.15, rely=0.53, width=700, height=60)

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        self.load_and_compute_modes()
        self.update_outputs(self.selected_option.get())

    def load_and_compute_modes(self):
        self.mode_results = {}
        self.median_results = {}
        self.mean_results = {}

        try:
            df = pd.read_excel(self.selected_file, engine='openpyxl')
            df = df[df['Назва станції'].str.strip() == "РАЕС"]

            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
            df = df.infer_objects(copy=False)
            df.interpolate(method='linear', limit_direction='both', inplace=True)

            options = {
                "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
                "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
                "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
                "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
                "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
                "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
                "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
                "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
                "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
                "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
                "Загальний об'єм водних скидів, куб. м": 10,
                "Сумарний показник радіоактивних викидів, %": 11,
                "Сумарний індекс скиду радіоактивних речовин, %": 12
            }

            for option in options:
                if option in df.columns:
                    series = df[option]
                    if series.dtype == object:
                        series = series.astype(str).str.replace(",", ".", regex=False)
                        series = series.str.replace(r"[^\d\.]+", "", regex=True)
                        series = pd.to_numeric(series, errors='coerce')

                    series = series.dropna()
                    self.mode_results[option] = series.mode().tolist() or ["—"]
                    self.median_results[option] = series.median() if not series.empty else "—"
                    self.mean_results[option] = series.mean() if not series.empty else "—"
                else:
                    self.mode_results[option] = ["—"]
                    self.median_results[option] = "—"
                    self.mean_results[option] = "—"

        except Exception as e:
            messagebox.showerror("Помилка завантаження даних", f"Не вдалося обробити файл:\n{e}")

    def update_outputs(self, selection):
        mode = self.mode_results.get(selection, ["—"])
        median = self.median_results.get(selection, "—")
        mean = self.mean_results.get(selection, "—")

        if hasattr(self, 'warning_label'):
            self.warning_label.destroy()

        threshold = self.thresholds.get(selection)
        mean_warning = ""
        mode_warning = ""
        median_warning = ""

        if isinstance(mean, (float, int)) and threshold is not None and mean > threshold:
            mean_warning = " ⚠"
        if isinstance(mode, (float, int)) and threshold is not None and mode > threshold:
            mode_warning = " ⚠"
        if isinstance(median, (float, int)) and threshold is not None and median > threshold:
            median_warning = " ⚠"

        self.output1.config(text=f"Мода: {', '.join(map(str, mode)) if mode else '—'}{mode_warning}" if isinstance(mode,
                                                                                                                   (
                                                                                                                   float,
                                                                                                                   int)) else f"Мода: {', '.join(map(str, mode)) if mode else '—'}")
        self.output2.config(
            text=f"Медіана: {median}{median_warning}" if isinstance(median, (float, int)) else f"Медіана: {median}")
        self.output3.config(text=f"Середнє арифметичне: {mean:.3f}{mean_warning}" if isinstance(mean, (
        float, int)) else f"Середнє арифметичне: {mean}")

        if mean_warning:
            self.warning_label = tk.Label(self.output3.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.68, rely=0.54, width=50, height=50)
            ToolTip(self.warning_label,
                    "Показник середнього арифметичного перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")
        if median_warning:
            self.warning_label = tk.Label(self.output2.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.78, rely=0.39, width=50, height=50)
            ToolTip(self.warning_label,
                    "Показник медіани перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")
        if mode_warning:
            self.warning_label = tk.Label(self.output1.master, bg="#FFFFFF", cursor="question_arrow")
            self.warning_label.place(relx=0.39, rely=0.39, width=50, height=50)
            ToolTip(self.warning_label,
                    "Показник моди перевищує допустимий ліміт!\nЦе може свідчити про потенційну загрозу.\nПотребує консультації з фахівцем.")

    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
            f"{self.output1.cget('text')}",
            f"{self.output2.cget('text')}",
            f"{self.output3.cget('text')}",
            f"Файл даних: {self.master.selected_file if hasattr(self.master, 'selected_file') else 'Невідомо'}",
            "",
            ""
        ]

        try:
            with open("результат_аналізу.txt", "a", encoding="utf-8") as f:
                f.write("\n".join(result_lines))
        except Exception as e:
            print(f"Error loading file: {e}")

    def go_back(self):
        AnalysisWindow1(self.master.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)

class ARIMA_ZAG(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.master = master
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")

        self.thresholds = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 100,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 50,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 200,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 75,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 300,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 60,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 150,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 120,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 130,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 110,
            "Загальний об'єм водних скидів, куб. м": 5000,
            "Сумарний показник радіоактивних викидів, %": 80,
            "Сумарний індекс скиду радіоактивних речовин, %": 90
        }
        self.title("ARIMA передбачення для всіх АЕС")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")
        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.18, width=800, height=440)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Прогноз для всіх АЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        options = ["Радіоактивність викидів інертних радіоактивних газів, ГБк/доба",
                   "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %",
                   "Радіоактивність викидів радіонуклідів йоду, кБк/доба",
                   "Відсоток від допустимого рівня викидів радіонуклідів йоду, %",
                   "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба",
                   "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць",
                   "Загальний об'єм водних скидів, куб. м",
                   "Сумарний показник радіоактивних викидів, %",
                   "Сумарний індекс скиду радіоактивних речовин, %"]
        self.selected_option = tk.StringVar(value=options[0])
        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)

        self.output1 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output1.place(relx=0.15, rely=0.38, width=320, height=30)

        self.output2 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output2.place(relx=0.53, rely=0.38, width=320, height=30)

        self.output3 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output3.place(relx=0.15, rely=0.46, width=320, height=30)

        self.output4 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output4.place(relx=0.53, rely=0.46, width=320, height=30)

        self.output5 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output5.place(relx=0.15, rely=0.54, width=320, height=30)

        self.output6 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output6.place(relx=0.53, rely=0.54, width=320, height=30)

        self.output7 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output7.place(relx=0.15, rely=0.62, width=320, height=30)

        self.output8 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output8.place(relx=0.53, rely=0.62, width=320, height=30)

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        self.load_and_compute_ARIMA()
        self.update_outputs(self.selected_option.get())

    def load_and_compute_ARIMA(self):
        self.forecast = {}

        try:
            df = pd.read_excel(self.selected_file, engine='openpyxl')

            df.index.freq = 'QE-DEC'
            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
            df = df.infer_objects(copy=False)
            df.interpolate(method='linear', limit_direction='both', inplace=True)

            df['date'] = pd.PeriodIndex.from_fields(
                year=df['Рік'],
                quarter=df['Квартал'],
                freq='Q-DEC'
            ).to_timestamp()
            df.set_index('date', inplace=True)

            options = {
                "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
                "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
                "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
                "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
                "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
                "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
                "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
                "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
                "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
                "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
                "Загальний об'єм водних скидів, куб. м": 10,
                "Сумарний показник радіоактивних викидів, %": 11,
                "Сумарний індекс скиду радіоактивних речовин, %": 12
            }


            for option in options:
                if option in df.columns:
                    ts = df[option]

                    if ts.dtype == object or not pd.api.types.is_numeric_dtype(ts):
                        ts = ts.astype(str).str.replace(",", ".", regex=False)
                        ts = ts.str.replace(r"[^\d\.\-]+", "", regex=True)
                        ts = pd.to_numeric(ts, errors='coerce')

                    ts.dropna(inplace=True)

                    try:
                        model = ARIMA(ts, order=(5, 1, 0))
                        model_fit = model.fit()

                        forecast = model_fit.forecast(steps=8)
                        forecast_index = pd.date_range(
                            start=ts.index[-1] + pd.offsets.QuarterEnd(),
                            periods=8,
                            freq='Q'
                        )

                        forecast = pd.Series(forecast.values, index=forecast_index)

                        plt.figure(figsize=(10, 5))
                        plt.plot(ts, label="Історичні дані")
                        plt.plot(forecast.index, forecast, label="Прогноз", color="red")
                        plt.legend()
                        plt.title(f'Прогноз ARIMA для "{option}"')
                        plt.show()

                    except Exception as model_error:
                        print(f"Помилка обробки ARIMA для '{option}': {model_error}")

        except Exception as e:
            print("Помилка завантаження даних:", f"\n{e}")

    def update_outputs(self, selection):
        forecast = self.forecast.get(selection, ["—"] * 8)

        self.output1.config(text=f"Прогноз 1: {forecast[0]}")
        self.output2.config(text=f"Прогноз 2: {forecast[1]}")
        self.output3.config(text=f"Прогноз 3: {forecast[2]}")
        self.output4.config(text=f"Прогноз 4: {forecast[3]}")
        self.output5.config(text=f"Прогноз 5: {forecast[4]}")
        self.output6.config(text=f"Прогноз 6: {forecast[5]}")
        self.output7.config(text=f"Прогноз 7: {forecast[6]}")
        self.output8.config(text=f"Прогноз 8: {forecast[7]}")

    def go_back(self):
        AnalysisWindow3(self.master.master, self.selected_file)
        self.withdraw()
    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
            f"{self.output1.cget('text')}",
            f"{self.output2.cget('text')}",
            f"{self.output3.cget('text')}",
            f"{self.output4.cget('text')}",
            f"{self.output5.cget('text')}",
            f"{self.output6.cget('text')}",
            f"{self.output7.cget('text')}",
            f"{self.output8.cget('text')}",
            f"Файл даних: {self.master.selected_file if hasattr(self.master, 'selected_file') else 'Невідомо'}",
            "",
            ""
        ]

        try:
            with open("результат_аналізу.txt", "a", encoding="utf-8") as f:
                f.write("\n".join(result_lines))
        except Exception as e:
            print(f"Error loading file: {e}")

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)
class ARIMA_ZNPP(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.master = master
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")

        self.thresholds = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 100,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 50,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 200,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 75,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 300,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 60,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 150,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 120,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 130,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 110,
            "Загальний об'єм водних скидів, куб. м": 5000,
            "Сумарний показник радіоактивних викидів, %": 80,
            "Сумарний індекс скиду радіоактивних речовин, %": 90
        }
        self.title("ARIMA передбачення для ЗАЕС")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")
        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.18, width=800, height=440)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Прогноз для ЗАЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        options = ["Радіоактивність викидів інертних радіоактивних газів, ГБк/доба",
                   "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %",
                   "Радіоактивність викидів радіонуклідів йоду, кБк/доба",
                   "Відсоток від допустимого рівня викидів радіонуклідів йоду, %",
                   "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба",
                   "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць",
                   "Загальний об'єм водних скидів, куб. м",
                   "Сумарний показник радіоактивних викидів, %",
                   "Сумарний індекс скиду радіоактивних речовин, %"]
        self.selected_option = tk.StringVar(value=options[0])
        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)

        self.output1 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output1.place(relx=0.15, rely=0.38, width=320, height=30)

        self.output2 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output2.place(relx=0.53, rely=0.38, width=320, height=30)

        self.output3 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output3.place(relx=0.15, rely=0.46, width=320, height=30)

        self.output4 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output4.place(relx=0.53, rely=0.46, width=320, height=30)

        self.output5 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output5.place(relx=0.15, rely=0.54, width=320, height=30)

        self.output6 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output6.place(relx=0.53, rely=0.54, width=320, height=30)

        self.output7 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output7.place(relx=0.15, rely=0.62, width=320, height=30)

        self.output8 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output8.place(relx=0.53, rely=0.62, width=320, height=30)

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        self.load_and_compute_ARIMA()
        self.update_outputs(self.selected_option.get())

    def load_and_compute_ARIMA(self):
        self.forecast = {}

        try:
            df = pd.read_excel(self.selected_file, engine='openpyxl')
            df = df[df['Назва станції'].str.strip() == "ЗАЕС"]

            df.index.freq = 'QE-DEC'
            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
            df = df.infer_objects(copy=False)
            df.interpolate(method='linear', limit_direction='both', inplace=True)

            df['date'] = pd.PeriodIndex.from_fields(
                year=df['Рік'],
                quarter=df['Квартал'],
                freq='Q-DEC'
            ).to_timestamp()
            df.set_index('date', inplace=True)

            options = {
                "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
                "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
                "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
                "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
                "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
                "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
                "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
                "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
                "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
                "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
                "Загальний об'єм водних скидів, куб. м": 10,
                "Сумарний показник радіоактивних викидів, %": 11,
                "Сумарний індекс скиду радіоактивних речовин, %": 12
            }

            for option in options:
                if option in df.columns:
                    ts = df[option]

                    if ts.dtype == object or not pd.api.types.is_numeric_dtype(ts):
                        ts = ts.astype(str).str.replace(",", ".", regex=False)
                        ts = ts.str.replace(r"[^\d\.\-]+", "", regex=True)
                        ts = pd.to_numeric(ts, errors='coerce')

                    ts.dropna(inplace=True)

                    try:
                        model = ARIMA(ts, order=(5, 1, 0))
                        model_fit = model.fit()

                        forecast = model_fit.forecast(steps=8)
                        forecast_index = pd.date_range(
                            start=ts.index[-1] + pd.offsets.QuarterEnd(),
                            periods=8,
                            freq='Q'
                        )

                        forecast = pd.Series(forecast.values, index=forecast_index)
                        self.forecast[option] = {
                            'ts': ts,
                            'forecast': forecast
                        }



                    except Exception as model_error:
                        print(f" Помилка обробки ARIMA для '{option}': {model_error}")

        except Exception as e:
            print("Помилка завантаження даних:", f"\n{e}")

    def update_outputs(self, selection):
        data = self.forecast.get(selection)

        if data is not None:
            forecast = data['forecast']
            ts = data['ts']

            values = forecast.tolist()
            for i in range(8):
                getattr(self, f"output{i + 1}").config(text=f"Прогноз {i + 1}: {values[i]:.2f}")

            plt.figure(figsize=(10, 5))
            plt.plot(ts, label="Історичні дані")
            plt.plot(forecast.index, forecast, label="Прогноз", color="red")
            plt.legend()
            plt.title(f'Прогноз ARIMA для "{selection}"')
            plt.show()
        else:
            for i in range(8):
                getattr(self, f"output{i + 1}").config(text=f"Прогноз {i + 1}: —")

    def go_back(self):
        AnalysisWindow3(self.master.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
            f"{self.output1.cget('text')}",
            f"{self.output2.cget('text')}",
            f"{self.output3.cget('text')}",
            f"{self.output4.cget('text')}",
            f"{self.output5.cget('text')}",
            f"{self.output6.cget('text')}",
            f"{self.output7.cget('text')}",
            f"{self.output8.cget('text')}",
            f"Файл даних: {self.master.selected_file if hasattr(self.master, 'selected_file') else 'Невідомо'}",
            "",
            ""
        ]

        try:
            with open("результат_аналізу.txt", "a", encoding="utf-8") as f:
                f.write("\n".join(result_lines))
        except Exception as e:
            print(f"Error loading file: {e}")

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)
class ARIMA_PNPP(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.master = master
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")

        self.thresholds = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 100,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 50,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 200,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 75,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 300,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 60,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 150,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 120,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 130,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 110,
            "Загальний об'єм водних скидів, куб. м": 5000,
            "Сумарний показник радіоактивних викидів, %": 80,
            "Сумарний індекс скиду радіоактивних речовин, %": 90
        }
        self.title("ARIMA передбачення для ПАЕС")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")
        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.18, width=800, height=440)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Прогноз для ПАЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        options = ["Радіоактивність викидів інертних радіоактивних газів, ГБк/доба",
                   "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %",
                   "Радіоактивність викидів радіонуклідів йоду, кБк/доба",
                   "Відсоток від допустимого рівня викидів радіонуклідів йоду, %",
                   "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба",
                   "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць",
                   "Загальний об'єм водних скидів, куб. м",
                   "Сумарний показник радіоактивних викидів, %",
                   "Сумарний індекс скиду радіоактивних речовин, %"]
        self.selected_option = tk.StringVar(value=options[0])
        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)

        self.output1 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output1.place(relx=0.15, rely=0.38, width=320, height=30)

        self.output2 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output2.place(relx=0.53, rely=0.38, width=320, height=30)

        self.output3 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output3.place(relx=0.15, rely=0.46, width=320, height=30)

        self.output4 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output4.place(relx=0.53, rely=0.46, width=320, height=30)

        self.output5 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output5.place(relx=0.15, rely=0.54, width=320, height=30)

        self.output6 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output6.place(relx=0.53, rely=0.54, width=320, height=30)

        self.output7 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output7.place(relx=0.15, rely=0.62, width=320, height=30)

        self.output8 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output8.place(relx=0.53, rely=0.62, width=320, height=30)

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        self.load_and_compute_ARIMA()
        self.update_outputs(self.selected_option.get())

    def load_and_compute_ARIMA(self):
        self.forecast = {}

        try:
            df = pd.read_excel(self.selected_file, engine='openpyxl')
            df = df[df['Назва станції'].str.strip() == "ПАЕС"]

            df.index.freq = 'QE-DEC'
            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
            df = df.infer_objects(copy=False)
            df.interpolate(method='linear', limit_direction='both', inplace=True)

            df['date'] = pd.PeriodIndex.from_fields(
                year=df['Рік'],
                quarter=df['Квартал'],
                freq='Q-DEC'
            ).to_timestamp()
            df.set_index('date', inplace=True)

            options = {
                "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
                "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
                "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
                "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
                "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
                "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
                "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
                "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
                "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
                "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
                "Загальний об'єм водних скидів, куб. м": 10,
                "Сумарний показник радіоактивних викидів, %": 11,
                "Сумарний індекс скиду радіоактивних речовин, %": 12
            }

            for option in options:
                if option in df.columns:
                    ts = df[option]

                    if ts.dtype == object or not pd.api.types.is_numeric_dtype(ts):
                        ts = ts.astype(str).str.replace(",", ".", regex=False)
                        ts = ts.str.replace(r"[^\d\.\-]+", "", regex=True)
                        ts = pd.to_numeric(ts, errors='coerce')

                    ts.dropna(inplace=True)

                    try:
                        model = ARIMA(ts, order=(5, 1, 0))
                        model_fit = model.fit()

                        forecast = model_fit.forecast(steps=8)
                        forecast_index = pd.date_range(
                            start=ts.index[-1] + pd.offsets.QuarterEnd(),
                            periods=8,
                            freq='Q'
                        )

                        forecast = pd.Series(forecast.values, index=forecast_index)
                        self.forecast[option] = {
                            'ts': ts,
                            'forecast': forecast
                        }



                    except Exception as model_error:
                        print(f" Помилка обробки ARIMA для '{option}': {model_error}")

        except Exception as e:
            print("Помилка завантаження даних:", f"\n{e}")

    def update_outputs(self, selection):
        data = self.forecast.get(selection)

        if data is not None:
            forecast = data['forecast']
            ts = data['ts']

            values = forecast.tolist()
            for i in range(8):
                getattr(self, f"output{i + 1}").config(text=f"Прогноз {i + 1}: {values[i]:.2f}")

            plt.figure(figsize=(10, 5))
            plt.plot(ts, label="Історичні дані")
            plt.plot(forecast.index, forecast, label="Прогноз", color="red")
            plt.legend()
            plt.title(f'Прогноз ARIMA для "{selection}"')
            plt.show()
        else:
            for i in range(8):
                getattr(self, f"output{i + 1}").config(text=f"Прогноз {i + 1}: —")

    def go_back(self):
        AnalysisWindow3(self.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
            f"{self.output1.cget('text')}",
            f"{self.output2.cget('text')}",
            f"{self.output3.cget('text')}",
            f"{self.output4.cget('text')}",
            f"{self.output5.cget('text')}",
            f"{self.output6.cget('text')}",
            f"{self.output7.cget('text')}",
            f"{self.output8.cget('text')}",
            f"Файл даних: {self.master.selected_file if hasattr(self.master, 'selected_file') else 'Невідомо'}",
            "",
            ""
        ]

        try:
            with open("результат_аналізу.txt", "a", encoding="utf-8") as f:
                f.write("\n".join(result_lines))
        except Exception as e:
            print(f"Error loading file: {e}")

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)
class ARIMA_KhNPP(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.master = master
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")

        self.thresholds = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 100,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 50,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 200,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 75,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 300,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 60,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 150,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 120,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 130,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 110,
            "Загальний об'єм водних скидів, куб. м": 5000,
            "Сумарний показник радіоактивних викидів, %": 80,
            "Сумарний індекс скиду радіоактивних речовин, %": 90
        }
        self.title("ARIMA передбачення для ХАЕС")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")
        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.18, width=800, height=440)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Прогноз для ХАЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        options = ["Радіоактивність викидів інертних радіоактивних газів, ГБк/доба",
                   "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %",
                   "Радіоактивність викидів радіонуклідів йоду, кБк/доба",
                   "Відсоток від допустимого рівня викидів радіонуклідів йоду, %",
                   "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба",
                   "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць",
                   "Загальний об'єм водних скидів, куб. м",
                   "Сумарний показник радіоактивних викидів, %",
                   "Сумарний індекс скиду радіоактивних речовин, %"]
        self.selected_option = tk.StringVar(value=options[0])
        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)


        self.output1 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output1.place(relx=0.15, rely=0.38, width=320, height=30)

        self.output2 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output2.place(relx=0.53, rely=0.38, width=320, height=30)

        self.output3 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output3.place(relx=0.15, rely=0.46, width=320, height=30)

        self.output4 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output4.place(relx=0.53, rely=0.46, width=320, height=30)

        self.output5 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output5.place(relx=0.15, rely=0.54, width=320, height=30)

        self.output6 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output6.place(relx=0.53, rely=0.54, width=320, height=30)

        self.output7 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output7.place(relx=0.15, rely=0.62, width=320, height=30)

        self.output8 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output8.place(relx=0.53, rely=0.62, width=320, height=30)




        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        self.load_and_compute_ARIMA()
        self.update_outputs(self.selected_option.get())

    def load_and_compute_ARIMA(self):
        self.forecast = {}

        try:
            df = pd.read_excel(self.selected_file, engine='openpyxl')
            df = df[df['Назва станції'].str.strip() == "ХАЕС"]

            df.index.freq = 'QE-DEC'
            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
            df = df.infer_objects(copy=False)
            df.interpolate(method='linear', limit_direction='both', inplace=True)

            df['date'] = pd.PeriodIndex.from_fields(
                year=df['Рік'],
                quarter=df['Квартал'],
                freq='Q-DEC'
            ).to_timestamp()
            df.set_index('date', inplace=True)

            options = {
                "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
                "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
                "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
                "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
                "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
                "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
                "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
                "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
                "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
                "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
                "Загальний об'єм водних скидів, куб. м": 10,
                "Сумарний показник радіоактивних викидів, %": 11,
                "Сумарний індекс скиду радіоактивних речовин, %": 12
            }

            for option in options:
                if option in df.columns:
                    ts = df[option]

                    if ts.dtype == object or not pd.api.types.is_numeric_dtype(ts):
                        ts = ts.astype(str).str.replace(",", ".", regex=False)
                        ts = ts.str.replace(r"[^\d\.\-]+", "", regex=True)
                        ts = pd.to_numeric(ts, errors='coerce')

                    ts.dropna(inplace=True)

                    try:
                        model = ARIMA(ts, order=(5, 1, 0))
                        model_fit = model.fit()

                        forecast = model_fit.forecast(steps=8)
                        forecast_index = pd.date_range(
                            start=ts.index[-1] + pd.offsets.QuarterEnd(),
                            periods=8,
                            freq='Q'
                        )

                        forecast = pd.Series(forecast.values, index=forecast_index)
                        self.forecast[option] = {
                            'ts': ts,
                            'forecast': forecast
                        }



                    except Exception as model_error:
                        print(f" Помилка обробки ARIMA для '{option}': {model_error}")

        except Exception as e:
            print("Помилка завантаження даних:", f"\n{e}")

    def update_outputs(self, selection):
        data = self.forecast.get(selection)

        if data is not None:
            forecast = data['forecast']
            ts = data['ts']

            values = forecast.tolist()
            for i in range(8):
                getattr(self, f"output{i + 1}").config(text=f"Прогноз {i + 1}: {values[i]:.2f}")

            plt.figure(figsize=(10, 5))
            plt.plot(ts, label="Історичні дані")
            plt.plot(forecast.index, forecast, label="Прогноз", color="red")
            plt.legend()
            plt.title(f'Прогноз ARIMA для "{selection}"')
            plt.show()
        else:
            for i in range(8):
                getattr(self, f"output{i + 1}").config(text=f"Прогноз {i + 1}: —")

    def go_back(self):
        AnalysisWindow3(self.master.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
            f"{self.output1.cget('text')}",
            f"{self.output2.cget('text')}",
            f"{self.output3.cget('text')}",
            f"{self.output4.cget('text')}",
            f"{self.output5.cget('text')}",
            f"{self.output6.cget('text')}",
            f"{self.output7.cget('text')}",
            f"{self.output8.cget('text')}",
            f"Файл даних: {self.master.selected_file if hasattr(self.master, 'selected_file') else 'Невідомо'}",
            "",
            ""
        ]

        try:
            with open("результат_аналізу.txt", "a", encoding="utf-8") as f:
                f.write("\n".join(result_lines))
        except Exception as e:
            print(f"Error loading file: {e}")

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)
class ARIMA_RNPP(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.master = master
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")

        self.thresholds = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 100,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 50,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 200,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 75,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 300,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 60,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 150,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 120,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 130,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 110,
            "Загальний об'єм водних скидів, куб. м": 5000,
            "Сумарний показник радіоактивних викидів, %": 80,
            "Сумарний індекс скиду радіоактивних речовин, %": 90
        }
        self.title("ARIMA передбачення для РАЕС")
        self.geometry("1000x600")
        self.resizable(False, False)
        self.configure(bg="#D9D9D9")
        self.label8 = tk.Label(self, text="")
        self.label8.configure(bg="#5B8386", fg="#FFFFFF", anchor="center")
        self.label8.place(relx=0.12, rely=0.18, width=800, height=440)

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Прогноз для РАЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        options = ["Радіоактивність викидів інертних радіоактивних газів, ГБк/доба",
                   "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %",
                   "Радіоактивність викидів радіонуклідів йоду, кБк/доба",
                   "Відсоток від допустимого рівня викидів радіонуклідів йоду, %",
                   "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба",
                   "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць",
                   "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць",
                   "Загальний об'єм водних скидів, куб. м",
                   "Сумарний показник радіоактивних викидів, %",
                   "Сумарний індекс скиду радіоактивних речовин, %"]
        self.selected_option = tk.StringVar(value=options[0])
        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)

        self.output1 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output1.place(relx=0.15, rely=0.38, width=320, height=30)

        self.output2 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                                highlightbackground="black", highlightthickness=1)
        self.output2.place(relx=0.53, rely=0.38, width=320, height=30)

        self.output3 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output3.place(relx=0.15, rely=0.46, width=320, height=30)

        self.output4 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output4.place(relx=0.53, rely=0.46, width=320, height=30)

        self.output5 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output5.place(relx=0.15, rely=0.54, width=320, height=30)

        self.output6 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output6.place(relx=0.53, rely=0.54, width=320, height=30)

        self.output7 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output7.place(relx=0.15, rely=0.62, width=320, height=30)

        self.output8 = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF",
                                anchor="center", highlightbackground="black", highlightthickness=1)
        self.output8.place(relx=0.53, rely=0.62, width=320, height=30)

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        self.load_and_compute_ARIMA()
        self.update_outputs(self.selected_option.get())

    def load_and_compute_ARIMA(self):
        self.forecast = {}

        try:
            df = pd.read_excel(self.selected_file, engine='openpyxl')
            df = df[df['Назва станції'].str.strip() == "РАЕС"]

            df.index.freq = 'QE-DEC'
            df.dropna(how='all', inplace=True)
            df.reset_index(drop=True, inplace=True)
            df = df.infer_objects(copy=False)
            df.interpolate(method='linear', limit_direction='both', inplace=True)

            df['date'] = pd.PeriodIndex.from_fields(
                year=df['Рік'],
                quarter=df['Квартал'],
                freq='Q-DEC'
            ).to_timestamp()
            df.set_index('date', inplace=True)

            options = {
                "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
                "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
                "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
                "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
                "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
                "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
                "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
                "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
                "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
                "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
                "Загальний об'єм водних скидів, куб. м": 10,
                "Сумарний показник радіоактивних викидів, %": 11,
                "Сумарний індекс скиду радіоактивних речовин, %": 12
            }

            for option in options:
                if option in df.columns:
                    ts = df[option]

                    if ts.dtype == object or not pd.api.types.is_numeric_dtype(ts):
                        ts = ts.astype(str).str.replace(",", ".", regex=False)
                        ts = ts.str.replace(r"[^\d\.\-]+", "", regex=True)
                        ts = pd.to_numeric(ts, errors='coerce')

                    ts.dropna(inplace=True)

                    try:
                        model = ARIMA(ts, order=(5, 1, 0))
                        model_fit = model.fit()

                        forecast = model_fit.forecast(steps=8)
                        forecast_index = pd.date_range(
                            start=ts.index[-1] + pd.offsets.QuarterEnd(),
                            periods=8,
                            freq='Q'
                        )

                        forecast = pd.Series(forecast.values, index=forecast_index)
                        self.forecast[option] = {
                            'ts': ts,
                            'forecast': forecast
                        }



                    except Exception as model_error:
                        print(f" Помилка обробки ARIMA для '{option}': {model_error}")

        except Exception as e:
            print("Помилка завантаження даних:", f"\n{e}")

    def update_outputs(self, selection):
        data = self.forecast.get(selection)

        if data is not None:
            forecast = data['forecast']
            ts = data['ts']

            values = forecast.tolist()
            for i in range(8):
                getattr(self, f"output{i + 1}").config(text=f"Прогноз {i + 1}: {values[i]:.2f}")

            plt.figure(figsize=(10, 5))
            plt.plot(ts, label="Історичні дані")
            plt.plot(forecast.index, forecast, label="Прогноз", color="red")
            plt.legend()
            plt.title(f'Прогноз ARIMA для "{selection}"')
            plt.show()
        else:
            for i in range(8):
                getattr(self, f"output{i + 1}").config(text=f"Прогноз {i + 1}: —")

    def go_back(self):
        AnalysisWindow3(self.master.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
            f"{self.output1.cget('text')}",
            f"{self.output2.cget('text')}",
            f"{self.output3.cget('text')}",
            f"{self.output4.cget('text')}",
            f"{self.output5.cget('text')}",
            f"{self.output6.cget('text')}",
            f"{self.output7.cget('text')}",
            f"{self.output8.cget('text')}",
            f"Файл даних: {self.master.selected_file if hasattr(self.master, 'selected_file') else 'Невідомо'}",
            "",
            ""
        ]

        try:
            with open("результат_аналізу.txt", "a", encoding="utf-8") as f:
                f.write("\n".join(result_lines))
        except Exception as e:
            print(f"Error loading file: {e}")

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)


class TemporalBlock(nn.Module):
    def __init__(self, in_channels, out_channels, kernel_size, dilation):
        super().__init__()
        padding = dilation * (kernel_size - 1)
        self.net = nn.Sequential(
            nn.Conv1d(in_channels, out_channels, kernel_size,
                      padding=padding, dilation=dilation),
            nn.ReLU(),
            nn.Conv1d(out_channels, out_channels, kernel_size,
                      padding=padding, dilation=dilation),
            nn.ReLU()
        )
        self.downsample = nn.Conv1d(in_channels, out_channels, 1) if in_channels != out_channels else None

    def forward(self, x):
        out = self.net(x)
        out = out[:, :, :x.size(2)]
        res = x if self.downsample is None else self.downsample(x)
        return out + res

class TCN(nn.Module):
    def __init__(self, seq_length, input_channels=1, n_outputs=1):
        super().__init__()
        self.tcn = nn.Sequential(
            TemporalBlock(input_channels, 16, kernel_size=3, dilation=1),
            TemporalBlock(16, 32, kernel_size=3, dilation=2),
            nn.Flatten()
        )
        self.linear = nn.Linear(32 * seq_length, n_outputs)

    def forward(self, x):
        x = x.transpose(1, 2)
        x = self.tcn(x)
        return self.linear(x)



class TCN_ZAG(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.master = master
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")
        self.seq_length = 8
        self.scaler = StandardScaler()
        self.forecast = {}

        options = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
            "Загальний об'єм водних скидів, куб. м": 10,
            "Сумарний показник радіоактивних викидів, %": 11,
            "Сумарний індекс скиду радіоактивних речовин, %": 12
        }



        self.setup_gui(options)
        self.load_and_train()
        self.update_outputs(self.selected_option.get())

    def setup_gui(self, options):
        self.title("TCN прогноз для ХАЕС")
        self.geometry("1000x600")
        self.configure(bg="#D9D9D9")
        self.options = options

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Прогноз для ХАЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)
        self.outputs = []

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        for i in range(self.seq_length):
            lbl = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                           highlightbackground="black", highlightthickness=1)
            col = 0.15 if i % 2 == 0 else 0.53
            row = 0.38 + (i // 2) * 0.08
            lbl.place(relx=col, rely=row, width=320, height=30)
            self.outputs.append(lbl)

    def load_and_train(self):
        df = pd.read_excel(self.selected_file, engine='openpyxl')
        df.columns = df.columns.str.strip().str.replace(r"[\r\n\t]", "", regex=True)

        df = df.dropna(subset=['Рік', 'Квартал'])

        df['Рік'] = df['Рік'].astype(int)
        df['Квартал'] = df['Квартал'].astype(int)

        df['date'] = pd.PeriodIndex.from_fields(year=df['Рік'], quarter=df['Квартал'], freq='Q-DEC').to_timestamp()

        exclude_cols = ['Рік', 'Квартал', 'date']
        for col in df.columns:
            if col not in exclude_cols:
                df[col] = df[col].astype(str).str.replace(",", ".", regex=False)  # заміна ком
                df[col] = df[col].str.replace(r"[^\d\.\-]+", "", regex=True)  # видалення одиниць виміру, тексту
                df[col] = pd.to_numeric(df[col], errors='coerce')  # конвертація в float

        df.set_index('date', inplace=True)
        df = df.groupby('date').mean(numeric_only=True)


        for col in self.options:

            try:
                ts = df[col]
                if ts.dtype == object or not pd.api.types.is_numeric_dtype(ts):
                    ts = ts.astype(str).str.replace(",", ".", regex=False)
                    ts = ts.str.replace(r"[^\d\.\-]+", "", regex=True)
                    ts = pd.to_numeric(ts, errors='coerce')


                ts.astype(float).interpolate()



                data = self.scaler.fit_transform(ts.values.reshape(-1, 1)).flatten()
                X, y = self.create_sequences(data)
                X = X[:, :, np.newaxis].astype(np.float32)
                y = y.astype(np.float32).reshape(-1, 1)


                net = NeuralNetRegressor(TCN, module__seq_length=self.seq_length, max_epochs=100, lr=1e-3,
                                         batch_size=8, optimizer=torch.optim.Adam, iterator_train__shuffle=True,
                                         verbose=0)
                net.fit(X, y)

                preds = self.forecast_future(net, data[-self.seq_length:], 8)
                future = self.scaler.inverse_transform(preds.reshape(-1, 1)).flatten()
                idx = pd.date_range(start=ts.index[-1] + pd.offsets.QuarterEnd(), periods=8, freq='Q')
                future_ts = pd.Series(future, index=idx)

                self.forecast[col] = {'ts': ts, 'forecast': future_ts}
            except Exception as e:
                print(f"Проблема з колонкою '{col}': {e}")

    def create_sequences(self, data):
        X, y = [], []
        for i in range(len(data) - self.seq_length):
            X.append(data[i:i + self.seq_length])
            y.append(data[i + self.seq_length])
        return np.array(X), np.array(y)

    def forecast_future(self, net, last_seq, n_steps):
        seq = last_seq.copy()
        preds = []
        for _ in range(n_steps):
            x = torch.tensor(seq[np.newaxis, :, np.newaxis], dtype=torch.float32)
            with torch.no_grad():
                p = net.predict(x.numpy())[0]
            preds.append(p)
            seq = np.roll(seq, -1)
            seq[-1] = p
        return np.array(preds)

    def update_outputs(self, selection):
        data = self.forecast.get(selection)

        if data:
            vals = data['forecast'].tolist()
            for i, lbl in enumerate(self.outputs):
                lbl.config(text=f"Прогноз {i + 1}: {vals[i]:.2f}")
            plt.figure(figsize=(10, 5))
            plt.plot(data['ts'], label="Історичні дані")
            plt.plot(data['forecast'], label="Прогноз TCN", color='red')
            plt.legend()
            plt.title(f'TCN прогноз для "{selection}"')
            plt.show()
        else:
            for lbl in self.outputs:
                lbl.config(text="—")

    def go_back(self):
        AnalysisWindow2(self.master.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
        ]

        # Додати всі прогнози з self.outputs
        for i, lbl in enumerate(self.outputs):
            result_lines.append(f"{lbl.cget('text')}")

        # Додати шлях до файлу
        result_lines.append(f"Файл даних: {self.selected_file if hasattr(self, 'selected_file') else 'Невідомо'}")
        result_lines.append("")
        result_lines.append("")

        # Зберігаємо у файл
        filename = "результат_аналізу.txt"
        with open(filename, "a", encoding="utf-8") as f:
            f.write("\n".join(result_lines))

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)


class TCN_ZNPP(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.master = master
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")
        self.seq_length = 8
        self.scaler = StandardScaler()
        self.forecast = {}

        options = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
            "Загальний об'єм водних скидів, куб. м": 10,
            "Сумарний показник радіоактивних викидів, %": 11,
            "Сумарний індекс скиду радіоактивних речовин, %": 12
        }

        self.setup_gui(options)
        self.load_and_train()
        self.update_outputs(self.selected_option.get())

    def setup_gui(self, options):
        self.title("TCN прогноз для ЗАЕС")
        self.geometry("1000x600")
        self.configure(bg="#D9D9D9")
        self.options = options

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Прогноз для ЗАЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)
        self.outputs = []

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        for i in range(self.seq_length):
            lbl = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                           highlightbackground="black", highlightthickness=1)
            col = 0.15 if i % 2 == 0 else 0.53
            row = 0.38 + (i // 2) * 0.08
            lbl.place(relx=col, rely=row, width=320, height=30)
            self.outputs.append(lbl)

    def load_and_train(self):
        df = pd.read_excel(self.selected_file, engine='openpyxl')
        df = df[df['Назва станції'].str.strip() == "ЗАЕС"].copy()
        df['date'] = pd.PeriodIndex.from_fields(year=df['Рік'], quarter=df['Квартал'], freq='Q-DEC').to_timestamp()
        df.set_index('date', inplace=True)

        for col in self.options:
            try:
                ts = df[col]
                if ts.dtype == object or not pd.api.types.is_numeric_dtype(ts):
                    ts = ts.astype(str).str.replace(",", ".", regex=False)
                    ts = ts.str.replace(r"[^\d\.\-]+", "", regex=True)
                    ts = pd.to_numeric(ts, errors='coerce')

                ts.dropna(inplace=True)
                ts.astype(float).interpolate().dropna()

                data = self.scaler.fit_transform(ts.values.reshape(-1, 1)).flatten()
                X, y = self.create_sequences(data)
                X = X[:, :, np.newaxis].astype(np.float32)
                y = y.astype(np.float32).reshape(-1, 1)

                net = NeuralNetRegressor(TCN, module__seq_length=self.seq_length, max_epochs=100, lr=1e-3,
                                         batch_size=8, optimizer=torch.optim.Adam, iterator_train__shuffle=True,
                                         verbose=0)
                net.fit(X, y)

                preds = self.forecast_future(net, data[-self.seq_length:], 8)
                future = self.scaler.inverse_transform(preds.reshape(-1, 1)).flatten()
                idx = pd.date_range(start=ts.index[-1] + pd.offsets.QuarterEnd(), periods=8, freq='Q')
                future_ts = pd.Series(future, index=idx)

                self.forecast[col] = {'ts': ts, 'forecast': future_ts}
            except Exception as e:
                print(f"Проблема з колонкою '{col}': {e}")

    def create_sequences(self, data):
        X, y = [], []
        for i in range(len(data) - self.seq_length):
            X.append(data[i:i + self.seq_length])
            y.append(data[i + self.seq_length])
        return np.array(X), np.array(y)

    def forecast_future(self, net, last_seq, n_steps):
        seq = last_seq.copy()
        preds = []
        for _ in range(n_steps):
            x = torch.tensor(seq[np.newaxis, :, np.newaxis], dtype=torch.float32)
            with torch.no_grad():
                p = net.predict(x.numpy())[0]
            preds.append(p)
            seq = np.roll(seq, -1)
            seq[-1] = p
        return np.array(preds)

    def update_outputs(self, selection):
        data = self.forecast.get(selection)

        if data:
            vals = data['forecast'].tolist()
            for i, lbl in enumerate(self.outputs):
                lbl.config(text=f"Прогноз {i + 1}: {vals[i]:.2f}")
            plt.figure(figsize=(10, 5))
            plt.plot(data['ts'], label="Історичні дані")
            plt.plot(data['forecast'], label="Прогноз TCN", color='red')
            plt.legend()
            plt.title(f'TCN прогноз для "{selection}"')
            plt.show()
        else:
            for lbl in self.outputs:
                lbl.config(text="—")

    def go_back(self):
        AnalysisWindow2(self.master.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
        ]

        # Додати всі прогнози з self.outputs
        for i, lbl in enumerate(self.outputs):
            result_lines.append(f"{lbl.cget('text')}")

        # Додати шлях до файлу
        result_lines.append(f"Файл даних: {self.selected_file if hasattr(self, 'selected_file') else 'Невідомо'}")
        result_lines.append("")
        result_lines.append("")

        # Зберігаємо у файл
        filename = "результат_аналізу.txt"
        with open(filename, "a", encoding="utf-8") as f:
            f.write("\n".join(result_lines))

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)
class TCN_PNPP(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.master = master
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")
        self.seq_length = 8
        self.scaler = StandardScaler()
        self.forecast = {}

        options = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
            "Загальний об'єм водних скидів, куб. м": 10,
            "Сумарний показник радіоактивних викидів, %": 11,
            "Сумарний індекс скиду радіоактивних речовин, %": 12
        }

        self.setup_gui(options)
        self.load_and_train()
        self.update_outputs(self.selected_option.get())

    def setup_gui(self, options):
        self.title("TCN прогноз для ПАЕС")
        self.geometry("1000x600")
        self.configure(bg="#D9D9D9")
        self.options = options

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Прогноз для ПАЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)
        self.outputs = []

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        for i in range(self.seq_length):
            lbl = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                           highlightbackground="black", highlightthickness=1)
            col = 0.15 if i % 2 == 0 else 0.53
            row = 0.38 + (i // 2) * 0.08
            lbl.place(relx=col, rely=row, width=320, height=30)
            self.outputs.append(lbl)

    def load_and_train(self):
        df = pd.read_excel(self.selected_file, engine='openpyxl')
        df = df[df['Назва станції'].str.strip() == "ПАЕС"].copy()
        df['date'] = pd.PeriodIndex.from_fields(year=df['Рік'], quarter=df['Квартал'], freq='Q-DEC').to_timestamp()
        df.set_index('date', inplace=True)

        for col in self.options:
            try:
                ts = df[col]
                if ts.dtype == object or not pd.api.types.is_numeric_dtype(ts):
                    ts = ts.astype(str).str.replace(",", ".", regex=False)
                    ts = ts.str.replace(r"[^\d\.\-]+", "", regex=True)
                    ts = pd.to_numeric(ts, errors='coerce')

                ts.dropna(inplace=True)
                ts.astype(float).interpolate().dropna()

                data = self.scaler.fit_transform(ts.values.reshape(-1, 1)).flatten()
                X, y = self.create_sequences(data)
                X = X[:, :, np.newaxis].astype(np.float32)
                y = y.astype(np.float32).reshape(-1, 1)

                net = NeuralNetRegressor(TCN, module__seq_length=self.seq_length, max_epochs=100, lr=1e-3,
                                         batch_size=8, optimizer=torch.optim.Adam, iterator_train__shuffle=True,
                                         verbose=0)
                net.fit(X, y)

                preds = self.forecast_future(net, data[-self.seq_length:], 8)
                future = self.scaler.inverse_transform(preds.reshape(-1, 1)).flatten()
                idx = pd.date_range(start=ts.index[-1] + pd.offsets.QuarterEnd(), periods=8, freq='Q')
                future_ts = pd.Series(future, index=idx)

                self.forecast[col] = {'ts': ts, 'forecast': future_ts}
            except Exception as e:
                print(f"Проблема з колонкою '{col}': {e}")

    def create_sequences(self, data):
        X, y = [], []
        for i in range(len(data) - self.seq_length):
            X.append(data[i:i + self.seq_length])
            y.append(data[i + self.seq_length])
        return np.array(X), np.array(y)

    def forecast_future(self, net, last_seq, n_steps):
        seq = last_seq.copy()
        preds = []
        for _ in range(n_steps):
            x = torch.tensor(seq[np.newaxis, :, np.newaxis], dtype=torch.float32)
            with torch.no_grad():
                p = net.predict(x.numpy())[0]
            preds.append(p)
            seq = np.roll(seq, -1)
            seq[-1] = p
        return np.array(preds)

    def update_outputs(self, selection):
        data = self.forecast.get(selection)

        if data:
            vals = data['forecast'].tolist()
            for i, lbl in enumerate(self.outputs):
                lbl.config(text=f"Прогноз {i + 1}: {vals[i]:.2f}")
            plt.figure(figsize=(10, 5))
            plt.plot(data['ts'], label="Історичні дані")
            plt.plot(data['forecast'], label="Прогноз TCN", color='red')
            plt.legend()
            plt.title(f'TCN прогноз для "{selection}"')
            plt.show()
        else:
            for lbl in self.outputs:
                lbl.config(text="—")

    def go_back(self):
        AnalysisWindow2(self.master.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
        ]

        # Додати всі прогнози з self.outputs
        for i, lbl in enumerate(self.outputs):
            result_lines.append(f"{lbl.cget('text')}")

        # Додати шлях до файлу
        result_lines.append(f"Файл даних: {self.selected_file if hasattr(self, 'selected_file') else 'Невідомо'}")
        result_lines.append("")
        result_lines.append("")

        # Зберігаємо у файл
        filename = "результат_аналізу.txt"
        with open(filename, "a", encoding="utf-8") as f:
            f.write("\n".join(result_lines))

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)
class TCN_KhNPP(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.master = master
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")
        self.seq_length = 8
        self.scaler = StandardScaler()
        self.forecast = {}

        options = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
            "Загальний об'єм водних скидів, куб. м": 10,
            "Сумарний показник радіоактивних викидів, %": 11,
            "Сумарний індекс скиду радіоактивних речовин, %": 12
        }

        self.setup_gui(options)
        self.load_and_train()
        self.update_outputs(self.selected_option.get())

    def setup_gui(self, options):
        self.title("TCN прогноз для ХАЕС")
        self.geometry("1000x600")
        self.configure(bg="#D9D9D9")
        self.options = options

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Прогноз для ХАЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)
        self.outputs = []

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        for i in range(self.seq_length):
            lbl = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                           highlightbackground="black", highlightthickness=1)
            col = 0.15 if i % 2 == 0 else 0.53
            row = 0.38 + (i // 2) * 0.08
            lbl.place(relx=col, rely=row, width=320, height=30)
            self.outputs.append(lbl)

    def load_and_train(self):
        df = pd.read_excel(self.selected_file, engine='openpyxl')
        df = df[df['Назва станції'].str.strip() == "ХАЕС"].copy()
        df['date'] = pd.PeriodIndex.from_fields(year=df['Рік'], quarter=df['Квартал'], freq='Q-DEC').to_timestamp()
        df.set_index('date', inplace=True)

        for col in self.options:
            try:
                ts = df[col]
                if ts.dtype == object or not pd.api.types.is_numeric_dtype(ts):
                    ts = ts.astype(str).str.replace(",", ".", regex=False)
                    ts = ts.str.replace(r"[^\d\.\-]+", "", regex=True)
                    ts = pd.to_numeric(ts, errors='coerce')

                ts.dropna(inplace=True)
                ts.astype(float).interpolate().dropna()

                data = self.scaler.fit_transform(ts.values.reshape(-1, 1)).flatten()
                X, y = self.create_sequences(data)
                X = X[:, :, np.newaxis].astype(np.float32)
                y = y.astype(np.float32).reshape(-1, 1)

                net = NeuralNetRegressor(TCN, module__seq_length=self.seq_length, max_epochs=100, lr=1e-3,
                                         batch_size=8, optimizer=torch.optim.Adam, iterator_train__shuffle=True,
                                         verbose=0)
                net.fit(X, y)

                preds = self.forecast_future(net, data[-self.seq_length:], 8)
                future = self.scaler.inverse_transform(preds.reshape(-1, 1)).flatten()
                idx = pd.date_range(start=ts.index[-1] + pd.offsets.QuarterEnd(), periods=8, freq='Q')
                future_ts = pd.Series(future, index=idx)

                self.forecast[col] = {'ts': ts, 'forecast': future_ts}
            except Exception as e:
                print(f"Проблема з колонкою '{col}': {e}")

    def create_sequences(self, data):
        X, y = [], []
        for i in range(len(data) - self.seq_length):
            X.append(data[i:i + self.seq_length])
            y.append(data[i + self.seq_length])
        return np.array(X), np.array(y)

    def forecast_future(self, net, last_seq, n_steps):
        seq = last_seq.copy()
        preds = []
        for _ in range(n_steps):
            x = torch.tensor(seq[np.newaxis, :, np.newaxis], dtype=torch.float32)
            with torch.no_grad():
                p = net.predict(x.numpy())[0]
            preds.append(p)
            seq = np.roll(seq, -1)
            seq[-1] = p
        return np.array(preds)

    def update_outputs(self, selection):
        data = self.forecast.get(selection)

        if data:
            vals = data['forecast'].tolist()
            for i, lbl in enumerate(self.outputs):
                lbl.config(text=f"Прогноз {i + 1}: {vals[i]:.2f}")
            plt.figure(figsize=(10, 5))
            plt.plot(data['ts'], label="Історичні дані")
            plt.plot(data['forecast'], label="Прогноз TCN", color='red')
            plt.legend()
            plt.title(f'TCN прогноз для "{selection}"')
            plt.show()
        else:
            for lbl in self.outputs:
                lbl.config(text="—")

    def go_back(self):
        AnalysisWindow2(self.master.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
        ]

        # Додати всі прогнози з self.outputs
        for i, lbl in enumerate(self.outputs):
            result_lines.append(f"{lbl.cget('text')}")

        # Додати шлях до файлу
        result_lines.append(f"Файл даних: {self.selected_file if hasattr(self, 'selected_file') else 'Невідомо'}")
        result_lines.append("")
        result_lines.append("")

        # Зберігаємо у файл
        filename = "результат_аналізу.txt"
        with open(filename, "a", encoding="utf-8") as f:
            f.write("\n".join(result_lines))

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)
class TCN_RNPP(tk.Toplevel):
    def __init__(self, master, selected_file):
        super().__init__(master)
        self.selected_file = selected_file
        self.master = master
        self.selected_option = tk.StringVar()
        self.selected_option.set("Радіоактивність викидів інертних радіоактивних газів, ГБк/доба")
        self.seq_length = 8
        self.scaler = StandardScaler()
        self.forecast = {}

        options = {
            "Радіоактивність викидів інертних радіоактивних газів, ГБк/доба": 0,
            "Відсоток від допустимого рівня викидів інертних радіоактивних газів, %": 1,
            "Радіоактивність викидів радіонуклідів йоду, кБк/доба": 2,
            "Відсоток від допустимого рівня викидів радіонуклідів йоду, %": 3,
            "Радіоактивність викидів довгоіснуючих радіонуклідів, кБк/доба": 4,
            "Відсоток від допустимого рівня викидів довгоіснуючих радіонуклідів, %": 5,
            "Радіоактивність газо-аерозольних викидів радіонукліду Cs-137, кБк/місяць": 6,
            "Радіоактивність газо-аерозольних викидів радіонукліду Co-60, кБк/місяць": 7,
            "Радіоактивність рідкого скиду радіонукліду Cs-137, кБк/місяць": 8,
            "Радіоактивність рідкого скиду радіонукліду Co-60, кБк/місяць": 9,
            "Загальний об'єм водних скидів, куб. м": 10,
            "Сумарний показник радіоактивних викидів, %": 11,
            "Сумарний індекс скиду радіоактивних речовин, %": 12
        }

        self.setup_gui(options)
        self.load_and_train()
        self.update_outputs(self.selected_option.get())

    def setup_gui(self, options):
        self.title("TCN прогноз для РАЕС")
        self.geometry("1000x600")
        self.configure(bg="#D9D9D9")
        self.options = options

        self.label7 = tk.Label(self, text="")
        self.label7.configure(bg="#FFFFFF", fg="#FFFFFF", anchor="center")
        self.label7.place(relx=0.1, rely=0.16, width=800, height=440)

        self.label = tk.Label(self, text="")
        self.label.configure(bg="#284B63", fg="#FFFFFF", anchor="center")
        self.label.place(relx=0.0, rely=0.0, width=1000, height=50)

        self.h_button = tk.Button(self, text="⌂", font=("Bahnschrift", 14), command=self.go_home)
        self.h_button.place(relx=0.020, rely=0.017, width=31, height=31)

        self.h1_button = tk.Button(self, text="←",
                                   font=("Bahnschrift", 14), command=self.go_back)
        self.h1_button.place(relx=0.950, rely=0.017, width=31, height=31)

        self.i_button = tk.Button(self, text="Інструкції",
                                  font=("Bahnschrift", 12), command=self.show_instructions)
        self.i_button.place(relx=0.060, rely=0.017, width=158, height=31)

        self.cur_button = tk.Label(self, text="Прогноз для РАЕС", font=("Bahnschrift", 12))
        self.cur_button.configure(bg="#C9D2D8")
        self.cur_button.place(relx=0.228, rely=0.017, width=241, height=31)

        self.option_menu = tk.OptionMenu(self, self.selected_option, *options, command=self.update_outputs)
        self.option_menu.config(font=("Bahnschrift", 13))
        self.option_menu.place(relx=0.15, rely=0.22, width=700, height=60)
        self.outputs = []

        self.save_button = tk.Button(self, text="Зберегти результат", font=("Bahnschrift", 16), bg="#3C6E71",
                                     fg="#FFFFFF", command=self.save_results)
        self.save_button.place(relx=0.15, rely=0.7, width=700, height=70)

        for i in range(self.seq_length):
            lbl = tk.Label(self, text="", font=("Bahnschrift", 15), bg="#FFFFFF", anchor="center",
                           highlightbackground="black", highlightthickness=1)
            col = 0.15 if i % 2 == 0 else 0.53
            row = 0.38 + (i // 2) * 0.08
            lbl.place(relx=col, rely=row, width=320, height=30)
            self.outputs.append(lbl)

    def load_and_train(self):
        df = pd.read_excel(self.selected_file, engine='openpyxl')
        df = df[df['Назва станції'].str.strip() == "РАЕС"].copy()
        df['date'] = pd.PeriodIndex.from_fields(year=df['Рік'], quarter=df['Квартал'], freq='Q-DEC').to_timestamp()
        df.set_index('date', inplace=True)

        for col in self.options:
            try:
                ts = df[col]
                if ts.dtype == object or not pd.api.types.is_numeric_dtype(ts):
                    ts = ts.astype(str).str.replace(",", ".", regex=False)
                    ts = ts.str.replace(r"[^\d\.\-]+", "", regex=True)
                    ts = pd.to_numeric(ts, errors='coerce')

                ts.dropna(inplace=True)
                ts.astype(float).interpolate().dropna()

                data = self.scaler.fit_transform(ts.values.reshape(-1, 1)).flatten()
                X, y = self.create_sequences(data)
                X = X[:, :, np.newaxis].astype(np.float32)
                y = y.astype(np.float32).reshape(-1, 1)

                net = NeuralNetRegressor(TCN, module__seq_length=self.seq_length, max_epochs=100, lr=1e-3,
                                         batch_size=8, optimizer=torch.optim.Adam, iterator_train__shuffle=True,
                                         verbose=0)
                net.fit(X, y)

                preds = self.forecast_future(net, data[-self.seq_length:], 8)
                future = self.scaler.inverse_transform(preds.reshape(-1, 1)).flatten()
                idx = pd.date_range(start=ts.index[-1] + pd.offsets.QuarterEnd(), periods=8, freq='Q')
                future_ts = pd.Series(future, index=idx)

                self.forecast[col] = {'ts': ts, 'forecast': future_ts}
            except Exception as e:
                print(f"Проблема з колонкою '{col}': {e}")

    def create_sequences(self, data):
        X, y = [], []
        for i in range(len(data) - self.seq_length):
            X.append(data[i:i + self.seq_length])
            y.append(data[i + self.seq_length])
        return np.array(X), np.array(y)

    def forecast_future(self, net, last_seq, n_steps):
        seq = last_seq.copy()
        preds = []
        for _ in range(n_steps):
            x = torch.tensor(seq[np.newaxis, :, np.newaxis], dtype=torch.float32)
            with torch.no_grad():
                p = net.predict(x.numpy())[0]
            preds.append(p)
            seq = np.roll(seq, -1)
            seq[-1] = p
        return np.array(preds)

    def update_outputs(self, selection):
        data = self.forecast.get(selection)

        if data:
            vals = data['forecast'].tolist()
            for i, lbl in enumerate(self.outputs):
                lbl.config(text=f"Прогноз {i + 1}: {vals[i]:.2f}")
            plt.figure(figsize=(10, 5))
            plt.plot(data['ts'], label="Історичні дані")
            plt.plot(data['forecast'], label="Прогноз TCN", color='red')
            plt.legend()
            plt.title(f'TCN прогноз для "{selection}"')
            plt.show()
        else:
            for lbl in self.outputs:
                lbl.config(text="—")

    def go_back(self):
        AnalysisWindow2(self.master.master, self.selected_file)
        self.withdraw()

    def go_home(self):
        self.destroy()
        self.master.deiconify()

    def save_results(self):
        result_lines = [
            f"Вибрана опція: {self.selected_option.get()}",
        ]

        # Додати всі прогнози з self.outputs
        for i, lbl in enumerate(self.outputs):
            result_lines.append(f"{lbl.cget('text')}")

        # Додати шлях до файлу
        result_lines.append(f"Файл даних: {self.selected_file if hasattr(self, 'selected_file') else 'Невідомо'}")
        result_lines.append("")
        result_lines.append("")

        # Зберігаємо у файл
        filename = "результат_аналізу.txt"
        with open(filename, "a", encoding="utf-8") as f:
            f.write("\n".join(result_lines))

    def show_instructions(self):
        if hasattr(self, 'instruction_window') and self.instruction_window.winfo_exists():
            self.instruction_window.lift()
            return

        self.instruction_window = tk.Toplevel(self)
        self.instruction_window.title("Інструкції")
        self.instruction_window.geometry("800x600")
        self.instruction_window.resizable(False, False)
        self.instruction_window.transient(self)
        self.instruction_window.grab_set()

        label = tk.Label(self.instruction_window,
                         text='Інструкція:\n\n1. Ознайомтеся зі структурою файлу "шаблон.xlsx".\n'
                              '2. Введіть дані для опрацювання за зразком.\n'
                              '3. Переконайтеся, що маєте встановлену мову програмування Python та середовище програмування, яке її підтримує.\n'
                              '4. Запустіть програму та дотримуйтеся подальших інструкцій для окремих сторінок.\n\n'
                              'Домашня сторінка:\n'
                              '- На верхній панелі керування можна знайти кнопки "Додому", "Інструкції" та "Назад", а також переглянути назву поточного вікна.\n'
                              '- Маємо додаткову кнопку "Інструкції", з якими варто ознайомитися перед початком роботи.\n'
                              '- Є кнопка "Про автора", де можна знайти інформацію про автора програми.\n'
                              '- Натисніть кнопку "Для початку роботи оберіть фвйл для аналізу", аби обрати файл.\n\n'
                              'Оберіть дію:\n'
                              '- Оберіть аналіз чи передбачення, які Вас цікавлять, і натисніть на відповідну кнопку. \nВам відкриється нове вікно з варіантами АЕС для аналізу.\n\n'
                              'Оберіть АЕС для аналізу:\n'
                              '- Оберіть АЕС, по значеннях показників якої Ви хочете отримати аналіз чи передбачення.\n\n'
                              'Аналіз:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках "Мода", "Медіана" та "Середнє арифметичне" будуть автоматично обчислюватися\n ці значення для обраного показника.\n'
                              '- Якщо показник будь-якої з цих величин перевищує норму, то біля цього значення буде висвітлюватися\n знак помилки. При наведені в ту область можно побачити підказку до даної помилки.\n'
                              '- Кнопка "Зберегти результати" дозволяє виводити дані результати у файл розширення txt, куди будуть занотовуватися всі збережені дані.\n\n'
                              'Передбачення:\n'
                              '- З випадаючого вікна оберіть той показник, показники якого Ви хочете проаналізувати.\n'
                              '- У віконечках показані дані, розраховані на вісім сезонів уперед.\n'
                              '- В окремому вікні можна побачити графік, побудлваний до обраної величини.',
                         justify="left")
        label.pack(padx=0, pady=20)
        close_button = tk.Button(self.instruction_window, text="Закрити", command=self.instruction_window.destroy)
        close_button.pack(pady=10)


if __name__ == "__main__":
    app = FileSelectorApp()
    app.mainloop()
