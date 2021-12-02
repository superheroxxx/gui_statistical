import tkinter as tk
import tkinter.messagebox as mb
import tkinter.filedialog as fd
import os
import subprocess
import psutil
import datetime
import pandas as pd
import time

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Statistical Data")
        self.geometry("350x200")

        self.tit = tk.Label(self, text = "Для сбора статистики необходимо выбрать интервал\nсбора статистики и выбрать файл для запуска процесса")
        self.tit.grid(row=0, column=0, columnspan=2)

        self.i_l = tk.Label(self, text = "Интервал")
        self.i_l.grid(row = 1, column = 0)
        vcmd = (self.register(self.validate), '%P')
        self.i_e = tk.Entry(self,validate='key', validatecommand=vcmd)
        self.i_e.grid(row = 1, column = 1)
        self.name_f = tk.Entry(self)
        self.name_f.grid(row = 2, column = 1)
        self.name_f.xview_scroll

        self.btn_file = tk.Button(self, text="Выбрать файл", command = self.choose_file)
        self.btn_file.grid(row = 2, column = 0)
        self.tit = tk.Label(self, text = "Для запуска нажмите \"Start\" ")
        self.tit.grid(row=3, column=0, columnspan=2)

        button = tk.Button(self, text = "Start", command = self.start)
        button.grid(row = 4, column = 0, columnspan=2)

    def validate(self, new_value):
        return new_value == "" or new_value.isnumeric()
    
    def choose_file(self):
        path_file = fd.askopenfilename()
        if path_file:
            self.name_f.delete(0,'end')
            self.name_f.insert(0,path_file)

    def start(self):
        if self.name_f.index("end") == 0 or self.i_e.index("end") == 0:
            self.show_warning()
            return
        str_e = self.name_f.get()
        try:
            os.startfile(str_e)
        except FileNotFoundError:
            self.show_error()
        mypid = self.find_process_pid(str_e[str_e.rfind("/")+1:])
        
        dt_obj = datetime.datetime.now()
        dt_string = dt_obj.strftime("%d.%m.%Y.%H.%M.%S.%f")
        n_file = 'report_'+dt_string+'.xlsx'

        p = psutil.Process(mypid)
        pi_pid=[]
        time_w=[]
        cpu_percent=[]
        rss=[]
        vms=[]
        num_handles=[]
        while psutil.pid_exists(mypid):
            pi_pid.append(mypid)
            dt_obj1 = datetime.datetime.now()
            dt_string1 = dt_obj1.strftime("%H:%M:%S.%f")
            time_w.append(dt_string1)
            with p.oneshot():
                cpu_percent.append(p.cpu_percent(interval=1))
                rss.append(p.memory_info().rss)
                vms.append(p.memory_info().vms)
                num_handles.append(p.num_handles())
            time.sleep(int(self.i_e.get())-1)
            
        df = pd.DataFrame({'pi_pid': pi_pid,
                           'time_w': time_w,
                           'cpu_percent': cpu_percent,
                           'rss': rss,
                           'vms': vms,
                           'num_handles': num_handles})
        df.to_excel(n_file)

    def find_process_pid(self,process_name):
        for process in psutil.process_iter():
            if process.name() == process_name:
                return process.pid

    def show_error(self):
        msg = "Файл не найден"
        mb.showerror("Ошибка", msg)

    def show_warning(self):
        msg = "Введите интервал и выберите файл"
        mb.showwarning("Предупреждение", msg)
        
    
if __name__ == "__main__":
    app = App()
    app.mainloop()
