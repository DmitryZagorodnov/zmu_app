import os
import sys
import io
import tkintermapview
import requests
import matplotlib
from tkinter import *
from tkinter import filedialog as fd
from PIL import Image, ImageTk
from gpx_parser import GpxParser
from docxtpl import DocxTemplate
from docxtpl import InlineImage
from docx.shared import Mm
from tkinter import messagebox as mb
from tkinter import ttk
from tkcalendar import DateEntry
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg,
)

matplotlib.use('TkAgg')


class MyApp:

    def __init__(self):
        self.root = Tk()
        self.entr_list = []
        self.fields_values = {}
        self.aliases = {}

        self.context = {}
        self.dates = {}

        self.tabs_master = None
        self.tab_create_report = None
        self.tab_day_profiles = None
        self.tab_area_profiles = None
        self.setup_window = None

        self.reports = []
        self.day_profiles = []
        self.area_profiles = []

        self.reports_count = 0
        self.day_profiles_count = 0
        self.area_profiles_count = 0

        self.areas = {}
        self.days = {}

        self.button1 = None
        self.button2 = None
        self.button3 = None

        self.areas_cbs = []
        self.days_cbs = []

        self.template_way = './template.docx'
        self.api_url = 'https://static-maps.yandex.ru/1.x/'

    def draw_root(self):
        self.tabs_master = ttk.Notebook(self.root)
        self.tabs_master.grid(row=0, column=0)
        main_menu = Menu(self.root)
        self.root.config(menu=main_menu)

        options = Menu(main_menu, tearoff=0)
        options.add_command(label="Настройки", command=self.call_setup)
        options.add_separator()
        options.add_command(label="Выход", command=sys.exit)

        main_menu.add_cascade(label="Опции", menu=options)

    def draw_tab1(self):
        self.tab_create_report = Frame(self.tabs_master)
        self.tabs_master.add(self.tab_create_report, text="Создание ведомости")

        # Init columns' headers
        borderwidth = 1
        relief = 'solid'
        width = 20
        row = 0

        Label(master=self.tab_create_report, text='Имя ведомости', width=width, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=0)
        Label(master=self.tab_create_report, text='Номер маршрута', width=width - 5, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=1)
        Label(master=self.tab_create_report, text='Профиль территории', width=width, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=2)
        Label(master=self.tab_create_report, text='Профиль дня', width=width, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=3)
        Label(master=self.tab_create_report, text='Учётчик', width=width, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=4)
        Label(master=self.tab_create_report, text='Файл трека', width=width, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=5)
        Label(master=self.tab_create_report, text='Файл меток', width=width, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=6)

        self.reports_count = 1

        self.button1 = Button(self.tab_create_report, text='+', command=self.create_report, width=19)
        self.button1.grid(row=self.reports_count, column=0)

    def create_report(self):
        rn = Entry(self.tab_create_report, fg="black", bg="white", width=23, justify=RIGHT)
        rn.insert(0, f"Ведомость №{self.reports_count}")
        tn = ttk.Spinbox(self.tab_create_report, from_=0, to=1000, width=15, justify=RIGHT)
        ap = ttk.Combobox(self.tab_create_report, values=[], postcommand=self.get_areas, width=20)
        self.areas_cbs.append(ap)
        dp = ttk.Combobox(self.tab_create_report, values=[], postcommand=self.get_days, width=20)
        self.days_cbs.append(dp)
        uch = Entry(self.tab_create_report, fg="black", bg="white", width=23, justify=RIGHT)
        ft = ttk.Combobox(self.tab_create_report, values=[], width=20, postcommand=lambda: self.get_tracksfile(ft))
        fw = ttk.Combobox(self.tab_create_report, values=[], width=20, postcommand=lambda: self.get_tracksfile(fw))

        new_report = [rn, tn, ap, dp, uch, ft, fw]
        self.reports.append(new_report)
        self.reports_count = self.reports_count + 1
        self.draw_report(new_report)

    def draw_report(self, report):
        self.button1.destroy()
        for ind, field in enumerate(report):
            field.grid(row=self.reports_count, column=ind)
        Button(self.tab_create_report, text='Создать ведомость', command=lambda: self.fill_report(report),
               width=15).grid(row=self.reports_count, column=len(report))
        Button(self.tab_create_report, text='Схема трека', command=lambda: self.show_track(report),
               width=10).grid(row=self.reports_count, column=len(report) + 1)
        Button(self.tab_create_report, text='Карта', command=lambda: self.show_map(report),
               width=10).grid(row=self.reports_count, column=len(report) + 2)
        self.button1 = Button(self.tab_create_report, text='+', command=self.create_report, width=19)
        self.button1.grid(row=self.reports_count + 1, column=0)

    def fill_report(self, report):
        output_file = report[0].get()
        route_number = report[1].get()
        area_profile = report[2].get()
        day_profile = report[3].get()
        accountant = report[4].get()
        track_file = report[5].get()
        waypoints_file = report[6].get()

        self.context['RN'] = route_number
        if area_profile:
            subject, district, area = self.areas[area_profile]
            self.context['SRF'] = subject
            self.context['MO'] = district
            self.context['IT'] = area

        if day_profile:
            self.prepare_day_context(day_profile)

        self.context['UCH'] = accountant

        pars = GpxParser()
        if track_file:
            try:
                pars.parse_track(track_file)
            except ValueError:
                mb.showerror(title='Error', message="Your file doesn't contains any track")
                return
        if waypoints_file:
            try:
                pars.parse_waypoints(waypoints_file)
                pars.parse()
            except ValueError:
                mb.showerror(title='Error', message="Your file doesn't contains any waypoints")
                return

        doc = DocxTemplate(self.template_way)
        image = self.get_track(track_file=track_file, waypoints_file=waypoints_file)
        if not isinstance(image, int):
            image.savefig('temp.png')
            self.context['image'] = InlineImage(doc, 'temp.png', width=Mm(100))

        cont = pars.prepare_context()
        # self.parse_dates()
        for key, value in cont.items():
            self.context[key] = value
        doc.render(context=self.context)
        doc.save(f"{output_file}.docx")
        if image in self.context:
            os.remove('./temp.png')
        self.context = {}
        mb.showinfo('Success!', 'Your file successfully created!')

    def get_track(self, track_file=None, waypoints_file=None):
        if not track_file and not waypoints_file:
            return 0
        else:
            pars = GpxParser()
            f = Figure(figsize=(7, 7), dpi=100)
            f_plot = f.add_subplot(111)
            if track_file:
                try:
                    pars.parse_track(track_file)
                    f_plot.plot(pars.longs, pars.lats, color='red')
                except ValueError:
                    return -1

            if waypoints_file:
                try:
                    pars.parse_waypoints(waypoints_file)
                except ValueError:
                    return -1
                for ind, pt in enumerate(pars.waypoints):
                    f_plot.plot(pt[1], pt[2], color='gray', marker='o')

                    if ind % 2 == 0:
                        f_plot.text(pt[1], pt[2], pt[0], ha='left', verticalalignment='top', backgroundcolor='white',
                                    bbox={'visible': True, 'facecolor': 'white'})
                    elif ind % 2 == 1:
                        f_plot.text(pt[1], pt[2], pt[0], ha='right', verticalalignment='top', backgroundcolor='white',
                                    bbox={'visible': True, 'facecolor': 'white'})
                    elif ind % 2 == 2:
                        f_plot.text(pt[1], pt[2], pt[0], ha='right', verticalalignment='bottom',
                                    backgroundcolor='white',
                                    bbox={'visible': True, 'facecolor': 'white'})
                    elif ind % 2 == 3:
                        f_plot.text(pt[1], pt[2], pt[0], ha='left', verticalalignment='bottom', backgroundcolor='white',
                                    bbox={'visible': True, 'facecolor': 'white'})

            f_plot.text(0, 1, "С\u2191", transform=f_plot.transAxes, fontsize=20)
            return f

    def show_track(self, report):
        track_file = report[5].get()
        waypoints_file = report[6].get()
        img = self.get_track(track_file=track_file, waypoints_file=waypoints_file)
        if img == 0:
            mb.showerror(title='Missed files', message='You should fill any filename entry')
            return
        elif img == -1:
            mb.showerror(title='Wrong files', message='Some of your files doesn\'t contains required content')
            return
        else:
            daughter = Tk()
            daughter.title('Track')
            b = Button(daughter, text="Save", command=lambda: self.save_pic(img), width=20)
            b.pack(side=BOTTOM)
            canvs = FigureCanvasTkAgg(img, daughter)
            canvs.get_tk_widget().pack(side=TOP, fill=BOTH, expand=1)
            daughter.mainloop()

    def save_pic(self, pic):
        file_to_save = fd.asksaveasfilename(initialdir=os.getcwd(),
                                            title="Select the output file name",
                                            filetypes=[('Png files', '.png')])
        pic.savefig(f'{file_to_save}.png')
        mb.showinfo('Success!', 'Your image successfully saved!')

    def show_map(self, report):
        track_file = report[5].get()
        waypoints_file = report[6].get()
        if not track_file and not waypoints_file:
            mb.showerror(title='Ошибка', message='Вы не указали ни одного gpx файла!')
            return

        params = {
            'l': 'sat'
        }

        p = GpxParser()
        if track_file:
            try:
                p.parse_track(track_file)
                pl = ''
                for i in range(len(p.lats) // 5):
                    pl += f'{round(p.longs[i * 5], 6)},{round(p.lats[i * 5], 6)},'
                params['pl'] = pl[:-1]
            except ValueError:
                mb.showerror(title='Ошибка', message=f'Указанный файл трека {track_file} '
                                                     f'не содержит в себе координат маршрута')
                return

        if waypoints_file:
            try:
                p.parse_waypoints(waypoints_file)
                pt = ""
                for i in range(len(p.waypoints)):
                    wp = p.waypoints[i]
                    if i != len(p.waypoints) - 1:
                        pt += f"{wp[1]},{wp[2]},pmors{i + 1}~"
                params['pt'] = pt[:-1]
            except ValueError:
                mb.showerror(title='Ошибка', message=f'Указанный файл меток {waypoints_file} не содержит в себе меток')
                return

        try:
            r = requests.get(self.api_url, params)
            print(r.status_code)
            image = ImageTk.PhotoImage(data=r.content)

            daughter = Toplevel()

            label = Label(daughter, image=image)
            label.image = image
            label.pack(side=TOP)

            Button(daughter, text="Сохранить изображение", command=lambda: self.save_map(r.content)).pack(side=BOTTOM)

        except requests.exceptions.ConnectionError:
            mb.showerror(title="Ошибка", message="Нет интернет-соединения")
            return


    def save_map(self, map):
        filename = fd.asksaveasfilename(initialdir=os.getcwd(),
                                        title="Select the output file name",
                                        filetypes=[('Png files', '.png')])
        with open(f"{filename}.png", 'wb') as fn:
            fn.write(map)
        mb.showinfo(title='Успешно!', message="Ваш файл успешно сохранен!")

    def get_areas(self):
        for area_cb in self.areas_cbs:
            area_cb['values'] = list(self.areas.keys())

    def get_days(self):
        for day_cb in self.days_cbs:
            day_cb['values'] = list(self.days.keys())

    def get_tracksfile(self, entry):
        entry.delete(0, 'end')
        filetype = [('gpx files', '*.gpx')]
        filename = fd.askopenfilename(filetype=filetype)
        entry.insert(0, filename)

    def draw_tab2(self):
        self.tab_area_profiles = Frame(self.tabs_master)
        self.tabs_master.add(self.tab_area_profiles, text="Профили территории")

        borderwidth = 1
        relief = 'solid'
        width = 30
        row = 0

        Label(master=self.tab_area_profiles, text='Название профиля', width=width, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=0)
        Label(master=self.tab_area_profiles, text='Субъект Российской Федерации', width=width, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=1)
        Label(master=self.tab_area_profiles, text='Муниципальное образование (район)', width=width,
              borderwidth=borderwidth, relief=relief).grid(row=row, column=2)
        Label(master=self.tab_area_profiles, text='Исследуемая территория', width=width, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=3)
        self.area_profiles_count = 1

        self.button2 = Button(self.tab_area_profiles, text='+', command=self.create_new_area, width=width)
        self.button2.grid(row=self.area_profiles_count, column=0)

    def create_new_area(self):
        new_area_window = Tk()
        new_area_window.title('Создание нового профиля территории')
        Label(new_area_window, text="Название профиля:", width=35).grid(row=0, column=0)
        pn = Entry(new_area_window, fg="black", bg="white", width=35, justify=RIGHT)
        pn.insert(0, f"Профиль №{self.area_profiles_count}")
        pn.grid(row=0, column=1)
        Label(new_area_window, text="Субъект Российской федерации:", width=35).grid(row=1, column=0)
        sn = Entry(new_area_window, fg='black', bg='white', width=35, justify=RIGHT)
        sn.grid(row=1, column=1)
        Label(new_area_window, text="Муниципальное образование (район):", width=35).grid(row=2, column=0)
        dn = Entry(new_area_window, fg='black', bg='white', width=35, justify=RIGHT)
        dn.grid(row=2, column=1)
        Label(new_area_window, text="Исследуемая территория:", width=35).grid(row=3, column=0)
        an = Entry(new_area_window, fg='black', bg='white', width=35, justify=RIGHT)
        an.grid(row=3, column=1)

        new_area = [pn, sn, dn, an]

        Button(new_area_window, text="Сохранить",
               command=lambda: self.save_area(new_area, new_area_window)).grid(row=4, column=0)

    def save_area(self, area, window, cur_row=None):
        name = area[0].get()
        subject = area[1].get()
        district = area[2].get()
        r_area = area[3].get()
        if cur_row:
            old_name = self.tab_area_profiles.grid_slaves(row=cur_row, column=0)[0].cget('text')
            del self.areas[old_name]
        else:
            self.area_profiles_count = self.area_profiles_count + 1
        if any((name, subject, district, r_area)):
            self.areas[name] = [subject, district, r_area]
        self.draw_new_area([name, subject, district, r_area], cur_row)
        window.destroy()

    def draw_new_area(self, area, cur_row=None):
        if cur_row:
            old_labels = self.tab_area_profiles.grid_slaves(row=cur_row)
            for label in old_labels:
                label.grid_forget()
            for ind, field in enumerate(area):
                Label(master=self.tab_area_profiles, width=30, text=field,
                      justify=RIGHT).grid(row=cur_row, column=ind)
            Button(self.tab_area_profiles, text="Редактировать профиль",
                   command=lambda: self.edit_area(area, cur_row),
                   width=30).grid(row=cur_row, column=len(area))
        else:
            self.button2.destroy()
            for ind, field in enumerate(area):
                Label(master=self.tab_area_profiles, width=30, text=field,
                      justify=RIGHT).grid(row=self.area_profiles_count, column=ind)
            self.button2 = Button(self.tab_area_profiles, text='+', command=self.create_new_area, width=30)
            self.button2.grid(row=self.area_profiles_count + 1, column=0)
            cur_row = self.area_profiles_count
            Button(self.tab_area_profiles, text="Редактировать профиль",
                   command=lambda: self.edit_area(area, cur_row),
                   width=30).grid(row=self.area_profiles_count, column=len(area))

    def edit_area(self, area, cur_row):
        edit_area_window = Tk()
        edit_area_window.title('Редактирование профиля территории')
        Label(edit_area_window, text="Название профиля:", width=35).grid(row=0, column=0)
        pn = Entry(edit_area_window, fg="black", bg="white", width=35, justify=RIGHT)
        pn.insert(0, area[0])
        pn.grid(row=0, column=1)
        Label(edit_area_window, text="Субъект Российской федерации:", width=35).grid(row=1, column=0)
        sn = Entry(edit_area_window, fg='black', bg='white', width=35, justify=RIGHT)
        sn.insert(0, area[1])
        sn.grid(row=1, column=1)
        Label(edit_area_window, text="Муниципальное образование (район):", width=35).grid(row=2, column=0)
        dn = Entry(edit_area_window, fg='black', bg='white', width=35, justify=RIGHT)
        dn.insert(0, area[2])
        dn.grid(row=2, column=1)
        Label(edit_area_window, text="Исследуемая территория:", width=35).grid(row=3, column=0)
        an = Entry(edit_area_window, fg='black', bg='white', width=35, justify=RIGHT)
        an.insert(0, area[3])
        an.grid(row=3, column=1)

        edited_area = [pn, sn, dn, an]

        Button(edit_area_window, text="Сохранить",
               command=lambda: self.save_area(edited_area, edit_area_window, cur_row=cur_row)).grid(row=4, column=0)

    def draw_tab3(self):
        self.tab_day_profiles = Frame(self.tabs_master)
        self.tabs_master.add(self.tab_day_profiles, text="Профили дней учета")

        borderwidth = 1
        relief = 'solid'
        width = 30
        row = 0

        Label(master=self.tab_day_profiles, text='Название профиля', width=20, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=0)
        Label(master=self.tab_day_profiles, text='Окончание пороши', width=20, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=1)
        Label(master=self.tab_day_profiles, text='Исп-е транспорта', width=20, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=2)
        Label(master=self.tab_day_profiles, text='Исп-е навигатора', width=20, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=3)
        Label(master=self.tab_day_profiles, text='Затирка', width=20, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=4)
        Label(master=self.tab_day_profiles, text='Учет следов', width=20, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=5)
        Label(master=self.tab_day_profiles, text='Снег', width=20, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=6)
        Label(master=self.tab_day_profiles, text='Погода', width=20, borderwidth=borderwidth,
              relief=relief).grid(row=row, column=7)

        self.day_profiles_count = 1

        self.button3 = Button(self.tab_day_profiles, text='+', command=self.create_new_day, width=20)
        self.button3.grid(row=self.day_profiles_count, column=0)

    def create_new_day(self, day=None, cur_row=None):
        new_day_window = Tk()
        if not cur_row:
            new_day_window.title("Создание нового профиля дня")
        else:
            new_day_window.title("Редактирование профиля дня")

        new_day_fields = []

        i = 0
        Label(master=new_day_window, text='Название профиля', width=30).grid(row=i, column=0)
        day_name = Entry(master=new_day_window, fg="black", bg="white", width=30, justify=RIGHT)
        if not cur_row:
            day_name.insert(0, f"Профиль №{self.day_profiles_count}")
        new_day_fields.append(day_name)
        day_name.grid(row=i, column=1)
        i = i + 1

        Label(master=new_day_window, text='Дата окончания последней пороши', width=30).grid(row=i, column=0)
        dp_date = DateEntry(master=new_day_window, fg="black", bg="white", width=27, justify=RIGHT,
                            date_pattern='dd-MM-yyyy')
        new_day_fields.append(dp_date)
        dp_date.delete(0, 'end')
        dp_date.grid(row=i, column=1)

        Label(master=new_day_window, text="Время окончания:", width=20).grid(row=i, column=2)
        dp_h = Entry(master=new_day_window, fg="black", bg="white", width=20, justify=RIGHT)
        new_day_fields.append(dp_h)
        dp_h.grid(row=i, column=3)
        Label(master=new_day_window, text="час.").grid(row=i, column=4)
        i = i + 1

        Label(master=new_day_window, text="Использование транспортного средства").grid(row=i, column=0)
        its = ttk.Combobox(master=new_day_window, values=["Да", "Нет"], width=27, justify=RIGHT)
        new_day_fields.append(its)
        its.grid(row=i, column=1)
        Label(master=new_day_window, text="(да/нет)").grid(row=i, column=2)
        i = i + 1

        Label(master=new_day_window, text="Использование спутникового навигатора").grid(row=i, column=0)
        isn = ttk.Combobox(master=new_day_window, values=["Да", "Нет"], width=27, justify=RIGHT)
        new_day_fields.append(isn)
        isn.grid(row=i, column=1)
        Label(master=new_day_window, text="да/нет").grid(row=i, column=2)
        i = i + 1

        Label(master=new_day_window, text="Дата затирки").grid(row=i, column=0)
        dz_date = DateEntry(master=new_day_window, fg="black", bg="white", width=27, justify=RIGHT,
                            date_pattern='dd-MM-yyyy')
        new_day_fields.append(dz_date)
        dz_date.delete(0, 'end')
        dz_date.grid(row=i, column=1)

        Label(master=new_day_window, text="Время начала:", justify=RIGHT).grid(row=i, column=2)
        dz_hb = Entry(master=new_day_window, fg="black", bg="white", width=20, justify=RIGHT)
        new_day_fields.append(dz_hb)
        dz_hb.grid(row=i, column=3)
        Label(master=new_day_window, text="час.").grid(row=i, column=4)
        dz_mb = Entry(master=new_day_window, fg="black", bg="white", width=20, justify=RIGHT)
        new_day_fields.append(dz_mb)
        dz_mb.grid(row=i, column=5)
        Label(master=new_day_window, text="мин., окончание").grid(row=i, column=6)
        dz_he = Entry(master=new_day_window, fg="black", bg="white", width=20, justify=RIGHT)
        new_day_fields.append(dz_he)
        dz_he.grid(row=i, column=7)
        Label(master=new_day_window, text="час.").grid(row=i, column=8)
        dz_me = Entry(master=new_day_window, fg="black", bg="white", width=20, justify=RIGHT)
        new_day_fields.append(dz_me)
        dz_me.grid(row=i, column=9)
        Label(master=new_day_window, text="мин.").grid(row=i, column=10)
        i = i + 1

        Label(master=new_day_window, text="Дата учета следов").grid(row=i, column=0)
        dus_date = DateEntry(master=new_day_window, fg="black", bg="white", width=27, justify=RIGHT,
                             date_pattern='dd-MM-yyyy')
        dus_date.delete(0, 'end')
        new_day_fields.append(dus_date)
        dus_date.grid(row=i, column=1)

        Label(master=new_day_window, text="Время начала:").grid(row=i, column=2)
        dus_hb = Entry(master=new_day_window, fg="black", bg="white", width=20, justify=RIGHT)
        new_day_fields.append(dus_hb)
        dus_hb.grid(row=i, column=3)
        Label(master=new_day_window, text="час.").grid(row=i, column=4)
        dus_mb = Entry(master=new_day_window, fg="black", bg="white", width=20, justify=RIGHT)
        new_day_fields.append(dus_mb)
        dus_mb.grid(row=i, column=5)
        Label(master=new_day_window, text="мин., окончание").grid(row=i, column=6)
        dus_he = Entry(master=new_day_window, fg="black", bg="white", width=20, justify=RIGHT)
        new_day_fields.append(dus_he)
        dus_he.grid(row=i, column=7)
        Label(master=new_day_window, text="час.").grid(row=i, column=8)
        dus_me = Entry(master=new_day_window, fg="black", bg="white", width=20, justify=RIGHT)
        new_day_fields.append(dus_me)
        dus_me.grid(row=i, column=9)
        Label(master=new_day_window, text="мин.").grid(row=i, column=10)
        i = i + 1

        Label(master=new_day_window, text="Высота снега").grid(row=i, column=0)
        sh = Entry(master=new_day_window, fg="black", bg="white", width=30, justify=RIGHT)
        new_day_fields.append(sh)
        sh.grid(row=i, column=1)
        Label(master=new_day_window, text="см.").grid(row=i, column=2)
        i = i + 1

        Label(master=new_day_window, text="Характер снега (рыхлый, плотный и др.)").grid(row=i, column=0)
        sc = Entry(master=new_day_window, fg="black", bg="white", width=30, justify=RIGHT)
        new_day_fields.append(sc)
        sc.grid(row=i, column=1)
        i = i + 1

        Label(master=new_day_window, text="Погода в день учета следов: температура").grid(row=i, column=0)
        wt = Entry(master=new_day_window, fg="black", bg="white", width=30, justify=RIGHT)
        new_day_fields.append(wt)
        wt.grid(row=i, column=1)
        i = i + 1
        Label(master=new_day_window, text="Погода в день учета следов: осадки").grid(row=i, column=0)
        wd = Entry(master=new_day_window, fg="black", bg="white", width=30, justify=RIGHT)
        new_day_fields.append(wd)
        wd.grid(row=i, column=1)
        i = i + 1

        if cur_row:
            for ind, field in enumerate(new_day_fields):
                field.insert(0, day[ind])

        Button(master=new_day_window, text="Сохранить",
               command=lambda: self.save_day(new_day_fields, new_day_window, cur_row)).grid(row=i, column=0)

    def save_day(self, day, window, cur_row=None):
        if cur_row:
            old_name = self.tab_day_profiles.grid_slaves(row=cur_row, column=0)[0].cget('text')
            del self.days[old_name]
        else:
            self.day_profiles_count = self.day_profiles_count + 1

        values = []
        for field in day:
            values.append(field.get())
        if any(values):
            self.days[values[0]] = values[1:]
            print(self.days)
        self.draw_new_day(values, cur_row)
        window.destroy()

    def draw_new_day(self, day, cur_row=None):
        day_info = self.prepare_day_to_draw(day)
        if cur_row:
            old_labels = self.tab_day_profiles.grid_slaves(row=cur_row)
            for label in old_labels:
                label.grid_forget()
            for ind, field in enumerate(day_info):
                Label(master=self.tab_day_profiles, width=20, text=field,
                      justify=RIGHT).grid(row=cur_row, column=ind)
            Button(self.tab_day_profiles, text="Редактировать профиль",
                   command=lambda: self.edit_day(day, cur_row),
                   width=30).grid(row=cur_row, column=len(day))
        else:
            self.button3.destroy()
            for ind, field in enumerate(day_info):
                Label(master=self.tab_day_profiles, width=20, text=field,
                      justify=RIGHT).grid(row=self.day_profiles_count, column=ind)
            self.button3 = Button(self.tab_day_profiles, text='+', command=self.create_new_day, width=20)
            self.button3.grid(row=self.day_profiles_count + 1, column=0)
            cur_row = self.day_profiles_count
            Button(self.tab_day_profiles, text="Редактировать профиль",
                   command=lambda: self.edit_day(day, cur_row),
                   width=30).grid(row=self.day_profiles_count, column=len(day))

    def prepare_day_context(self, day_name):
        day = self.days[day_name]
        self.context["DP"] = day[0]
        self.context["DPH"] = day[1]
        self.context["ITS"] = day[2]
        self.context["ISN"] = day[3]

        zatirka_date = day[4].split('-')
        self.context["DZD"] = zatirka_date[0]
        self.context["DZM"] = zatirka_date[1]
        self.context["DZY"] = zatirka_date[2]
        self.context["DZHB"] = day[5]
        self.context["DZMB"] = day[6]
        self.context["DZHE"] = day[7]
        self.context["DZME"] = day[8]

        uchet_date = day[9].split('-')
        self.context["DUSD"] = uchet_date[0]
        self.context["DUSM"] = uchet_date[1]
        self.context["DUSY"] = uchet_date[2]
        self.context["DUSHB"] = day[10]
        self.context["DUSMB"] = day[11]
        self.context["DUSHE"] = day[12]
        self.context["DUSME"] = day[13]

        self.context["SH"] = day[14]
        self.context["SC"] = day[15]
        self.context["WT"] = day[16]
        self.context["WD"] = day[17]

    def prepare_day_to_draw(self, day):
        day_to_draw = []

        day_name = day[0]
        day_to_draw.append(day_name)

        porosha_date = day[1]
        porosha_time = day[2]
        if porosha_date and porosha_time:
            day_to_draw.append(f"{porosha_date}, {porosha_time} ч.")
        elif porosha_date:
            day_to_draw.append(f"{porosha_date}")
        else:
            day_to_draw.append(f"{porosha_time} ч.")

        vehicle_used = day[3]
        day_to_draw.append(vehicle_used)

        navigator_used = day[4]
        day_to_draw.append(navigator_used)

        zatirka_date = day[5]
        zatirka_beginning_time = f"{day[6]}:{day[7]}"
        zatirka_ending_time = f"{day[8]}:{day[9]}"
        temp = ""
        if zatirka_date:
            temp += f"{zatirka_date}"
        if day[6] or day[7]:
            temp += f", {zatirka_beginning_time}"
        if day[8] or day[9]:
            temp += f"-{zatirka_ending_time}"
        day_to_draw.append(temp)

        uchet_date = day[10]
        uchet_beginning_time = f"{day[11]}:{day[12]}"
        uchet_ending_time = f"{day[13]}:{day[14]}"
        temp = ""
        if zatirka_date:
            temp += f"{uchet_date}"
        if day[11] or day[12]:
            temp += f", {uchet_beginning_time}"
        if day[13] or day[14]:
            temp += f"-{uchet_ending_time}"
        day_to_draw.append(temp)

        snow_height = day[15]
        snow_character = day[16]
        temp = ""
        if snow_height:
            temp += f"{snow_height} см."
        if snow_character:
            temp += f", {snow_character}"
        day_to_draw.append(temp)

        temperature = day[17]
        drops = day[18]
        temp = ""
        if temperature:
            temp += f"{temperature}°"
        if drops:
            temp += f", {drops}"
        day_to_draw.append(temp)

        return day_to_draw

    def edit_day(self, day, cur_row):
        self.create_new_day(day=day, cur_row=cur_row)

    def call_setup(self):
        if self.setup_window is not None:
            self.setup_window.lift()
        else:
            self.setup_window = Tk()
            self.setup_window.title('Настройки')
            Label(self.setup_window, width=40, text="Путь к шаблону:").grid(row=0, column=0)
            template_entry = Entry(self.setup_window, fg='black', bg='white', width=40, justify=RIGHT)
            template_entry.grid(row=0, column=1)
            if self.template_way != './template.docx':
                template_entry.insert(0, self.template_way)
            b = Button(self.setup_window, text="Обзор", width=20,
                       command=lambda: self.change_template_way(template_entry))
            b.grid(row=0, column=2)

            save_button = Button(self.setup_window, text="Сохранить", width=20,
                                 command=lambda: self.save_setup(template_entry))
            save_button.grid(row=1, column=0)

    def change_template_way(self, template_entry):
        new_way = fd.askopenfilename(initialdir=os.getcwd(),
                                     title="Укажите новый путь к шаблону",
                                     filetypes=[("Docx files", "*.docx")])
        template_entry.delete(0, 'end')
        template_entry.insert(0, new_way)

    def save_setup(self, template_entry):
        new_way = template_entry.get()
        if new_way:
            self.template_way = new_way
        ans = mb.askyesno(title="Success!", message="Настройки успешно сохранены. Закрыть окно?")
        if ans:
            self.setup_window.destroy()
            self.setup_window = None
        else:
            self.setup_window.lift()


a = MyApp()
a.draw_root()
a.draw_tab1()
a.draw_tab2()
a.draw_tab3()
a.root.mainloop()
