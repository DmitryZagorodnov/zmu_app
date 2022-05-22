import os
import matplotlib
from tkinter import *
from tkinter import filedialog as fd
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


class App:

    def __init__(self):
        self.root = Tk()
        self.entr_list = []
        self.fields_values = {}
        self.aliases = {}

        self.ent_chosen_tracks: Entry
        self.tracksfile = ""
        self.ent_chosen_waypoints: Entry
        self.waypointsfile = ""
        self.context = {}
        self.dates = {}
        self.init_forms()
        self.root.mainloop()

    def choose_track(self):
        filetype = [('gpx files', '*.gpx')]
        self.tracksfile = fd.askopenfilename(filetype=filetype)
        self.ent_chosen_tracks.delete(0, END)
        self.ent_chosen_tracks.insert(0, self.tracksfile)

    def choose_waypoints(self):
        filetype = [('gpx files', '*.gpx')]
        self.waypointsfile = fd.askopenfilename(filetype=filetype)
        self.ent_chosen_waypoints.delete(0, END)
        self.ent_chosen_waypoints.insert(0, self.waypointsfile)

    def init_values(self):
        for entry in self.entr_list:
            cont = entry.get()
            if cont is not None:
                self.fields_values[entry] = cont
            else:
                self.fields_values[entry] = ""

    def check_required_fields(self):
        return not (self.tracksfile == "" or self.waypointsfile == "")

    def parse_dates(self):
        for alias, value in self.dates.items():
            date = value.get_date()
            self.context[f"{alias}D"] = date.day
            self.context[f"{alias}M"] = date.month
            self.context[f"{alias}Y"] = date.year

    def create_doc(self):
        self.init_values()
        if not self.check_required_fields():
            mb.showerror(title='Error', message="Missed required fields")
            return -1
        else:
            pars = GpxParser()
            try:
                pars.parse_track(tracksfile=self.tracksfile)
            except ValueError:
                mb.showerror(title='Error', message="Your file doesn't contains any track")
                return -1
            try:
                pars.parse_waypoints(waypointsfile=self.waypointsfile)
            except ValueError:
                mb.showerror(title='Error', message="Your file doesn't contains any waypoints")
                return -1

            pars.parse()
            doc = DocxTemplate("template.docx")
            image = self.get_track()
            image.savefig('temp.png')
            self.context['image'] = InlineImage(doc, 'temp.png', width=Mm(100))

            cont = pars.prepare_context()
            self.parse_dates()
            for entry in self.entr_list:
                self.context[self.aliases[entry]] = self.fields_values[entry]
            for key, value in cont.items():
                self.context[key] = value
            doc.render(context=self.context)
            file_to_save = fd.asksaveasfilename(initialdir=os.getcwd(),
                                                title="Select the output file name",
                                                filetypes=[('Docx files', '.docx')])
            doc.save(f"{file_to_save}.docx")
            os.remove('./temp.png')
            mb.showinfo('Success!', 'Your file successfully created!')

    def get_track(self):
        self.tracksfile = self.ent_chosen_tracks.get()
        self.waypointsfile = self.ent_chosen_waypoints.get()
        if not self.tracksfile and not self.waypointsfile:
            mb.showerror(title='Missed files', message='You should fill any filename entry')
            return -1
        else:
            pars = GpxParser()
            f = Figure(figsize=(7, 7), dpi=100)
            f_plot = f.add_subplot(111)
            if self.tracksfile:
                try:
                    pars.parse_track(self.tracksfile)
                    f_plot.plot(pars.longs, pars.lats, color='red')
                except ValueError:
                    mb.showerror(title='Error', message="Your file doesn't contains any track")
                    return -1

            if self.waypointsfile:
                try:
                    pars.parse_waypoints(self.waypointsfile)
                except ValueError:
                    mb.showerror(title='Error', message="Your file doesn't contains any waypoints")
                    return -1
                for pt in pars.waypoints:
                    f_plot.plot(pt[1], pt[2], color='gray', marker='o')
                    f_plot.text(pt[1], pt[2], pt[0], rotation='horizontal', backgroundcolor='white',
                                bbox={'visible': True, 'facecolor': 'white'})
            return f

    def print_track(self):
        img = self.get_track()
        if img != -1:
            daughter = Tk()
            daughter.wm_title = 'Track'
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

    def init_forms(self):
        i = 0

        Label(text="Ведомость зимнего маршрутного учета").grid(row=i, column=0)
        i = i + 1

        Label(text="Маршрут №").grid(row=i, column=0)
        rn = ttk.Spinbox(from_=0, to=1000, width=17)
        self.entr_list.append(rn)
        self.aliases[rn] = "RN"
        rn.grid(row=i, column=1)
        i = i + 1

        Label(text="Субъект Российской Федерации").grid(row=i, column=0)
        srf = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(srf)
        self.aliases[srf] = "SRF"
        srf.grid(row=i, column=1)
        i = i + 1

        Label(text="Муниципальное образование (район)").grid(row=i, column=0)
        srf = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(srf)
        self.aliases[srf] = "MO"
        srf.grid(row=i, column=1)
        i = i + 1

        Label(text="Исследуемая территория").grid(row=i, column=0)
        it = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(it)
        self.aliases[it] = 'IT'
        it.grid(row=i, column=1)
        i = i + 1

        Label(text="Учетчик").grid(row=i, column=0)
        uch = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(uch)
        self.aliases[uch] = "UCH"
        uch.grid(row=i, column=1)
        i = i + 1

        Label(text="Дата окончания последней пороши:").grid(row=i, column=0)
        dp_date = DateEntry(fg="black", bg="white", width="17", justify=RIGHT)
        dp_date.delete(0, 'end')
        dp_date.grid(row=i, column=1)
        self.dates["DP"] = dp_date

        Label(text="Время:").grid(row=i, column=2)
        dp_h = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(dp_h)
        self.aliases[dp_h] = "DPH"
        dp_h.grid(row=i, column=3)
        Label(text="час.").grid(row=i, column=4)
        i = i + 1

        Label(text="Использование транспортного средства").grid(row=i, column=0)
        its = ttk.Combobox(values=["Да", "Нет"], width=17)
        self.entr_list.append(its)
        self.aliases[its] = "ITS"
        its.grid(row=i, column=1)
        Label(text="(да/нет)").grid(row=i, column=2)
        i = i + 1

        Label(text="Использование спутникового навигатора").grid(row=i, column=0)
        isn = ttk.Combobox(values=["Да", "Нет"], width=17)
        self.entr_list.append(isn)
        self.aliases[isn] = "ISN"
        isn.grid(row=i, column=1)
        Label(text="(да/нет)").grid(row=i, column=2)
        i = i + 1

        Label(text="Дата затирки").grid(row=i, column=0)
        dz_date = DateEntry(fg="black", bg="white", width="17", justify=RIGHT)
        dz_date.delete(0, 'end')
        dz_date.grid(row=i, column=1)
        self.dates['DZ'] = dz_date

        Label(text="Время начала").grid(row=i, column=2)
        dz_hb = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(dz_hb)
        self.aliases[dz_hb] = "DZHB"
        dz_hb.grid(row=i, column=3)
        Label(text="час.").grid(row=i, column=4)
        dz_mb = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(dz_mb)
        self.aliases[dz_mb] = "DZMB"
        dz_mb.grid(row=i, column=5)
        Label(text="мин., окончание").grid(row=i, column=6)
        dz_he = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(dz_he)
        self.aliases[dz_he] = "DZHE"
        dz_he.grid(row=i, column=7)
        Label(text="час.").grid(row=i, column=8)
        dz_me = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(dz_me)
        self.aliases[dz_me] = "DZME"
        dz_me.grid(row=i, column=9)
        Label(text="мин.").grid(row=i, column=10)
        i = i + 1

        Label(text="Дата учета следов").grid(row=i, column=0)
        dus_date = DateEntry(fg="black", bg="white", width="17", justify=RIGHT)
        dus_date.delete(0, 'end')
        dus_date.grid(row=i, column=1)

        Label(text="Время начала:").grid(row=i, column=2)
        dus_hb = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(dus_hb)
        self.aliases[dus_hb] = "DUSHB"
        dus_hb.grid(row=i, column=3)
        Label(text="час.").grid(row=i, column=4)
        dus_mb = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(dus_mb)
        self.aliases[dus_mb] = "DUSMB"
        dus_mb.grid(row=i, column=5)
        Label(text="мин., окончание").grid(row=i, column=6)
        dus_he = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(dus_he)
        self.aliases[dus_he] = "DUSHE"
        dus_he.grid(row=i, column=7)
        Label(text="час.").grid(row=i, column=8)
        dus_me = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(dus_me)
        self.aliases[dus_me] = "DUSME"
        dus_me.grid(row=i, column=9)
        Label(text="мин.").grid(row=i, column=10)
        i = i + 1

        Label(text="Высота снега").grid(row=i, column=0)
        sh = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(sh)
        self.aliases[sh] = "SH"
        sh.grid(row=i, column=1)
        Label(text="см.").grid(row=i, column=2)
        i = i + 1

        Label(text="Характер снега (рыхлый, плотный и др.)").grid(row=i, column=0)
        sc = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(sc)
        self.aliases[sc] = "SC"
        sc.grid(row=i, column=1)
        i = i + 1

        Label(text="Погода в день учета следов: температура").grid(row=i, column=0)
        wt = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(wt)
        self.aliases[wt] = "WT"
        wt.grid(row=i, column=1)
        i = i + 1
        Label(text="Погода в день учета следов: осадки").grid(row=i, column=0)
        wd = Entry(fg="black", bg="white", width="20", justify=RIGHT)
        self.entr_list.append(wd)
        self.aliases[wd] = "WD"
        wd.grid(row=i, column=1)
        i = i + 1

        for i in range(20):
            self.root.rowconfigure(i, weight=1)
            for j in range(20):
                self.root.columnconfigure(j, weight=1)

        Button(self.root, text="Выбрать файл с треком", command=self.choose_track, width=20).grid(row=i, column=0)
        self.ent_chosen_tracks = Entry(self.root, fg="black", bg="white", width="50", justify=LEFT)
        self.ent_chosen_tracks.grid(row=i, column=1, columnspan=3)
        i = i + 1

        Button(self.root, text="Выбрать файл с метками", command=self.choose_waypoints).grid(row=i, column=0)
        self.ent_chosen_waypoints = Entry(self.root, fg="black", bg="white", width="50", justify=LEFT)
        self.ent_chosen_waypoints.grid(row=i, column=1, columnspan=3)
        i = i + 1

        Button(text='Составить ведомость', command=self.create_doc, width=20).grid(row=i, column=0)
        i = i + 1
        Button(text='Отобразить трек', command=self.print_track, width=20).grid(row=i, column=0)
