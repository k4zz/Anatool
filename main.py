import sys
import getopt
import openpyxl
import logging
import queue
import tkinter as tk
from tkinter import ttk, VERTICAL, HORIZONTAL, N, S, E, W, Label, Button, Entry, Checkbutton, IntVar
from tkinter import filedialog as fd
from tkinter.scrolledtext import ScrolledText

logger = logging.getLogger(__name__)
debug_state = any


class QueueHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(record)


class Protokol:
    def __init__(self, number, names):
        self.number = number
        self.names = names

    def __repr__(self):
        return f"Protokol(number={self.number}, names={self.names})"

    def add_names(self, name):
        self.names.extend(name)


class Zestawienie:
    def __init__(self, name, numbers):
        self.name = name
        self.numbers = numbers

    def __repr__(self):
        return f"Zestawienie(name='{self.name}', numbers='{self.numbers})"

    def add_positions(self, positions):
        self.numbers.extend(positions)


class PathUI:
    def __init__(self, frame):
        self.protokol_path = tk.StringVar()
        self.zestawienie_path = tk.StringVar()
        self.frame = frame
        Label(self.frame, text='Protokół').grid(column=0, row=0, sticky=W)
        Label(self.frame, text='Zestawienie').grid(column=0, row=1, sticky=W)
        Entry(self.frame, textvariable=self.protokol_path, width=60).grid(column=1, row=0, sticky=(W, E))
        Entry(self.frame, textvariable=self.zestawienie_path, width=60).grid(column=1, row=1, sticky=(W, E))
        Button(self.frame, text='...', command=self.open_protokol).grid(column=2, row=0, sticky=E)
        Button(self.frame, text='...', command=self.open_zestawienie).grid(column=2, row=1, sticky=E)
        Button(self.frame, text='Analiza', command=self.analyze).grid(column=0, row=2, columnspan=3, sticky=(W, E))

    def open_protokol(self):
        filetypes = (
            ('Excel', '*.xlsx'),
            ('All files', '*.*')
        )

        self.protokol_path.set(fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes))

    def open_zestawienie(self):
        filetypes = (
            ('Excel', '*.xlsx'),
            ('All files', '*.*')
        )

        self.zestawienie_path.set(fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes))

    def analyze(self):
        ready = True
        if self.protokol_path.get() == "":
            logger.log(logging.ERROR, "Brakująca ścieżka do protokołu")
            ready = False
        elif not self.protokol_path.get().endswith('.xlsx'):
            logger.log(logging.ERROR, "Podana sciezka dla protokolu nie jest plikiem excel (.xlsx)")
            ready = False

        if self.zestawienie_path.get() == "":
            logger.log(logging.ERROR, "Brakująca ścieżka do zestawienia")
            ready = False
        elif not self.zestawienie_path.get().endswith('.xlsx'):
            logger.log(logging.ERROR, "Podana sciezka dla zestawienia nie jest plikiem excel (.xlsx)")
            ready = False

        if ready:
            Analyzer(self.protokol_path.get(), self.zestawienie_path.get())


class ConsoleUI:
    def __init__(self, frame):
        self.frame = frame
        self.scrolled_text = ScrolledText(frame, state='disabled', height=30, width=150)
        self.scrolled_text.grid(row=0, column=0, sticky=(N, S, W, E))
        self.cb_debug = Checkbutton(frame, text='Debug', variable=debug_state).grid(row=1, column=0, sticky=W)
        self.log_queue = queue.Queue()
        self.queue_handler = QueueHandler(self.log_queue)
        logger.addHandler(self.queue_handler)
        self.frame.after(100, self.poll_log_queue)

    def clear_console(self):
        self.scrolled_text.delete(1.0, tk.END)

    def display(self, record):
        msg = self.queue_handler.format(record)
        self.scrolled_text.configure(state='normal')
        self.scrolled_text.insert(tk.END, msg + '\n', record.levelname)
        self.scrolled_text.configure(state='disabled')
        # Autoscroll to the bottom
        self.scrolled_text.yview(tk.END)

    def poll_log_queue(self):
        while True:
            try:
                record = self.log_queue.get(block=False)
            except queue.Empty:
                break
            else:
                self.display(record)
        self.frame.after(100, self.poll_log_queue)


class App:

    def __init__(self, root):
        self.root = root
        root.title('Anatool')
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        vertical_pane = ttk.PanedWindow(self.root, orient=VERTICAL)
        vertical_pane.grid(row=0, column=0, sticky="nsew")
        horizontal_pane = ttk.PanedWindow(vertical_pane, orient=HORIZONTAL)
        vertical_pane.add(horizontal_pane)

        path_frame = ttk.LabelFrame(horizontal_pane, text="Ścieżki")
        path_frame.columnconfigure(1, weight=1)
        horizontal_pane.add(path_frame, weight=1)

        console_frame = ttk.LabelFrame(horizontal_pane, text="Log")
        console_frame.columnconfigure(0, weight=1)
        console_frame.rowconfigure(0, weight=1)
        horizontal_pane.add(console_frame, weight=1)

        self.path = PathUI(path_frame)
        self.console = ConsoleUI(console_frame)


def Cmd(argv):
    first = ''
    second = ''

    try:
        opts, args = getopt.getopt(argv, "hp:z:", ["protokol=", "zestawienie="])
    except getopt.GetoptError:
        print('main.py -p <sciezka_do_protokolu> -z <sciezka_do_zestawienia>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print('main.py -p <sciezka_do_protokolu> -z <sciezka_do_zestawienia>')
            sys.exit()
        elif opt in ("-p", "--protokol"):
            if arg.endswith('.xlsx'):
                first = arg
            else:
                print("Podana sciezka dla protokolu nie jest plikiem excel (.xlsx)")
                sys.exit()
        elif opt in ("-z", "--zestawienie"):
            if arg.endswith('.xlsx'):
                second = arg
            else:
                print("Podana sciezka dla protokolu nie jest plikiem excel (.xlsx)")
                sys.exit()

    Analyzer(first, second)

class Analyzer:
    def __init__(self, protokol_path, zestawienie_path):
        self.protokol_path = protokol_path
        self.zestawienie_path = zestawienie_path

        self.sheet_protokol = any
        self.sheet_zestawienie = any

        self.zestawienie = {}
        self.protokol = {}

        logger.log(logging.INFO, "----Analiza----")
        self.get_sheets()
        self.get_objects()
        self.analyze()
        logger.log(logging.INFO, "----Koniec analizy----")

    def get_sheets(self):
        try:
            wb_protokol = openpyxl.open(self.protokol_path)
        except:
            logger.log(logging.ERROR, "Nie można otworzyć pliku protokołu")
            sys.exit()

        try:
            wb_zestawienie = openpyxl.open(self.zestawienie_path)
        except:
            logger.log(logging.ERROR, "Nie można otworzyć pliku zestawienia")
            sys.exit()


        self.sheet_protokol = wb_protokol[wb_protokol.sheetnames[0]]
        self.sheet_zestawienie = wb_zestawienie[wb_zestawienie.sheetnames[0]]

    def get_objects(self):
        # try:
        rowsFromSheet = self.sheet_protokol.iter_rows()
        rows = iter(rowsFromSheet)
        for row in rows:
            number = str(row[0].value)
            if number == 'None':
                continue

            str_names = row[2].value
            split_names = str_names.split('\n')

            # fix for spaces
            temp_list = []
            for name in split_names:
                if name.endswith(' '):
                    name = name.rstrip()
                if name == '':
                    continue
                temp_list.append(name)
            split_names = temp_list

            if number in self.protokol:
                self.protokol[number].add_names(split_names)
            else:
                self.protokol[number] = Protokol(number, split_names)
        # except:
        #     print("Unexpected error:", sys.exc_info()[0])
        #     logger.log(logging.ERROR, "Błąd parsowania pliku protokołu")

        currentRow = 0
        try:
            rowsFromSheet = self.sheet_zestawienie.iter_rows()
            rows = iter(rowsFromSheet)
            for row in rows:
                currentRow += 1
                name = row[1].value
                strPlotNumbers = row[3].value
                listPlotNumbers = strPlotNumbers.split(", ")
                if name in self.zestawienie:
                    self.zestawienie[name].add_positions(listPlotNumbers)
                else:
                    self.zestawienie[name] = Zestawienie(row[1].value, listPlotNumbers)
        except:
            logger.log(logging.ERROR, "Błąd parsowania pliku zestawienia dla lini  " + currentRow)

    def analyze(self):
        for num, pozycja in self.protokol.items():
            for nazwisko in pozycja.names:
                not_found = -2

                if nazwisko in self.zestawienie:
                    not_found = -1

                    if num in self.zestawienie[nazwisko].numbers:
                        not_found = 0

                if not_found == -1:
                    msg = "Brakujaca pozycja " + pozycja.number + " dla nazwiska " + nazwisko + " w zestawieniu."
                    logger.log(logging.ERROR, msg)
                elif not_found == -2:
                    msg = "Brakujace nazwisko " + nazwisko + " w zestawieniu" + "; Pozycja w protokole: " + pozycja.number + "."
                    logger.log(logging.ERROR, msg)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        Cmd(sys.argv[1:])
    else:
        logging.basicConfig(level=logging.DEBUG)
        root = tk.Tk()
        debug_state = IntVar()
        app = App(root)
        app.root.mainloop()
