import sys
import getopt
import openpyxl
import logging
import queue
import tkinter as tk
from tkinter import ttk, VERTICAL, HORIZONTAL, N, S, E, W, Label, Button, Entry, IntVar
from tkinter import filedialog as fd
from tkinter.scrolledtext import ScrolledText

# Logger
logger = logging.getLogger(__name__)


class QueueHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(record)


# Data containers
class Protocol:
    def __init__(self, number, names, row):
        self.number = number
        self.names = names
        self.row = row

    def __repr__(self):
        return f"Protokol(number={self.number}, names={self.names})"

    def add_names(self, name):
        self.names.extend(name)


class Collation:
    def __init__(self, name, numbers, row):
        self.name = name
        self.numbers = numbers
        self.row = row
        self.amount_of_checked = 0

    def __repr__(self):
        return f"Zestawienie(name='{self.name}', numbers='{self.numbers})"

    def add_positions(self, positions):
        self.numbers.extend(positions)

    def number_checked(self):
        self.amount_of_checked += 1


# UI
class PathUI:
    def __init__(self, frame):
        self.protocol_path = tk.StringVar()
        self.collation_path = tk.StringVar()
        self.frame = frame
        Label(self.frame, text='Protokół').grid(column=0, row=0, sticky=W)
        Label(self.frame, text='Zestawienie').grid(column=0, row=1, sticky=W)
        Entry(self.frame, textvariable=self.protocol_path, width=60).grid(column=1, row=0, sticky=(W, E))
        Entry(self.frame, textvariable=self.collation_path, width=60).grid(column=1, row=1, sticky=(W, E))
        Button(self.frame, text='...', command=self.open_protocol).grid(column=2, row=0, sticky=E)
        Button(self.frame, text='...', command=self.open_collation).grid(column=2, row=1, sticky=E)
        Button(self.frame, text='Analiza', command=self.analyze).grid(column=0, row=2, columnspan=3, sticky=(W, E))

    def open_protocol(self):
        filetypes = (
            ('Excel', '*.xlsx'),
            ('All files', '*.*')
        )

        self.protocol_path.set(fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes))

    def open_collation(self):
        filetypes = (
            ('Excel', '*.xlsx'),
            ('All files', '*.*')
        )

        self.collation_path.set(fd.askopenfilename(
            title='Open a file',
            initialdir='/',
            filetypes=filetypes))

    def analyze(self):
        ready = True
        if self.protocol_path.get() == "":
            logger.log(logging.ERROR, "Brakująca ścieżka do protokołu")
            ready = False
        elif not self.protocol_path.get().endswith('.xlsx'):
            logger.log(logging.ERROR, "Podana sciezka dla protokolu nie jest plikiem excel (.xlsx)")
            ready = False

        if self.collation_path.get() == "":
            logger.log(logging.ERROR, "Brakująca ścieżka do zestawienia")
            ready = False
        elif not self.collation_path.get().endswith('.xlsx'):
            logger.log(logging.ERROR, "Podana sciezka dla zestawienia nie jest plikiem excel (.xlsx)")
            ready = False

        if ready:
            Analyzer(self.protocol_path.get(), self.collation_path.get())


class ConsoleUI:
    def __init__(self, frame):
        self.frame = frame
        self.scrolled_text = ScrolledText(frame, state='disabled', height=30, width=150)
        self.scrolled_text.grid(row=0, column=0, sticky=(N, S, W, E))
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
        self.root.title('Anatool')
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
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


class Cmd:
    def __init__(self, argv):
        self.protocol = ''
        self.collation = ''
        self.argv = argv

        try:
            opts, args = getopt.getopt(self.argv, "hp:z:", ["protokol=", "zestawienie="])
        except getopt.GetoptError:
            print('main.py -p <sciezka_do_protokolu> -z <sciezka_do_zestawienia>')
            sys.exit(2)
        for opt, arg in opts:
            if opt == '-h':
                print('main.py -p <sciezka_do_protokolu> -z <sciezka_do_zestawienia>')
                sys.exit()
            elif opt in ("-p", "--protokol"):
                if arg.endswith('.xlsx'):
                    self.protocol = arg
                else:
                    print("Podana sciezka dla protokolu nie jest plikiem excel (.xlsx)")
                    sys.exit()
            elif opt in ("-z", "--zestawienie"):
                if arg.endswith('.xlsx'):
                    self.collation = arg
                else:
                    print("Podana sciezka dla protokolu nie jest plikiem excel (.xlsx)")
                    sys.exit()

        Analyzer(self.protocol, self.collation)


class Analyzer:
    def __init__(self, protocol_path, collation_path):
        self.protocol_path = protocol_path
        self.collation_path = collation_path

        self.sheet_protocol = any
        self.sheet_collation = any

        self.collation = {}
        self.protocol = {}

        logger.log(logging.INFO, "----Analiza----")
        if not self.get_sheets():
            logger.log(logging.INFO, "----Analiz zakończona błędem----")
            return

        if not self.get_objects():
            logger.log(logging.INFO, "----Analiz zakończona błędem----")
            return

        if not self.analyze():
            logger.log(logging.INFO, "----Analiz zakończona błędem----")
            return

        logger.log(logging.INFO, "----Koniec analizy----")

    def get_sheets(self):
        try:
            wb_protocol = openpyxl.open(self.protocol_path)
        except:
            logger.log(logging.ERROR, "Nie można otworzyć pliku protokołu")
            return False

        try:
            wb_collation = openpyxl.open(self.collation_path)
        except:
            logger.log(logging.ERROR, "Nie można otworzyć pliku zestawienia")
            return False

        self.sheet_protocol = wb_protocol[wb_protocol.sheetnames[0]]
        self.sheet_collation = wb_collation[wb_collation.sheetnames[0]]

        return True

    def get_objects(self):

        # Protocol
        rows_from_sheet = self.sheet_protocol.iter_rows()
        rows = iter(rows_from_sheet)
        for row in rows:
            number = str(row[0].value)
            if number == 'None':
                continue

            str_names = row[2].value
            # Split string to list
            split_names = str_names.split('\n')
            # Remove white spaces
            split_names = [n.strip() for n in split_names]
            # Remove empty strings (new lines)
            split_names = [n for n in split_names if n != '']

            if number in self.protocol:
                new_list = self.protocol[number].names
                new_list.extend(split_names)
                new_list = list(dict.fromkeys(new_list))
                self.protocol[number].names = new_list
            else:
                self.protocol[number] = Protocol(number, split_names, row[0].row)

        # Collation
        current_row = 0
        rows_from_sheet = self.sheet_collation.iter_rows()
        rows = iter(rows_from_sheet)
        for row in rows:
            current_row += 1
            try:
                name = row[1].value.rstrip()
                str_plot_numbers = row[2].value
                list_plot_numbers = str_plot_numbers.split(",")
                list_plot_numbers = [n.strip() for n in list_plot_numbers]
                if name in self.collation:
                    self.collation[name].add_positions(list_plot_numbers)
                else:
                    self.collation[name] = Collation(row[1].value, list_plot_numbers, row[0].row)
            except:
                logger.log(logging.ERROR, "Błąd parsowania pliku zestawienia dla wiersza " + str(current_row))

        return True

    def analyze(self):
        for num, position in self.protocol.items():
            for name in position.names:
                not_found = -2

                if name in self.collation:
                    not_found = -1

                    if num in self.collation[name].numbers:
                        not_found = 0
                        self.collation[name].number_checked()

                if not_found == -1:
                    msg = "Brakujaca pozycja " + position.number + " dla nazwiska " + name + " w zestawieniu."
                    logger.log(logging.ERROR, msg)
                elif not_found == -2:
                    msg = "Brakujace nazwisko " + name + " w zestawieniu" + "; " \
                          + "Pozycja w protokole: " + position.number + "."
                    logger.log(logging.ERROR, msg)

        for name, collation in self.collation.items():
            if len(collation.numbers) > collation.amount_of_checked:
                logger.log(logging.ERROR, "Nadmiarowa pozycja w zestawieniu; Nazwa: " + collation.name + "; Linia: " +
                           str(collation.row))


if __name__ == "__main__":
    if len(sys.argv) > 1:
        Cmd(sys.argv[1:])
    else:
        logging.basicConfig(level=logging.DEBUG)
        root = tk.Tk()
        debug_state = IntVar()
        app = App(root)
        app.root.mainloop()
