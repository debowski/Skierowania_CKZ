import tkinter as tk
from tkinter import messagebox
import os
# import re

class App:
    def __init__(self, master):
        self.master = master
        self.main_frame = tk.Frame(master)
        
        # Kolorowanie kadrowe głównej przeglądy
        self.backgroundColor = "#f0f0f0"
        self.secondaryColor = "#ffffff"
        
        # Inicjalizacja pulpitu
        self.pult = tk.Label(
            master,
            text="System obsługi dokumentów",
            font=('Arial', 20, 'bold'),
            bg=self.backgroundColor,
            fg='black'
        )
        self.pult.pack(expand=True, fill=tk.X)
        
        # Tworzenie obszarów dla sekcji kredytów
        self.credits_frame = tk.LabelFrame(self.master, text="Credyty")
        self.credits_frame.pack(pady=20)
        
        # Metody potrzebne dla funkcjonalności
        self.open_folder_wykazy = lambda: self.open_folder("Wykazy")
        self.open_folder_skierowania = lambda: self.open_folder("Skierowania")
        self.create_pdf = lambda folder_name: self.create_folder_pdf(folder_name)
        
        # Tworzenie obszarów dla sekcji otworzonego folderu
        self.open_folder_frame = tk.LabelFrame(self.master, text="Otwórne folderów")
        self.open_folder_frame.pack(pady=10)
        
        # Dodanie przycisków otwierających folderów
        self.btn_wykazy = tk.Button(
            self.open_folder_frame,
            text="Wykazy",
            command=self.open_folder_wykazy,
            bg='blue',
            fg='white'
        )
        self.btn_skierowania = tk.Button(
            self.open_folder_frame,
            text="Skierowania",
            command=self.open_folder_skierowania,
            bg='blue',
            fg='white'
        )
        
        # Dodanie obszarów dla tworzenia PDF
        self.create_pdf_frame = tk.LabelFrame(self.master, text="Tworzenie PDF")
        self.create_pdf_frame.pack(pady=30)
        
        # Tworzenie przycisków do tworzenia PDF
        self.btn_create_wykaz = tk.Button(
            self.create_pdf_frame,
            text="Twórz Wykazy",
            command=lambda: self.utworz_wykaz_pdf(),
            bg='green',
            fg='white'
        )
        self.btn_create_skierowanie = tk.Button(
            self.create_pdf_frame,
            text="Twórz Skierowania",
            command=lambda: self.utworz_skierowania_pdf(),
            bg='green',
            fg='white'
        )
        
        # Tworzenie obszarów dla sekcji informacji
        self.info_frame = tk.LabelFrame(self.master, text="Informacja")
        self.info_frame.pack(pady=50)
        
        # Dodanie obszarów dla sekcji informacji
        self.info_label = tk.Label(
            self.info_frame,
            text=self.credits(),
            justify=tk.LEFT,
            wrap=tk.WORD,
            font=('Arial', 10, 'italic')
        )
        self.info_label.pack()
        
        # Tworzenie przycisków otwierających sekcję informacji
        self.btn_show_info = tk.Button(
            master,
            text="Pokaż więcej informacji",
            command=lambda: self.master.focus_set(),
            bg='blue',
            fg='white'
        )
        self.btn_show_info.pack(pady=10, padx=20)
        
    def credits(self):
        """Zwraca napisy z sekcji kredytów"""
        return (
            "Autor: Jan Kowalski\n"
            f"Wersja: {self.get_version()}\n"
            "Lizenzja: Liczba publiczna (MIT)\n"
            "\nDisclaimer:\nAplikacja nie posiadaje załógów na rzecz użytkownika."
        )
    
    def get_version(self):
        """Pobiera wersję aplikacji"""
        return 1.0
    
    def open_folder(self, folder_name):
        """Otwierające funkcjonalność otwierająego folderu"""
        file_path = os.path.join(os.path.dirname(__file__), folder_name)
        if not os.path.exists(file_path):
            messagebox.showerror("Błąd", f"Folder '{folder_name}' nie istnieje!")
        else:
            os.startfile(file_path, 'nk')
    
    def utworz_wykaz_pdf(self):
        """Tworzy Wykazy w formacie PDF"""
        file_path = os.path.join(os.path.dirname(__file__), "wykazy", "nowe_wyklady.pdf")
        if not os.path.exists(file_path):
            file_path = os.path.join(file_path, "nowe_wyklady.pdf")
            with open(file_path, 'w') as f:
                f.write("Treści do wykazu nie zostały dodane!")
        else:
            os.startfile(file_path, 'r')
    
    def utworz_skierowania_pdf(self):
        """Tworzy Skierowania w formacie PDF"""
        file_path = os.path.join(os.path.dirname(__file__), "skierowania", "nowe_skierowania.pdf")
        if not os.path.exists(file_path):
            file_path = os.path.join(file_path, "nowe_skierowania.pdf")
            with open(file_path, 'w') as f:
                f.write("Treści do skierowań nie zostały dodane!")
        else:
            os.startfile(file_path, 'r')
    
    def create_folder_pdf(self, folder_name):
        """Tworzy nowy plik PDF w wybranym folderze"""
        folder_path = os.path.join(os.path.dirname(__file__), folder_name)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
            with open(os.path.join(folder_path, "nowe_dane.pdf"), 'w') as f:
                f.write("Nowe dane nie zostały dodane!")
        else:
            pass
    
    def run(self):
        """Główna funkcja aplikacji"""
        self.run()

if __name__ == "__main__":
    app = App(master=None)
    app.run()