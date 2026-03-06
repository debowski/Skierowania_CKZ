import json
import os
import sys
import threading
import pythoncom
import tkinter as tk
from tkinter import StringVar, filedialog, messagebox
import openpyxl
import ttkbootstrap as ttkb
from docx2pdf import convert
from docxtpl import DocxTemplate
from concurrent.futures import ThreadPoolExecutor

sys.stdout.reconfigure(encoding="utf-8")

def convert_with_com(file_path):
    """Wrapper function to initialize COM for the conversion."""
    pythoncom.CoInitialize()
    try:
        convert(file_path)
    finally:
        pythoncom.CoUninitialize()

class DataManager:
    def __init__(self):
        self.workbook = None
        self.sheet = None
        self.headers = []
        self.data = []
        self.indices = {}

    def load_file(self, file_path):
        try:
            self.workbook = openpyxl.load_workbook(file_path, data_only=True)
            self.sheet = self.workbook.active
            self.headers = [cell.value for cell in self.sheet[1] if cell.value is not None]
            
            expected_columns = [
                "Imię", "Drugie imię", "Nazwisko", 
                "Dane oddziału", "PESEL", "Specjalność/Zawód"
            ]
            
            if not set(expected_columns).issubset(set(self.headers)):
                return False, "Plik nie zawiera wszystkich wymaganych kolumn."

            self.indices = {header: self.headers.index(header) for header in self.headers}
            
            self.data = []
            for row in self.sheet.iter_rows(min_row=2, values_only=True):
                if any(row):
                    # Ensure the row has enough elements to match headers
                    row_data = list(row) + [None] * (len(self.headers) - len(row))
                    self.data.append(dict(zip(self.headers, row_data)))
            
            return True, "Sukces"
        except Exception as e:
            return False, str(e)

    def get_zawody_by_klasa(self):
        zawody_klasy = {"1": set(), "2": set(), "3": set()}
        for row in self.data:
            oddzial = str(row.get("Dane oddziału", ""))
            if oddzial:
                klasa = oddzial.split()[0][0]
                zawod = row.get("Specjalność/Zawód")
                if klasa in zawody_klasy and zawod:
                    zawody_klasy[klasa].add(zawod)
        return {k: sorted(list(v)) for k, v in zawody_klasy.items()}

    def filter_data(self, klasa, zawod):
        filtered = []
        for row in self.data:
            oddzial = str(row.get("Dane oddziału", "")).lower()
            specjalnosc = str(row.get("Specjalność/Zawód", "")).lower()
            
            if (klasa.lower() in oddzial and zawod.lower() in specjalnosc):
                filtered.append(row)
        return filtered

class App:
    def __init__(self):
        self.data_manager = DataManager()
        self.app_dir = self.get_resource_path("")
        self.data_dir = os.path.join(self.app_dir, "Data")
        self.json_file_path = os.path.join(self.app_dir, "zawody.json")
        
        self.zawody_dict = self.load_json(self.json_file_path)
        
        self.root = ttkb.Window(themename="solar")
        self.root.title("Skierowania 0.40 - Optimized")
        self.root.columnconfigure(0, weight=0, minsize=500)
        self.root.columnconfigure(1, weight=1, minsize=400)
        self.root.rowconfigure(0, weight=1)
        
        self.zawody_by_klasa = {"1": [], "2": [], "3": []}
        self.current_plik = ""
        
        self.dodaj_widzety()
        self.credits()

    def get_resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)

    def load_json(self, file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                return json.load(file)
        except Exception as e:
            print(f"Błąd ładowania JSON: {e}")
            return {}

    def dodaj_widzety(self):
        self.frame = ttkb.Frame(self.root)
        self.frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        self.frame2 = ttkb.Frame(self.root, bootstyle="success")
        self.frame2.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)

        self.frame.columnconfigure((0, 1, 2), weight=1)
        self.frame2.columnconfigure(0, weight=1)
        self.frame2.rowconfigure(0, weight=1)

        self.btn_wyb_plik = ttkb.Button(
            self.frame, text="Wybierz plik", bootstyle="warning", command=self.otwarcie_pliku
        )
        self.btn_wyb_plik.grid(row=0, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)

        self.var_klasa = StringVar(value="1")
        for i in range(1, 4):
            rb = ttkb.Radiobutton(
                self.frame, text=f"Klasa {i}", variable=self.var_klasa, value=str(i),
                bootstyle="success-outline-toolbutton", command=self.update_combobox
            )
            rb.grid(row=1, column=i-1, sticky="nsew", padx=5, pady=5)

        self.combo_zawod = ttkb.Combobox(self.frame, bootstyle="success", state="readonly")
        self.combo_zawod.grid(row=2, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)
        self.combo_zawod.bind("<<ComboboxSelected>>", self.wypisanie_osob)

        # Date Entries
        self.add_label_entry(3, "Data wystawienia", "data_wystawienia")
        self.add_label_entry(4, "Data rozpoczęcia", "data_rozpoczecia")
        self.add_label_entry(5, "Data zakończenia", "data_zakonczenia")

        # Time Spinboxes
        ttkb.Label(self.frame, text="Godzina rozpoczęcia").grid(row=6, column=0, sticky="nsew", padx=5, pady=5)
        self.spin_godz = ttkb.Spinbox(self.frame, from_=0, to=23, justify="center", format="%02.0f", width=5)
        self.spin_godz.grid(row=6, column=1, sticky="nsew", padx=5, pady=5)
        self.spin_godz.set("08")
        
        self.spin_min = ttkb.Spinbox(self.frame, from_=0, to=59, justify="center", format="%02.0f", width=5)
        self.spin_min.grid(row=6, column=2, sticky="nsew", padx=5, pady=5)
        self.spin_min.set("00")

        # Buttons
        self.btn_wykaz = self.add_button(7, 0, "Utwórz wykaz", self.utworz_wykaz)
        self.btn_wykaz_pdf = self.add_button(7, 1, "Konwersja PDF", self.utworz_wykaz_pdf, bootstyle="dark")
        self.add_button(7, 2, "Folder Wykazy", self.otworz_folder_wykazy)

        self.btn_skier = self.add_button(8, 0, "Utwórz skierowania", self.utworz_skierowania)
        self.btn_skier_pdf = self.add_button(8, 1, "Konwersja PDF", self.utworz_skierowania_pdf, bootstyle="dark")
        self.add_button(8, 2, "Folder Skierowania", self.otworz_folder_skierowania)

        self.wynik = ttkb.Label(self.frame, text="Wynik", bootstyle="inverse-dark")
        self.wynik.grid(row=9, column=0, columnspan=3, sticky="sew", padx=5, pady=5)

        self.pole_tekstowe = tk.Text(self.frame2)
        self.pole_tekstowe.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)

    def add_label_entry(self, row, text, attr_name):
        ttkb.Label(self.frame, text=text).grid(row=row, column=0, sticky="nsew", padx=5, pady=5)
        ent = ttkb.DateEntry(self.frame, firstweekday=0, bootstyle="success")
        ent.grid(row=row, column=1, columnspan=2, sticky="nsew", padx=5, pady=5)
        setattr(self, attr_name, ent)

    def add_button(self, row, col, text, cmd, bootstyle="success"):
        btn = ttkb.Button(self.frame, text=text, bootstyle=bootstyle, command=cmd)
        btn.grid(row=row, column=col, sticky="nsew", padx=5, pady=5)
        return btn

    def update_combobox(self):
        klasa = self.var_klasa.get()
        values = self.zawody_by_klasa.get(klasa, [])
        self.combo_zawod["values"] = values
        if values:
            self.combo_zawod.current(0)
        else:
            self.combo_zawod.set("")
        self.wypisanie_osob()

    def otwarcie_pliku(self):
        path = filedialog.askopenfilename(
            title="Wybierz plik", initialdir=self.data_dir, 
            filetypes=(("Arkusze", "*.xlsx"), ("Wszystkie pliki", "*.*"))
        )
        if path:
            success, msg = self.data_manager.load_file(path)
            if success:
                self.current_plik = path
                self.btn_wyb_plik.configure(text=os.path.basename(path), bootstyle="success")
                self.zawody_by_klasa = self.data_manager.get_zawody_by_klasa()
                self.update_combobox()
            else:
                messagebox.showerror("Błąd", msg)
                self.btn_wyb_plik.configure(text="Niepoprawny plik", bootstyle="danger")

    def wypisanie_osob(self, event=None):
        if not self.current_plik:
            self.pole_tekstowe.delete(1.0, tk.END)
            self.pole_tekstowe.insert(tk.END, "Nie wybrano pliku")
            return
        
        filtered = self.data_manager.filter_data(self.var_klasa.get(), self.combo_zawod.get())
        tekst = "\n".join([f"{i+1}. {r['Imię']} {r.get('Drugie imię') or ''} {r['Nazwisko']}" for i, r in enumerate(filtered)])
        self.pole_tekstowe.delete(1.0, tk.END)
        self.pole_tekstowe.insert(tk.END, tekst or "Brak danych")

    def symbol_zawodu(self, zawod):
        return self.zawody_dict.get(zawod, "N/A")

    def utworz_wykaz(self):
        filtered = self.data_manager.filter_data(self.var_klasa.get(), self.combo_zawod.get())
        if not filtered: return

        template_path = os.path.join(self.app_dir, "Szablony", "szablon_wykaz_v3.docx")
        doc = DocxTemplate(template_path)
        
        lista_osob = "\n".join([f"{i+1}. {r['Imię']} {r['Nazwisko']}" for i, r in enumerate(filtered)])
        
        context = {
            "dataWyst": self.data_wystawienia.entry.get(),
            "zawod": self.combo_zawod.get(),
            "kodZawodu": self.symbol_zawodu(self.combo_zawod.get()),
            "dataRozp": self.data_rozpoczecia.entry.get(),
            "dataZako": self.data_zakonczenia.entry.get(),
            "godzRozp": f"{self.spin_godz.get()}:{self.spin_min.get()}",
            "stopien": self.var_klasa.get(),
            "tabela": lista_osob,
        }
        
        doc.render(context)
        output_dir = os.path.join(self.app_dir, "Data", "Wykazy")
        os.makedirs(output_dir, exist_ok=True)
        
        save_path = os.path.join(output_dir, f"{context['stopien']}_{context['zawod']}.docx")
        doc.save(save_path)
        
        self.wynik.configure(text=f"Utworzono wykaz: {len(filtered)} osób")
        self.update_pdf_button_count(self.btn_wykaz_pdf, output_dir)

    def utworz_skierowania(self):
        filtered = self.data_manager.filter_data(self.var_klasa.get(), self.combo_zawod.get())
        if not filtered: return

        template_path = os.path.join(self.app_dir, "Szablony", "szablon_skierowania_v3.docx")
        output_dir = os.path.join(self.app_dir, "Data", "Skierowania")
        os.makedirs(output_dir, exist_ok=True)

        count = 0
        for r in filtered:
            doc = DocxTemplate(template_path)
            context = {
                "dataWyst": self.data_wystawienia.entry.get(),
                "imie": r['Imię'],
                "drugie_imie": r.get('Drugie imię') or "",
                "nazwisko": r['Nazwisko'],
                "PESEL": r['PESEL'],
                "zawod": self.combo_zawod.get(),
                "kodZawodu": self.symbol_zawodu(self.combo_zawod.get()),
                "dataRozp": self.data_rozpoczecia.entry.get(),
                "dataZako": self.data_zakonczenia.entry.get(),
                "godzRozp": f"{self.spin_godz.get()}:{self.spin_min.get()}",
                "stopien": self.var_klasa.get(),
            }
            doc.render(context)
            
            filename = f"{context['stopien']}_{context['zawod']}{context['imie']}{context['nazwisko']}.docx"
            save_path = self.get_unique_path(output_dir, filename)
            doc.save(save_path)
            count += 1

        self.wynik.configure(text=f"Utworzono: {count} dokumentów")
        self.update_pdf_button_count(self.btn_skier_pdf, output_dir)

    def get_unique_path(self, directory, filename):
        base, ext = os.path.splitext(filename)
        path = os.path.join(directory, filename)
        counter = 1
        while os.path.exists(path):
            path = os.path.join(directory, f"{base}_{counter}{ext}")
            counter += 1
        return path

    def update_pdf_button_count(self, btn, directory):
        if os.path.exists(directory):
            count = len([f for f in os.listdir(directory) if f.endswith(".docx") and not f.startswith("~")])
            btn.configure(text=f"PDF: {count} plików")

    def utworz_pdf_batch(self, directory):
        if not os.path.exists(directory):
            print("Folder nie istnieje.")
            return
            
        files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith(".docx") and not f.startswith("~")]
        if not files: return
        
        def run_conversion():
            with ThreadPoolExecutor(max_workers=4) as executor:
                executor.map(convert_with_com, files)
            self.root.after(0, lambda: messagebox.showinfo("Info", "Konwersja zakończona"))

        threading.Thread(target=run_conversion, daemon=True).start()

    def utworz_wykaz_pdf(self):
        self.utworz_pdf_batch(os.path.join(self.app_dir, "Data", "Wykazy"))

    def utworz_skierowania_pdf(self):
        self.utworz_pdf_batch(os.path.join(self.app_dir, "Data", "Skierowania"))

    def otworz_folder_wykazy(self):
        path = os.path.join(self.app_dir, "Data", "Wykazy")
        if not os.path.exists(path): os.makedirs(path)
        os.startfile(path)

    def otworz_folder_skierowania(self):
        path = os.path.join(self.app_dir, "Data", "Skierowania")
        if not os.path.exists(path): os.makedirs(path)
        os.startfile(path)

    def credits(self):
        license_text = "Ten program jest udostępniony na zasadach Licencji MIT.\n\n"
        disclaimer_text = "Autor tego programu nie ponosi odpowiedzialności za ewentualne szkody wynikające z jego użytkowania.\n\n"
        
        folders_text = (
            f"Foldery aplikacji:\n{self.app_dir}\n"
            f"Dane:\n{os.path.join(self.app_dir, 'Data')}\n"
            f"Skierowania:\n{os.path.join(self.app_dir, 'Data', 'Skierowania')}\n"
            f"Wykazy:\n{os.path.join(self.app_dir, 'Data', 'Wykazy')}\n"
        )
        
        version = "0.40 Optimized"
        text = (
            f"Autor: Piotr Dębowski\n"
            f"Wersja: {version}\n\n"
            + license_text + disclaimer_text + folders_text +
            "\nOptymalizacje:\n"
            "- Cache'owanie danych Excel (szybsze działanie)\n"
            "- ThreadPoolExecutor dla konwersji PDF\n"
            "- Poprawiona obsługa błędów i ścieżek\n"
            "- Brak zamrażania GUI podczas długich operacji\n"
        )
        self.pole_tekstowe.delete("1.0", tk.END)
        self.pole_tekstowe.insert(tk.END, text)

if __name__ == "__main__":
    app = App()
    app.root.mainloop()
