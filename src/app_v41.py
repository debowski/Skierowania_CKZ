import json
import os
import sys
import threading
import datetime
from typing import Dict, List, Tuple, Optional, Any

import pythoncom
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import openpyxl
import subprocess
import tempfile
import uuid
from docxtpl import DocxTemplate
from concurrent.futures import ThreadPoolExecutor

if sys.stdout is not None:
    sys.stdout.reconfigure(encoding="utf-8")

# Configure CustomTkinter appearance
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class Theme:
    """Centralized theme colors representing the requested dark-teal aesthetic."""
    BG = "#041419"
    FRAME_BG = "#09242b"
    PRIMARY = "#1e797a"
    PRIMARY_HOVER = "#155a5b"
    ACCENT = "#D05A3B"
    ACCENT_HOVER = "#A5452C"
    ENTRY_BG = "#0b313b"
    TEXT = "#E0E0E0"
    BORDER = "#144853"

def convert_with_libreoffice(file_path: str, output_dir: str):
    """Konwertuje plik DOCX na PDF używając LibreOffice w trybie headless."""
    tmp_env = None
    try:
        # Typowa ścieżka do LibreOffice w systemie Windows
        libreoffice_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
        if not os.path.exists(libreoffice_path):
            libreoffice_path = r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            
        if not os.path.exists(libreoffice_path):
            raise FileNotFoundError("Nie znaleziono instalacji LibreOffice (wymagane do konwersji PDF).")

        # We do NOT overwrite os.environ["USERPROFILE"] because Windows needs it for Desktop and Dialog boxes
        # Instead, we just pass -env:UserInstallation to libreoffice directly
        tmp_env = os.path.join(tempfile.gettempdir(), f"libreoffice_profile_{uuid.uuid4().hex}")
        env_arg = f"-env:UserInstallation=file:///{tmp_env.replace('\\', '/')}"

        # Uruchomienie z dodatkową flagą środowiska i z suppressowaniem okien
        subprocess.run(
            [libreoffice_path, env_arg, '--headless', '--nologo', '--nofirststartwizard', '--convert-to', 'pdf', '--outdir', output_dir, file_path],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL
        )
    except subprocess.CalledProcessError as e:
        raise Exception(f"Składnik LibreOffice zwrócił błąd. Status procesu: {e.returncode}")
    except Exception as e:
        raise Exception(f"Błąd konwersji pliku {os.path.basename(file_path)}: {str(e)}")
    finally:
        # Cleanup temp profile dir if it exists
        if tmp_env and os.path.exists(tmp_env):
            try:
                import shutil
                shutil.rmtree(tmp_env, ignore_errors=True)
            except:
                pass

class DataManager:
    """Class responsible for handling Excel data loading and processing."""
    def __init__(self):
        self.workbook: Optional[openpyxl.Workbook] = None
        self.sheet = None
        self.headers: List[str] = []
        self.data: List[Dict[str, Any]] = []
        self.indices: Dict[str, int] = {}

    def load_file(self, file_path: str) -> Tuple[bool, str]:
        try:
            self.workbook = openpyxl.load_workbook(file_path, data_only=True)
            self.sheet = self.workbook.active
            self.headers = [str(cell.value).strip() for cell in self.sheet[1] if cell.value is not None]
            
            expected_columns = [
                "Imię", "Drugie imię", "Nazwisko", 
                "Dane oddziału", "PESEL", "Specjalność/Zawód"
            ]
            
            if not set(expected_columns).issubset(set(self.headers)):
                return False, "Plik nie zawiera wszystkich wymaganych kolumn (Imię, Drugie imię, Nazwisko, Dane oddziału, PESEL, Specjalność/Zawód)."

            self.indices = {header: self.headers.index(header) for header in self.headers}
            
            self.data = []
            for row in self.sheet.iter_rows(min_row=2, values_only=True):
                if any(row):
                    row_data = list(row) + [None] * (len(self.headers) - len(row))
                    self.data.append(dict(zip(self.headers, row_data)))
            
            return True, "Sukces"
        except Exception as e:
            return False, f"Błąd podczas ładowania pliku: {e}"

    def get_zawody_by_klasa(self) -> Dict[str, List[str]]:
        zawody_klasy = {"1": set(), "2": set(), "3": set()}
        for row in self.data:
            oddzial = str(row.get("Dane oddziału", ""))
            if oddzial:
                klasa = oddzial.split()[0][0]
                zawod = row.get("Specjalność/Zawód")
                if klasa in zawody_klasy and zawod:
                    zawody_klasy[klasa].add(zawod)
        return {k: sorted(list(v)) for k, v in zawody_klasy.items()}

    def filter_data(self, klasa: str, zawod: str) -> List[Dict[str, Any]]:
        filtered = []
        for row in self.data:
            oddzial = str(row.get("Dane oddziału", "")).lower()
            specjalnosc = str(row.get("Specjalność/Zawód", "")).lower()
            
            if (klasa.lower() in oddzial and zawod.lower() in specjalnosc):
                filtered.append(row)
        return filtered

class App(ctk.CTk):
    """Main Application GUI using CustomTkinter."""
    def __init__(self):
        super().__init__()
        
        self.data_manager = DataManager()
        self.app_dir = self.get_resource_path("")
        self.data_dir = os.path.join(self.app_dir, "Data")
        self.json_file_path = os.path.join(self.app_dir, "zawody.json")
        
        self.zawody_dict = self.load_json(self.json_file_path)
        
        self.title("Skierowania 0.41 - Optimized (CustomTkinter)")
        self.geometry("1000x550")
        self.configure(fg_color=Theme.BG)
        
        # Grid layout for the main window
        self.grid_columnconfigure(0, weight=0, minsize=550)
        self.grid_columnconfigure(1, weight=1, minsize=400)
        self.grid_rowconfigure(0, weight=1)
        
        self.zawody_by_klasa = {"1": [], "2": [], "3": []}
        self.current_plik = ""
        
        self._setup_ui()
        self.credits()

    def get_resource_path(self, relative_path: str) -> str:
        try:
            base_path = sys._MEIPASS
        except AttributeError:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)

    def load_json(self, file_path: str) -> dict:
        if not os.path.exists(file_path):
            return {}
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                return json.load(file)
        except Exception as e:
            print(f"Błąd ładowania JSON: {e}")
            return {}

    def _setup_ui(self):
        # Left Panel (Controls)
        self.left_frame = ctk.CTkFrame(
            self, fg_color=Theme.FRAME_BG, 
            corner_radius=10, border_width=1, border_color=Theme.BORDER
        )
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=(15, 10), pady=15)
        self.left_frame.grid_columnconfigure((0, 1, 2), weight=1)
        
        # Right Panel (Information & Output)
        self.right_frame = ctk.CTkFrame(
            self, fg_color=Theme.FRAME_BG, 
            corner_radius=10, border_width=1, border_color=Theme.BORDER
        )
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 15), pady=15)
        self.right_frame.grid_columnconfigure(0, weight=1)
        self.right_frame.grid_rowconfigure(0, weight=1)
        
        self._create_left_panel()
        
        self.pole_tekstowe = ctk.CTkTextbox(
            self.right_frame, fg_color="transparent", 
            text_color=Theme.TEXT, font=ctk.CTkFont(size=14), wrap="word"
        )
        self.pole_tekstowe.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

    def _create_left_panel(self):
        pad = {"padx": 8, "pady": 6}
        
        # 1. File Selection Button
        self.btn_wyb_plik = ctk.CTkButton(
            self.left_frame, text="Wybierz plik", 
            fg_color=Theme.ACCENT, hover_color=Theme.ACCENT_HOVER, 
            text_color="white", font=ctk.CTkFont(size=14, weight="bold"),
            command=self.otwarcie_pliku
        )
        self.btn_wyb_plik.grid(row=0, column=0, columnspan=3, sticky="nsew", padx=8, pady=10)

        # 2. Segmented Button for 'Klasa'
        self.var_klasa = ctk.StringVar(value="Klasa 1")
        self.seg_klasa = ctk.CTkSegmentedButton(
            self.left_frame, 
            values=["Klasa 1", "Klasa 2", "Klasa 3"],
            variable=self.var_klasa,
            command=self.update_combobox_from_seg,
            selected_color=Theme.PRIMARY,
            selected_hover_color=Theme.PRIMARY_HOVER,
            unselected_color=Theme.ENTRY_BG,
            unselected_hover_color=Theme.PRIMARY_HOVER,
            font=ctk.CTkFont(size=13)
        )
        self.seg_klasa.grid(row=1, column=0, columnspan=3, sticky="nsew", **pad)

        # 3. ComboBox for Zawód
        self.combo_zawod = ctk.CTkComboBox(
            self.left_frame, state="readonly",
            fg_color=Theme.ENTRY_BG, border_color=Theme.PRIMARY,
            button_color=Theme.PRIMARY, button_hover_color=Theme.PRIMARY_HOVER,
            dropdown_fg_color=Theme.ENTRY_BG, dropdown_hover_color=Theme.PRIMARY,
            command=self.wypisanie_osob, font=ctk.CTkFont(size=13)
        )
        self.combo_zawod.grid(row=2, column=0, columnspan=3, sticky="nsew", **pad)
        self.combo_zawod.set("")

        # 4. Dates
        default_date = datetime.datetime.now().strftime("%d/%m/%y")
        self.data_wystawienia = self._add_label_entry(3, "Data wystawienia", default_date)
        self.data_rozpoczecia = self._add_label_entry(4, "Data rozpoczęcia", default_date)
        self.data_zakonczenia = self._add_label_entry(5, "Data zakończenia", default_date)

        # 5. Time Spinboxes equivalent
        ctk.CTkLabel(
            self.left_frame, text="Godzina rozpoczęcia", 
            anchor="w", font=ctk.CTkFont(size=13)
        ).grid(row=6, column=0, sticky="ew", **pad)
        
        self.spin_godz = ctk.CTkOptionMenu(
            self.left_frame, values=[f"{i:02d}" for i in range(24)],
            fg_color=Theme.ENTRY_BG, button_color=Theme.PRIMARY, 
            button_hover_color=Theme.PRIMARY_HOVER, font=ctk.CTkFont(size=13)
        )
        self.spin_godz.grid(row=6, column=1, sticky="nsew", **pad)
        self.spin_godz.set("08")
        
        self.spin_min = ctk.CTkOptionMenu(
            self.left_frame, values=[f"{i:02d}" for i in range(60)],
            fg_color=Theme.ENTRY_BG, button_color=Theme.PRIMARY, 
            button_hover_color=Theme.PRIMARY_HOVER, font=ctk.CTkFont(size=13)
        )
        self.spin_min.grid(row=6, column=2, sticky="nsew", **pad)
        self.spin_min.set("00")

        # 6. Action Buttons - Wykaz
        self.btn_wykaz = self._add_button(7, 0, "Utwórz wykaz", self.utworz_wykaz)
        self.btn_wykaz_pdf = self._add_button(7, 1, "Konwersja PDF", self.utworz_wykaz_pdf, is_secondary=True)
        self._add_button(7, 2, "Folder Wykazy", self.otworz_folder_wykazy)

        # 7. Action Buttons - Skierowania
        self.btn_skier = self._add_button(8, 0, "Utwórz skierowania", self.utworz_skierowania)
        self.btn_skier_pdf = self._add_button(8, 1, "Konwersja PDF", self.utworz_skierowania_pdf, is_secondary=True)
        self._add_button(8, 2, "Folder Skierowania", self.otworz_folder_skierowania)

        # 8. Result Label
        self.wynik = ctk.CTkLabel(
            self.left_frame, text="Wynik", text_color="#A4C2C2", 
            anchor="w", font=ctk.CTkFont(size=12)
        )
        self.wynik.grid(row=9, column=0, columnspan=3, sticky="ew", padx=10, pady=(15, 5))

    def _add_label_entry(self, row: int, text: str, default_val: str) -> ctk.CTkEntry:
        pad = {"padx": 8, "pady": 6}
        ctk.CTkLabel(
            self.left_frame, text=text, anchor="w", font=ctk.CTkFont(size=13)
        ).grid(row=row, column=0, sticky="ew", **pad)
        
        ent = ctk.CTkEntry(
            self.left_frame, fg_color=Theme.ENTRY_BG, border_color=Theme.PRIMARY,
            text_color=Theme.TEXT, font=ctk.CTkFont(size=13)
        )
        ent.insert(0, default_val)
        ent.grid(row=row, column=1, columnspan=2, sticky="nsew", **pad)
        return ent

    def _add_button(self, row: int, col: int, text: str, cmd, is_secondary: bool = False) -> ctk.CTkButton:
        pad = {"padx": 8, "pady": 6}
        
        if is_secondary:
            fg = "transparent"
            hover = Theme.ENTRY_BG
            border = 1
        else:
            fg = Theme.PRIMARY
            hover = Theme.PRIMARY_HOVER
            border = 0
            
        btn = ctk.CTkButton(
            self.left_frame, text=text, command=cmd, 
            fg_color=fg, hover_color=hover, 
            border_width=border, border_color=Theme.PRIMARY,
            text_color=Theme.TEXT, font=ctk.CTkFont(size=13)
        )
        btn.grid(row=row, column=col, sticky="nsew", **pad)
        return btn

    def get_selected_klasa(self) -> str:
        val = self.seg_klasa.get()
        if not val: return "1"
        return val.split(" ")[1] # "Klasa 1" -> "1"

    def update_combobox_from_seg(self, value: str = None):
        self.update_combobox()

    def update_combobox(self):
        klasa = self.get_selected_klasa()
        values = self.zawody_by_klasa.get(klasa, [])
        self.combo_zawod.configure(values=values)
        if values:
            self.combo_zawod.set(values[0])
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
                self.btn_wyb_plik.configure(
                    text=os.path.basename(path), 
                    fg_color=Theme.PRIMARY, hover_color=Theme.PRIMARY_HOVER
                )
                self.zawody_by_klasa = self.data_manager.get_zawody_by_klasa()
                self.update_combobox()
            else:
                messagebox.showerror("Błąd", msg)
                self.btn_wyb_plik.configure(
                    text="Niepoprawny plik", 
                    fg_color="#A93226", hover_color="#922B21"
                )

    def wypisanie_osob(self, event=None):
        if not self.current_plik:
            self.pole_tekstowe.delete("0.0", "end")
            self.pole_tekstowe.insert("end", "Nie wybrano pliku\n")
            return
        
        filtered = self.data_manager.filter_data(self.get_selected_klasa(), self.combo_zawod.get())
        
        lines = []
        for i, r in enumerate(filtered):
            imie = r.get('Imię', '')
            drugie = r.get('Drugie imię', '')
            nazwisko = r.get('Nazwisko', '')
            full_name = f"{imie} {drugie} {nazwisko}".replace("  ", " ").strip()
            lines.append(f"{i+1}. {full_name}")
            
        tekst = "\n".join(lines)
        self.pole_tekstowe.delete("0.0", "end")
        self.pole_tekstowe.insert("end", tekst if tekst else "Brak danych")

    def symbol_zawodu(self, zawod: str) -> str:
        return self.zawody_dict.get(zawod, "N/A")

    def utworz_wykaz(self):
        filtered = self.data_manager.filter_data(self.get_selected_klasa(), self.combo_zawod.get())
        if not filtered: return

        template_path = os.path.join(self.app_dir, "Szablony", "szablon_wykaz_v3.docx")
        if not os.path.exists(template_path):
            messagebox.showerror("Błąd", f"Nie znaleziono szablonu:\n{template_path}")
            return

        try:
            doc = DocxTemplate(template_path)
            
            lines = []
            for i, r in enumerate(filtered):
                imie = r.get('Imię', '')
                nazwisko = r.get('Nazwisko', '')
                lines.append(f"{i+1}. {imie} {nazwisko}".strip())
            lista_osob = "\n".join(lines)
            
            context = {
                "dataWyst": self.data_wystawienia.get(),
                "zawod": self.combo_zawod.get(),
                "kodZawodu": self.symbol_zawodu(self.combo_zawod.get()),
                "dataRozp": self.data_rozpoczecia.get(),
                "dataZako": self.data_zakonczenia.get(),
                "godzRozp": f"{self.spin_godz.get()}:{self.spin_min.get()}",
                "stopien": self.get_selected_klasa(),
                "tabela": lista_osob,
            }
            
            doc.render(context)
            output_dir = os.path.join(self.app_dir, "Data", "Wykazy")
            os.makedirs(output_dir, exist_ok=True)
            
            save_path = os.path.join(output_dir, f"{context['stopien']}_{context['zawod']}.docx")
            doc.save(save_path)
            
            self.wynik.configure(text=f"Utworzono wykaz: {len(filtered)} osób ({datetime.datetime.now().strftime('%H:%M:%S')})")
            self.update_pdf_button_count(self.btn_wykaz_pdf, output_dir)
        except Exception as e:
            messagebox.showerror("Błąd", f"Wystąpił błąd podczas tworzenia wykazu:\n{e}")

    def utworz_skierowania(self):
        filtered = self.data_manager.filter_data(self.get_selected_klasa(), self.combo_zawod.get())
        if not filtered: return

        template_path = os.path.join(self.app_dir, "Szablony", "szablon_skierowania_v3.docx")
        if not os.path.exists(template_path):
            messagebox.showerror("Błąd", f"Nie znaleziono szablonu:\n{template_path}")
            return

        output_dir = os.path.join(self.app_dir, "Data", "Skierowania")
        os.makedirs(output_dir, exist_ok=True)

        count = 0
        try:
            for r in filtered:
                doc = DocxTemplate(template_path)
                context = {
                    "dataWyst": self.data_wystawienia.get(),
                    "imie": r.get('Imię', ''),
                    "drugie_imie": r.get('Drugie imię') or "",
                    "nazwisko": r.get('Nazwisko', ''),
                    "PESEL": r.get('PESEL', ''),
                    "zawod": self.combo_zawod.get(),
                    "kodZawodu": self.symbol_zawodu(self.combo_zawod.get()),
                    "dataRozp": self.data_rozpoczecia.get(),
                    "dataZako": self.data_zakonczenia.get(),
                    "godzRozp": f"{self.spin_godz.get()}:{self.spin_min.get()}",
                    "stopien": self.get_selected_klasa(),
                }
                doc.render(context)
                
                filename = f"{context['stopien']}_{context['zawod']}_{context['imie']}_{context['nazwisko']}.docx"
                # Cleaning filename to be OS-safe
                filename = "".join(c for c in filename if c.isalnum() or c in (" ", "_", ".", "-")).rstrip()
                
                save_path = self.get_unique_path(output_dir, filename)
                doc.save(save_path)
                count += 1

            self.wynik.configure(text=f"Utworzono: {count} dokumentów skierowań ({datetime.datetime.now().strftime('%H:%M:%S')})")
            self.update_pdf_button_count(self.btn_skier_pdf, output_dir)
        except Exception as e:
            messagebox.showerror("Błąd", f"Wystąpił błąd podczas tworzenia skierowań:\n{e}")

    def get_unique_path(self, directory: str, filename: str) -> str:
        base, ext = os.path.splitext(filename)
        path = os.path.join(directory, filename)
        counter = 1
        while os.path.exists(path):
            path = os.path.join(directory, f"{base}_{counter}{ext}")
            counter += 1
        return path

    def update_pdf_button_count(self, btn: ctk.CTkButton, directory: str):
        if os.path.exists(directory):
            count = len([f for f in os.listdir(directory) if f.endswith(".docx") and not f.startswith("~")])
            btn.configure(text=f"PDF: {count} plików")

    def utworz_pdf_batch(self, directory: str):
        if not os.path.exists(directory):
            messagebox.showerror("Błąd", "Folder nie istnieje.")
            return
            
        files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith(".docx") and not f.startswith("~")]
        if not files: 
            messagebox.showinfo("Info", "Brak plików .docx do konwersji w tym folderze.")
            return
            
        self.wynik.configure(text="Konwersja PDF w toku...")
        
        def run_conversion():
            try:
                def convert_file(file):
                    convert_with_libreoffice(file, directory)
                    
                with ThreadPoolExecutor(max_workers=4) as executor:
                    list(executor.map(convert_file, files))
                    
                self.after(0, lambda: messagebox.showinfo("Info", "Konwersja zakończona pomyślnie"))
                self.after(0, lambda: self.wynik.configure(text=f"Konwersja PDF zakończona ({datetime.datetime.now().strftime('%H:%M:%S')})."))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Błąd", f"Błąd konwersji: {e}"))
                self.after(0, lambda: self.wynik.configure(text="Błąd konwersji PDF."))

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
            f"Foldery aplikacji:\n{self.app_dir}\n\n"
            f"Dane:\n{os.path.join(self.app_dir, 'Data')}\n\n"
            f"Skierowania:\n{os.path.join(self.app_dir, 'Data', 'Skierowania')}\n\n"
            f"Wykazy:\n{os.path.join(self.app_dir, 'Data', 'Wykazy')}\n\n"
        )
        
        version = "0.41 Optimized (CustomTkinter)"
        text = (
            f"Autor: Piotr Dębowski\n"
            f"Wersja: {version}\n\n"
            + license_text + disclaimer_text + folders_text +
            "Optymalizacje:\n"
            "- Cache'owanie danych Excel (szybsze działanie)\n"
            "- ThreadPoolExecutor dla konwersji PDF\n"
            "- Poprawiona obsługa błędów i ścieżek\n"
            "- Dodatkowy error handling\n"
            "- Nowoczesny interfejs użytkownika zgodny ze stylem CTK (CustomTkinter)\n"
        )
        self.pole_tekstowe.delete("0.0", "end")
        self.pole_tekstowe.insert("end", text)

if __name__ == "__main__":
    app = App()
    app.mainloop()
