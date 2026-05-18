import customtkinter as ctk
import os
import json
import shutil
import time
from datetime import date
from pathlib import Path
from gui.components.machine_card import MachineCard
from database.db_manager import DBManager
from logic.report_engine import ReportEngine
from config.paths import FOIL_REPORTS_PATH

class FoilApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.db_manager = DBManager()
        self.db = self.db_manager  
        
        # Zbiór do śledzenia wydrukowanych maszyn
        self.printed_machines = set()  
        
        # Przy starcie aplikacji czyścimy folder historii, usuwając pliki starsze niż 10 dni
        self.cleanup_history_folder()

        self.title("Drukowanie raportów - FOIL CUT REPORTER")
        self.geometry("900x600")
        
        # --- Sidebar ---
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        
        self.logo_label = ctk.CTkLabel(self.sidebar, text="Raporty dla folii\ndekoracyjnych", font=("Arial", 20, "bold"))
        self.logo_label.pack(pady=20, padx=20)

        self.btn_refresh = ctk.CTkButton(self.sidebar, text="Odśwież dane", command=self.refresh_machines)
        self.btn_refresh.pack(pady=10, padx=20)
        
        self.btn_history = ctk.CTkButton(self.sidebar, text="Historia raportów", command=self.open_history_folder, fg_color="#328354", hover_color="#2ecc71")
        self.btn_history.pack(pady=10, padx=20)

        # --- Main Area ---
        self.scrollable_frame = ctk.CTkScrollableFrame(self, label_text="Lista Maszyn Produkcyjnych")
        self.scrollable_frame.pack(side="right", fill="both", expand=True, padx=20, pady=20)

        self.refresh_machines()
        self.auto_refresh()

    def refresh_machines(self):
        print("[DEBUG] 2. Próbuję połączyć się z bazą Kronos...")
        active_machines = self.db_manager.fetch_active_machines()
        print(f"[DEBUG] 3. Sukces! Pobrane maszyny: {active_machines}")
        
        active_machines = self.db_manager.fetch_active_machines()
        
        # Czyszczenie widoku
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        if not active_machines:
            self.no_orders_label = ctk.CTkLabel(
                self.scrollable_frame, 
                text="BRAK NOWYCH ZLECEŃ\nRAPORT POJAWI SIĘ AUTOMATYCZNIE", 
                font=("Arial", 20, "bold"),
                text_color="gray"
            )
            self.no_orders_label.pack(pady=100)
            return

        # 2. Budujemy karty TYLKO dla aktywnych maszyn
        for name in active_machines:
            card = MachineCard(
                self.scrollable_frame, 
                machine_name=name
            )
            
            # Przypisujemy komendy po utworzeniu karty. 
            # Używamy sztuczki z lambda `n=name, c=card`, aby zamrozić wartości zmiennych w pętli.
            card.btn_print.configure(command=lambda n=name, c=card: self.print_machine_report(n, c))
            card.btn_delete.configure(command=lambda n=name: self.delete_machine_report(n))
            
            card.pack(fill="x", pady=5, padx=5)
            card.update_status(has_data=True)
            
            if name in self.printed_machines:
                card.mark_as_printed()

    def print_machine_report(self, machine_name, card):
        print(f"--- Uruchamiam generowanie raportu dla maszyny: {machine_name} ---")
        
        engine = ReportEngine(self.db)

        safe_machine_name = str(machine_name).replace("/", "-").replace("\\", "-")
        today_str = date.today().strftime("%Y-%m-%d")
        json_path = Path(FOIL_REPORTS_PATH) / f"{safe_machine_name}_{today_str}.json"
        
        # --- LOKALIZACJA HISTORII ---
        history_dir = Path(r"R:\Produkcja\Planowanie OKL\Production Counter Program\FoilReports\history")
        history_dir.mkdir(parents=True, exist_ok=True) # Tworzy katalog, jeśli nie istnieje
        
        if not json_path.exists():
            print(f"Błąd: Nie znaleziono pliku JSON -> {json_path}")
            return
            
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except Exception as e:
            print(f"Błąd odczytu pliku JSON: {e}")
            return
            
        final_report = payload.get("data", {})
        has_data = bool(
            final_report.get('outer_side') or 
            final_report.get('inner_side') or 
            final_report.get('protective') or 
            final_report.get('production_sequence') or 
            final_report.get('combined_side') or 
            final_report.get('top_side')
        )

        if has_data:
            # Drukujemy Worda OD RAZU do katalogu historii
            word_output_path = history_dir / f"Raport_Folie_{safe_machine_name}_{today_str}.docx"
            success = engine.generate_word_report(final_report, machine_name, str(word_output_path))
            
            if success:
                os.startfile(str(word_output_path))
                
                self.printed_machines.add(machine_name)
                card.mark_as_printed()
        else:
            print("Raport jest pusty po przeliczeniu.")

        print("--- Zakończono ---")

    def cleanup_history_folder(self):
        """Usuwa raporty Word starsze niż 10 dni z folderu historii."""
        history_dir = Path(r"R:\Produkcja\Planowanie OKL\Production Counter Program\FoilReports\history")
        
        # Jeśli katalog nie istnieje, nie mamy czego czyścić
        if not history_dir.exists():
            return
            
        print("--- Rozpoczynam czyszczenie starych raportów w historii ---")
        now = time.time()
        cutoff_seconds = 10 * 24 * 60 * 60  # 10 dni przeliczone na sekundy
        
        try:
            # Szukamy tylko plików .docx
            for file_path in history_dir.glob("*.docx"):
                if file_path.is_file():
                    # Obliczamy wiek pliku (obecny czas minus czas modyfikacji pliku)
                    file_age = now - file_path.stat().st_mtime
                    if file_age > cutoff_seconds:
                        os.remove(file_path)
                        print(f"Usunięto stary raport z historii: {file_path.name}")
        except Exception as e:
            print(f"Błąd podczas czyszczenia folderu historii: {e}")
            
    def open_history_folder(self):
        """Otwiera folder z historią raportów w Eksploratorze Windows."""
        history_dir = Path(r"R:\Produkcja\Planowanie OKL\Production Counter Program\FoilReports\history")
        
        # Jeśli z jakiegoś powodu katalog jeszcze nie istnieje (np. świeża instalacja), program go najpierw utworzy
        history_dir.mkdir(parents=True, exist_ok=True)
        
        try:
            os.startfile(str(history_dir))
            print(f"Otwarto folder historii: {history_dir}")
        except Exception as e:
            print(f"Błąd podczas otwierania folderu historii: {e}")        

    def delete_machine_report(self, machine_name):
        """Usuwa raport z bazy, czyści pamięć podręczną i trwale kasuje plik JSON."""
        print(f"--- Usuwanie raportu dla maszyny: {machine_name} ---")
        
        self.db.mark_report_done(machine_name)
        
        if machine_name in self.printed_machines:
            self.printed_machines.remove(machine_name)
            
        # --- BEZPOWROTNE USUWANIE PLIKU JSON ---
        safe_machine_name = str(machine_name).replace("/", "-").replace("\\", "-")
        today_str = date.today().strftime("%Y-%m-%d")
        json_path = Path(FOIL_REPORTS_PATH) / f"{safe_machine_name}_{today_str}.json"
        
        if json_path.exists():
            try:
                os.remove(json_path) # Używamy zwykłego os.remove
                print(f"Trwale usunięto plik JSON: {json_path.name}")
            except Exception as e:
                print(f"Błąd podczas usuwania pliku JSON: {e}")
        
        self.refresh_machines()
        
    def auto_refresh(self):
        self.refresh_machines()
        # 30000 ms = 30 sekund
        self.after(30000, self.auto_refresh)    
        
if __name__ == "__main__":
    print("[DEBUG] 1. Uruchamianie skryptu...")
    try:
        app = FoilApp()
        print("[DEBUG] 4. GUI zbudowane. Odpalam główne okno (mainloop)...")
        app.mainloop()
    except Exception as e:
        print(f"[BŁĄD KRYTYCZNY] Program wywalił się na starcie: {e}")