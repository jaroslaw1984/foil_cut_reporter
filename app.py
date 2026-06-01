import customtkinter as ctk
import os
import json
import time
from datetime import datetime
from pathlib import Path
from gui.components.machine_card import MachineCard
from database.db_manager import DBManager
from logic.report_engine import ReportEngine
from config.paths import FOIL_REPORTS_PATH, HISTORY_PATH
from config.version import PROGRAM_NAME, PROGRAM_VERSION
from gui.components.popup_about import AboutPopup

class FoilApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.db_manager = DBManager()
        self.db = self.db_manager  
        
        # Zbiór do śledzenia wydrukowanych maszyn
        self.printed_files = set()  
        
        # Przy starcie aplikacji czyścimy folder historii, usuwając pliki starsze niż 10 dni
        self.cleanup_history_folder()

        self.title(f"{PROGRAM_NAME} v{PROGRAM_VERSION}")
        self.geometry("900x600")
        
        # --- Sidebar ---
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        
        self.logo_label = ctk.CTkLabel(self.sidebar, text="Raporty dla folii\ndekoracyjnych", font=("Arial", 20, "bold"))
        self.logo_label.pack(pady=20, padx=20)

        self.btn_refresh = ctk.CTkButton(self.sidebar, text="Odśwież dane", command=self.refresh_machines)
        self.btn_refresh.pack(pady=10, padx=20)
        
        self.btn_history = ctk.CTkButton(self.sidebar, text="Historia raportów", command=self.open_history_folder, fg_color="#328354", hover_color="#28aa5e")
        self.btn_history.pack(pady=10, padx=20)
        
        self.btn_about = ctk.CTkButton(self.sidebar, text="O programie", command=self.popup_about, fg_color="#b64f13", hover_color="#d36120")
        self.btn_about.pack(pady=10, padx=20)
        
        # --- Main Area ---
        self.scrollable_frame = ctk.CTkScrollableFrame(self, label_text="Lista Maszyn Produkcyjnych")
        self.scrollable_frame.pack(side="right", fill="both", expand=True, padx=20, pady=20)

        self.refresh_machines()
        self.auto_refresh()

    def refresh_machines(self):
        print("[DEBUG] 2. Próbuję połączyć się z bazą Kronos...")
        active_machines = self.db_manager.fetch_active_machines()
        print(f"[DEBUG] 3. Sukces! Pobrane maszyny z bazy: {active_machines}")
        
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        # --- ZMIANA: Pobieramy wszystkie pliki JSON z folderu raportów ---
        json_files = list(Path(FOIL_REPORTS_PATH).glob("*.json"))

        # Jeśli brak plików oraz brak maszyn w bazie, wyświetl komunikat
        if not json_files and not active_machines:
            self.no_orders_label = ctk.CTkLabel(
                self.scrollable_frame, 
                text="BRAK NOWYCH ZLECEŃ\nRAPORT POJAWI SIĘ AUTOMATYCZNIE", 
                font=("Arial", 20, "bold"),
                text_color="gray"
            )
            self.no_orders_label.pack(pady=100)
            return

        # 2. Budujemy karty na podstawie fizycznych plików JSON
        for json_path in json_files:
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    payload = json.load(f)
                machine_name = payload.get("machine", "Nieznana Maszyna")
            except Exception:
                # W razie błędu odczytu, bierzemy nazwę ze splitu nazwy pliku
                machine_name = json_path.stem.split("_")[0]

            card = MachineCard(
                self.scrollable_frame, 
                machine_name=machine_name
            )
            
            # --- ZMIANA: Przekazujemy konkretną ścieżkę (jp) pliku do funkcji ---
            card.btn_print.configure(command=lambda n=machine_name, c=card, jp=json_path: self.print_machine_report(n, c, jp))
            card.btn_delete.configure(command=lambda n=machine_name, jp=json_path: self.delete_machine_report(n, jp))
            
            card.pack(fill="x", pady=5, padx=5)
            card.update_status(has_data=True)
            
            # Oznaczamy jako wydrukowane na podstawie ścieżki pliku
            if str(json_path) in self.printed_files:
                card.mark_as_printed()

    def print_machine_report(self, machine_name, card, json_path):
        print(f"--- Uruchamiam generowanie raportu dla pliku: {json_path.name} ---")
        
        engine = ReportEngine(self.db)
        history_dir = Path(HISTORY_PATH)
        history_dir.mkdir(parents=True, exist_ok=True)
        
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
            # --- ZMIANA: Unikalna nazwa pliku Worda z sekundnikiem ---
            safe_machine_name = str(machine_name).replace("/", "-").replace("\\", "-")
            stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            word_output_path = history_dir / f"Raport_Folie_{safe_machine_name}_{stamp}.docx"
            
            success = engine.generate_word_report(final_report, machine_name, str(word_output_path))
            
            if success:
                os.startfile(str(word_output_path))
                # Zapisujemy do pamięci, że TEN KONKRETNY PLIK JSON został wydrukowany
                self.printed_files.add(str(json_path))
                card.mark_as_printed()
        else:
            print("Raport jest pusty po przeliczeniu.")

        print("--- Zakończono ---")

    def cleanup_history_folder(self):
        """Usuwa raporty Word starsze niż 10 dni z folderu historii."""
        history_dir = Path(HISTORY_PATH)
        
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
        history_dir = Path(HISTORY_PATH)
        
        # Jeśli z jakiegoś powodu katalog jeszcze nie istnieje (np. świeża instalacja), program go najpierw utworzy
        history_dir.mkdir(parents=True, exist_ok=True)
        
        try:
            os.startfile(str(history_dir))
            print(f"Otwarto folder historii: {history_dir}")
        except Exception as e:
            print(f"Błąd podczas otwierania folderu historii: {e}")        

    def delete_machine_report(self, machine_name, json_path):
        """Usuwa konkretny raport, a jeśli to był ostatni dla tej maszyny - czyści flagę w bazie."""
        print(f"--- Usuwanie raportu: {json_path.name} ---")
        
        if str(json_path) in self.printed_files:
            self.printed_files.remove(str(json_path))
            
        # 1. Usuwamy konkretny plik JSON
        if json_path.exists():
            try:
                os.remove(json_path)
                print(f"Trwale usunięto plik JSON: {json_path.name}")
            except Exception as e:
                print(f"Błąd podczas usuwania pliku JSON: {e}")
                
        # 2. Inteligentne czyszczenie bazy: Sprawdzamy czy został jakiś inny plik dla tej maszyny
        remaining_files = False
        safe_machine_name = str(machine_name).replace("/", "-").replace("\\", "-")
        for p in Path(FOIL_REPORTS_PATH).glob("*.json"):
            if safe_machine_name in p.name:
                remaining_files = True
                break
                
        # 3. Jeśli usunęliśmy ostatni raport, gasimy sygnał w Kronosie
        if not remaining_files:
            self.db.mark_report_done(machine_name)
            print(f"Oczyszczono kolejkę dla {machine_name}. Wysłano sygnał do bazy.")
        
        self.refresh_machines()
        
    def auto_refresh(self):
        self.refresh_machines()
        # 30000 ms = 30 sekund
        self.after(30000, self.auto_refresh)    
        
    def popup_about(self):
        AboutPopup(self)
        
if __name__ == "__main__":
    print("[DEBUG] 1. Uruchamianie skryptu...")
    try:
        app = FoilApp()
        print("[DEBUG] 4. GUI zbudowane. Odpalam główne okno (mainloop)...")
        app.mainloop()
    except Exception as e:
        print(f"[BŁĄD KRYTYCZNY] Program wywalił się na starcie: {e}")