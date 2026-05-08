import customtkinter as ctk
import os
import json
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

        self.title("Drukowanie raportów - FOIL CUT REPORTER")
        self.geometry("900x600")

        # --- Sidebar ---
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.pack(side="left", fill="y")
        
        self.logo_label = ctk.CTkLabel(self.sidebar, text="Raporty dla folii\ndekoracyjnych", font=("Arial", 20, "bold"))
        self.logo_label.pack(pady=20, padx=20)

        self.btn_refresh = ctk.CTkButton(self.sidebar, text="Odśwież dane", command=self.refresh_machines)
        self.btn_refresh.pack(pady=10, padx=20)

        # --- Main Area ---
        self.scrollable_frame = ctk.CTkScrollableFrame(self, label_text="Lista Maszyn Produkcyjnych")
        self.scrollable_frame.pack(side="right", fill="both", expand=True, padx=20, pady=20)

        self.refresh_machines()
        self.auto_refresh()

    def refresh_machines(self):
        print("[DEBUG] 2. Próbuję połączyć się z bazą Kronos...")
        # 1. Pobieramy tylko te maszyny, które mają COŚ do pocięcia
        active_machines = self.db_manager.fetch_active_machines()
        print(f"[DEBUG] 3. Sukces! Pobrane maszyny: {active_machines}")
        # 1. Pobieramy tylko te maszyny, które mają COŚ do pocięcia
        # Docelowo użyjemy nowej metody: self.db.fetch_machines_with_active_reports()
        active_machines = self.db_manager.fetch_active_machines()
        
        # Czyszczenie widoku
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        if not active_machines:
            # Jeśli nie ma zleceń, wyświetlamy duży, czytelny komunikat
            self.no_orders_label = ctk.CTkLabel(
                self.scrollable_frame, 
                text="BRAK NOWYCH ZLECEŃ\nOdpocznij lub sprawdź SQL :)", 
                font=("Arial", 20, "bold"),
                text_color="gray"
            )
            self.no_orders_label.pack(pady=100)
            return

        # 2. Budujemy karty TYLKO dla aktywnych maszyn
        for name in active_machines:
            # Przekazujemy command, który wywoła naszą nową metodę
            card = MachineCard(
                self.scrollable_frame, 
                machine_name=name,
                print_command=lambda n=name: self.print_machine_report(n)
            )
            card.pack(fill="x", pady=5, padx=5)
            card.update_status(has_data=True)
            
    def auto_refresh(self):
        self.refresh_machines()
        # 30000 ms = 30 sekund
        self.after(30000, self.auto_refresh)
        
    def print_machine_report(self, machine_name):
        print(f"--- Uruchamiam generowanie raportu dla maszyny: {machine_name} ---")
        
        engine = ReportEngine(self.db)

        # 1. Tworzymy bezpieczną ścieżkę do JSONa z zachowaniem tych samych zasad tworzenia pliku
        safe_machine_name = str(machine_name).replace("/", "-").replace("\\", "-")
        json_path = Path(FOIL_REPORTS_PATH) / f"{safe_machine_name}.json"
        
        if not json_path.exists():
            print(f"Błąd: Nie znaleziono pliku JSON -> {json_path}")
            return
            
        # 2. Ładujemy przygotowane już dane
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except Exception as e:
            print(f"Błąd odczytu pliku JSON: {e}")
            return
            
        final_report = payload.get("data", {})

        # Sprawdzamy, czy jakakolwiek sekcja raportu zawiera dane
        has_data = bool(final_report.get('outer_side') or final_report.get('inner_side') or final_report.get('protective'))

        if has_data:
            # Tworzymy ścieżkę do zapisu Worda (np. w tym samym folderze co skrypt)
            word_output_path = f"Raport_Folie_{safe_machine_name}.docx"
            
            # 6. Generowanie Worda!
            success = engine.generate_word_report(final_report, machine_name, word_output_path)
            
            if success:
                os.startfile(word_output_path)
                
                # 7. Odznaczamy maszynę w bazie danych jako wykonaną (status 1)
                self.db.mark_report_done(machine_name)
                
                # 8. Odświeżamy widok GUI na maszynie produkcyjnej 
                self.refresh_machines()
        else:
            print("Raport jest pusty po przeliczeniu.")

        print("--- Zakończono ---")
        
if __name__ == "__main__":
    print("[DEBUG] 1. Uruchamianie skryptu...")
    try:
        app = FoilApp()
        print("[DEBUG] 4. GUI zbudowane. Odpalam główne okno (mainloop)...")
        app.mainloop()
    except Exception as e:
        print(f"[BŁĄD KRYTYCZNY] Program wywalił się na starcie: {e}")