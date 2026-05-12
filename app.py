import customtkinter as ctk
import os
import json
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
        active_machines = self.db_manager.fetch_active_machines()
        print(f"[DEBUG] 3. Sukces! Pobrane maszyny: {active_machines}")
        
        active_machines = self.db_manager.fetch_active_machines()
        
        # Czyszczenie widoku
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        if not active_machines:
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

        has_data = bool(final_report.get('outer_side') or final_report.get('inner_side') or final_report.get('protective'))

        if has_data:
            word_output_path = f"Raport_Folie_{safe_machine_name}_{today_str}.docx"
            success = engine.generate_word_report(final_report, machine_name, word_output_path)
            
            if success:
                os.startfile(word_output_path)
                
                # --- Zapisujemy, że maszyna została wydrukowana ---
                self.printed_machines.add(machine_name)
                card.mark_as_printed()
        else:
            print("Raport jest pusty po przeliczeniu.")

        print("--- Zakończono ---")

    def delete_machine_report(self, machine_name):
        """Nowa metoda obsługująca przycisk 'Usuń'"""
        print(f"--- Usuwanie raportu dla maszyny: {machine_name} ---")
        
        # Oznaczamy maszynę w bazie danych jako wykonaną (status 1)
        self.db.mark_report_done(machine_name)
        
        # --- NOWOŚĆ: Usuwamy z pamięci podręcznej, bo raport znika z listy ---
        if machine_name in self.printed_machines:
            self.printed_machines.remove(machine_name)
        
        # Odświeżamy widok GUI - maszyna zniknie z listy
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