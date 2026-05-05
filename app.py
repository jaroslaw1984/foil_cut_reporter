import customtkinter as ctk
from gui.components.machine_card import MachineCard
from database.db_manager import DBManager

class FoilApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.db_manager = DBManager()

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
            card = MachineCard(self.scrollable_frame, machine_name=name)
            card.pack(fill="x", pady=5, padx=5)
            card.update_status(has_data=True) # Zawsze zielona, bo tylko takie pokazujemy
            
    def auto_refresh(self):
        self.refresh_machines()
        # 30000 ms = 30 sekund
        self.after(30000, self.auto_refresh)

if __name__ == "__main__":
    app = FoilApp()
    app.mainloop()