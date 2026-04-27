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

    def refresh_machines(self):
        # Tutaj docelowo będzie fetch_available_machines() z SQL
        # Na razie zrobimy mock-up (sztuczne dane)
        mock_machines = self.db_manager.fetch_available_machines()

        # Czyszczenie starych kafelków (opcjonalnie)
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        for name in mock_machines:
            card = MachineCard(self.scrollable_frame, machine_name=name)
            card.pack(fill="x", pady=5, padx=5)
            
            # Test: Aktywujmy maszynę WLO-U001 jako przykład
            if name == "WLO-U001":
                card.update_status(has_data=True)

if __name__ == "__main__":
    app = FoilApp()
    app.mainloop()