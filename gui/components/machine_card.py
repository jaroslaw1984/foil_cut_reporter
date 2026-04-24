import customtkinter as ctk

class MachineCard(ctk.CTkFrame):
    def __init__(self, master, machine_name, **kwargs):
        super().__init__(master, **kwargs)
        
        # Konfiguracja siatki kafelka
        self.grid_columnconfigure(0, weight=1) # Nazwa maszyny
        self.grid_columnconfigure(1, weight=0) # Status
        self.grid_columnconfigure(2, weight=0) # Przycisk Podgląd
        self.grid_columnconfigure(3, weight=0) # Przycisk Drukuj

        # 1. Nazwa maszyny i ew. ikona
        self.label = ctk.CTkLabel(self, text=machine_name, font=("Arial", 14, "bold"))
        self.label.grid(row=0, column=0, padx=20, pady=10, sticky="w")

        # 2. Status (kropka)
        self.status_indicator = ctk.CTkLabel(self, text="●", text_color="gray", font=("Arial", 20))
        self.status_indicator.grid(row=0, column=1, padx=10)

        # 3. Przycisk Podgląd
        self.btn_preview = ctk.CTkButton(self, text="Podgląd", width=100, fg_color="#3B3B3B")
        self.btn_preview.grid(row=0, column=2, padx=5, pady=10)

        # 4. Przycisk Drukuj (Domyślnie wyłączony)
        self.btn_print = ctk.CTkButton(
            self, 
            text="Drukuj", 
            width=100, 
            state="disabled", 
            fg_color="gray"
        )
        self.btn_print.grid(row=0, column=3, padx=10, pady=10)

    def update_status(self, has_data):
        """Metoda do odblokowywania przycisku i zmiany koloru statusu"""
        if has_data:
            self.btn_print.configure(state="normal", fg_color="#1f538d")
            self.status_indicator.configure(text_color="green")
        else:
            self.btn_print.configure(state="disabled", fg_color="gray")
            self.status_indicator.configure(text_color="gray")