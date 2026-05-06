import customtkinter as ctk

class MachineCard(ctk.CTkFrame):
    # 1. Dodajemy print_command jako jawny parametr
    def __init__(self, master, machine_name, print_command=None, **kwargs):
        super().__init__(master, **kwargs)
        
        # Konfiguracja wag kolumn (dla wyśrodkowania kółka)
        self.grid_columnconfigure(0, weight=1) # Badge z nazwą
        self.grid_columnconfigure(1, weight=1) # Kółko statusu
        self.grid_columnconfigure(2, weight=0) # Przyciski
        self.grid_columnconfigure(3, weight=0)

        # Tworzymy ciemniejszą ramkę (tło dla nazwy)
        self.name_badge = ctk.CTkFrame(self, fg_color="#1A1A1A", corner_radius=8)
        self.name_badge.grid(row=0, column=0, padx=20, pady=10, sticky="w")

        # Napis wewnątrz ramki
        self.label = ctk.CTkLabel(
            self.name_badge, 
            text=machine_name, 
            font=("Arial", 16, "bold"), 
            text_color="#DCE4EE"
        )
        self.label.pack(padx=15, pady=5)

        # Kółko statusu
        self.status_indicator = ctk.CTkLabel(
            self, 
            text="●", 
            text_color="gray", 
            font=("Arial", 40)
        )
        self.status_indicator.grid(row=0, column=1, padx=20, pady=10)

        # Przycisk Podgląd
        self.btn_preview = ctk.CTkButton(self, text="Podgląd", width=100, fg_color="#3B3B3B")
        self.btn_preview.grid(row=0, column=2, padx=5, pady=10)

        # 2. Przycisk Drukuj (Domyślnie wyłączony, ale podpinamy komendę)
        self.btn_print = ctk.CTkButton(
            self, 
            text="Drukuj", 
            width=100, 
            state="disabled", 
            fg_color="gray",
            command=print_command  # <-- TUTAJ PODPINAMY AKCJĘ!
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