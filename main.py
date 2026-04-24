import customtkinter as ctk

class FoilApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Foil Cut Reporter")
        self.geometry("800x600")

if __name__ == "__main__":
    app = FoilApp()
    app.mainloop()