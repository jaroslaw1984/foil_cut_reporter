import customtkinter as ctk
import tkinter as tk
import sys
import os
import subprocess
import json
import threading
from config.version import PROGRAM_NAME, PROGRAM_VERSION, PROGRAM_YEAR, PROGRAM_AUTHOR, DESCRIPTION, COMPANY_MAIL, PRIVATE_MAIL
from config.paths import LATEST_JSON_PATH
from pathlib import Path
from tkinter import messagebox


class AboutPopup(ctk.CTkToplevel):
    def __init__(self, parent, discovered_version=None):
        super().__init__(parent)
        self.title("O programie")
        self.resizable(False, False)
        self.grab_set()
        self.grid_columnconfigure(0, weight=1)
        self.center_popup(parent)
        
        # --- Uruchamia budowę UI (tylko jedna funkcja budująca) ---
        self._build_ui()
        
        # --- ZMIANA: Sprawdzamy, czy okno zostało wywołane przez auto-updatera ---
        if discovered_version:
            # Wiemy już, że jest nowa wersja z pętli w tle głównego okna
            self._on_update_check_done(discovered_version, None)
        else:
            # Ręczne kliknięcie z menu bocznego - sprawdzamy plik JSON
            self._check_update_async()
        
    def center_popup(self, parent):
        try:
            parent.update_idletasks()
            self.update_idletasks()
            pw = self.winfo_width()
            ph = self.winfo_height()
            rw = parent.winfo_width()
            rh = parent.winfo_height()
            rx = parent.winfo_rootx()
            ry = parent.winfo_rooty()
            x = rx + (rw - pw) // 2
            y = ry + (rh - ph) // 2
            self.geometry(f"+{x}+{y}")
        except Exception:
            pass    
                
    def _build_ui(self):
        # --- kolumna nazwy okna --- 
        title_lbl = ctk.CTkLabel(self, 
                                 text=PROGRAM_NAME, 
                                 justify="center", 
                                 font=ctk.CTkFont(size=18, weight="bold")
                                 )
        title_lbl.grid(row=0, column=0, padx=20, pady=(18, 6), sticky="ew")
        
        # --- kolumna opisu ---
        desc_lbl = ctk.CTkLabel(self, 
                                text=DESCRIPTION.strip(), 
                                justify="center", wraplength=360, 
                                font=ctk.CTkFont(size=13))
        desc_lbl.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="ew")
        
        # --- kolumna mail ---
        mail_lbl = ctk.CTkLabel(self, 
                                text=(
                                    f"Email firmowy: {COMPANY_MAIL}\n"
                                    f"Email prywatny: {PRIVATE_MAIL}"
                                    ),
                                justify="center", 
                                wraplength=360,
                                font=ctk.CTkFont(size=13)
                                )
        
        mail_lbl.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        
        # --- kolumna wersji programu
        ver_lbl = ctk.CTkLabel(self, 
                               text=f"Wersja: {PROGRAM_VERSION}", 
                               justify="center", 
                               font=ctk.CTkFont(size=13, weight="bold")
                               )
        ver_lbl.grid(row=3, column=0, padx=20, pady=(0, 6), sticky="ew")
        
        # --- zmiana statusu 
        self.status_var = tk.StringVar(value="Sprawdzam aktualizacje…")
        status_lbl = ctk.CTkLabel(self,
                                  textvariable=self.status_var,
                                  justify="center",
                                  wraplength=360,
                                  font=ctk.CTkFont(size=12),
                                  text_color="#9aa0a6"
                                  )
        status_lbl.grid(row=4, column=0, padx=20, pady=(0, 10), sticky="ew")
        
        # --- Przycisk aktualizacji (zapisany jako self, by móc go ukrywać/pokazywać w innej funkcji) ---        
        self.update_btn = ctk.CTkButton(
            self,
            text="",
            fg_color="#1f6aa5",
            hover_color="#144a73",
            command=self._on_update_click,
            )    
            
        self.update_btn.grid(row=5, column=0, padx=20, pady=(0, 12), sticky="ew")
        self.update_btn.grid_remove()  # ukryty domyślnie

        # --- Stopka ---
        footer_lbl = ctk.CTkLabel(self,
                                  text=f"© Rok: {PROGRAM_YEAR} {PROGRAM_AUTHOR}",
                                  justify="center",
                                  font=ctk.CTkFont(size=12),
                                  text_color="#9aa0a6"
                                  )
        footer_lbl.grid(row=6, column=0, padx=20, pady=(0, 12), sticky="ew")

        # --- Przycisk OK ---
        ok_btn = ctk.CTkButton(self, text="OK", command=self.destroy)
        ok_btn.grid(row=7, column=0, padx=20, pady=(0, 16))
    
    # --- sekcja aktualizacji – ukryta domyślnie, pokażmy ją tylko jeśli jest aktualizacja ---
    def _on_update_click(self):
        app_exe = Path(sys.argv[0]).resolve()
        current_app_dir = app_exe.parent
        exe_name = app_exe.name
        print("app_exe:", app_exe)
        print("current_app_dir:", current_app_dir)
        print("exe_name:", exe_name)
        self._start_updater_and_exit(current_app_dir, exe_name)            
    
    def _start_updater_and_exit(self, current_app_dir: Path, exe_name: str) -> None:
        updater_exe = current_app_dir.parent / "Foil_cutreporter_updater.exe"
        if not updater_exe.exists():
            messagebox.showerror("Aktualizacja", f"Brak updatera:\n{updater_exe}")
            return

        pid = os.getpid()

        subprocess.Popen(
            [
                str(updater_exe),
                "--pid", str(pid),
                "--latest_json", LATEST_JSON_PATH,
                "--current_dir", str(current_app_dir),
                "--exe_name", exe_name,
            ],
            close_fds=True,
        )

        # --- zamykamy aplikację, żeby zwolnić pliki (na 100% kończymy proces) ---
        try:
            # --- najpierw zamknij okna GUI ---
            try:
                self.destroy()
            except Exception:
                pass

            try:
                self.winfo_toplevel().destroy()
            except Exception:
                pass

            # --- twarde wyjście = brak ryzyka, że PID dalej żyje i blokuje pliki ---
            os._exit(0)
        except Exception:
            os._exit(0)   

    # --- Logika Aktualizacji ---
    def _version_tuple(self, v: str) -> tuple[int, ...]:
        """Rozbija tekst '2.0.2' na krotkę liczb (2, 0, 2), co pozwala na łatwe porównywanie z operatorem '>'"""
        try:
            return tuple(int(x) for x in str(v).strip().split("."))
        except Exception:
            return (0,)

    # --- Pobiera i odczytuje plik JSON z sieci firmowej ---
    def _fetch_latest_info(self) -> dict:
        p = Path(LATEST_JSON_PATH)
        raw = p.read_text(encoding="utf-8")
        return json.loads(raw)

    # --- Uruchamia sprawdzanie w osobnym wątku (żeby okienko nie 'wisiało') ---
    def _check_update_async(self):
        def worker():
            try:
                data = self._fetch_latest_info()
                server_version = str(data.get("version", "")).strip()
                if not server_version:
                    raise ValueError("latest.json nie ma pola 'version'")
                
                # Zamiast dotykać GUI bezpośrednio, zlecamy to głównemu wątkowi okna przez .after(0, ...)
                self.after(0, lambda: self._on_update_check_done(server_version, None))
            except Exception as e:
                # ZMIANA: Zapisujemy treść błędu do zmiennej tekstowej, zanim 'e' zostanie usunięte z pamięci
                err_msg = f"{type(e).__name__}: {e}"
                self.after(0, lambda: self._on_update_check_done(None, err_msg))

        t = threading.Thread(target=worker, daemon=True)
        t.start()

    # --- Ta funkcja wywoływana jest po zakończeniu wątku i bezpiecznie aktualizuje widok ---
    def _on_update_check_done(self, server_version: str | None, error: str | None):
        # ZMIANA: Zabezpieczenie na wypadek, gdyby użytkownik zdążył zamknąć okno zanim wątek skończył pracę
        if not self.winfo_exists():
            return

        if error:
            self.status_var.set(f"Nie mogę sprawdzić aktualizacji:\n{error}")
            return

        assert server_version is not None
        if self._version_tuple(server_version) > self._version_tuple(PROGRAM_VERSION):
            self.status_var.set("Dostępna aktualizacja ✅")
            self.update_btn.configure(text=f"Pobierz nową wersję: {server_version}")
            self.update_btn.grid()  # Pokazujemy ukryty wcześniej przycisk
        else:
            self.status_var.set("Posiadasz najnowszą wersję programu.")
            self.update_btn.grid_remove()  