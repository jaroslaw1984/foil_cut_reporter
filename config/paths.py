from pathlib import Path

# --- podstawowa ścieżka projektu (katalog "project") ---
BASE_DIR = Path(__file__).resolve().parent.parent

# --- ścieżki do serwera
SERVER = r"sipdbprod\hydms1"
DATABASE = "hydrawlo"
VIEW_FULLNAME = "hydadm.SOP_Abfrage_Auftragsbestand_Sochacki"
SAP_SERVER = "kronos.sip.local"
SAP_DATABASE = "Raporty"

# --- ścieżka do folderu z raportami cięcia folii ---
FOIL_REPORTS_PATH = r"\\na02\groups\3.PROJEKTY\Production Counter Program\FoilReports"

# --- ścieżka do pliku JSON, który przechowuje wersję programu ---
LATEST_JSON_PATH = r"\\na02\groups\3.PROJEKTY\Production Counter Program\FoilReports\update\latest.json"

# --- ścieżka do folderu z historią raportów ---
HISTORY_PATH = r"\\na02\groups\3.PROJEKTY\Production Counter Program\FoilReports\history"

# --- ścieżka z informacjami o niestandardowych foliowych papierach ---
PAPER_FOILS = BASE_DIR / "config" / "paper_foils.json"
