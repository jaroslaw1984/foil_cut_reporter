from pathlib import Path

# --- podstawowa ścieżka projektu (katalog "project") ---
BASE_DIR = Path(__file__).resolve().parent.parent

# --- ścieżki do serwera
SERVER = r"sipdbprod\hydms1"
DATABASE = "hydrawlo"
VIEW_FULLNAME = "hydadm.SOP_Abfrage_Auftragsbestand_Sochacki"
SAP_SERVER = "kronos.sip.local"
SAP_DATABASE = "Raporty"