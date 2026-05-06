import pandas as pd
import pyodbc
import urllib.parse
from sqlalchemy import create_engine, text
from config.paths import SERVER, DATABASE, SAP_SERVER, SAP_DATABASE

class DBManager:
    def __init__(self):
        # Silnik 1: Serwer Hydra (Sipdbprod)
        self.hydra_engine = self._get_engine(SERVER, DATABASE)
        
        # Silnik 2: Serwer Kronos (Raporty)
        self.raporty_engine = self._get_engine(SAP_SERVER, SAP_DATABASE)
    
    def _pick_driver(self) -> str:
        drivers = pyodbc.drivers()
        for name in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server", "SQL Server"):
            if name in drivers:
                return name
        raise RuntimeError(f"No SQL Server ODBC driver found. Available: {drivers}")
    
    def _get_engine(self, server, database):
        """Uniwersalna metoda tworząca silnik SQLAlchemy dla podanego serwera i bazy."""
        driver = self._pick_driver()
        conn_str = (
            f"DRIVER={{{driver}}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            "Trusted_Connection=yes;"
            "Encrypt=yes;"
            "TrustServerCertificate=yes;"
        )
        quoted_conn_str = urllib.parse.quote_plus(conn_str)
        return create_engine(f"mssql+pyodbc:///?odbc_connect={quoted_conn_str}")
    
    def fetch_active_machines(self):
        """Pobiera tylko te maszyny, które mają aktywne zlecenia (status = 0)."""
        # Dodane WITH (NOLOCK), aby uniknąć zawieszania aplikacji przy zablokowanej tabeli!
        sql = """
            SELECT DISTINCT machine_name 
            FROM tblPlanowanieFoilReportsQueue WITH (NOLOCK)
            WHERE status = 0 
            ORDER BY machine_name
        """
        try:
            with self.raporty_engine.connect() as connection:
                df = pd.read_sql(text(sql), connection)
            return df["machine_name"].astype("string").str.strip().dropna().tolist()
        except Exception as e:
            print(f"Błąd podczas odpytywania tblPlanowanieFoilReportsQueue: {e}")
            return []