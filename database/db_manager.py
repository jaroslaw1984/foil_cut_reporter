import pandas as pd
import pyodbc
import urllib.parse
from sqlalchemy import create_engine, text
from config.paths import SERVER, DATABASE, VIEW_FULLNAME

class DBManager:
    def __init__(self):
        # Tworzymy silnik raz przy starcie klasy
        self.engine = self._get_hydra_engine()
    
    def _pick_driver(self) -> str:
        drivers = pyodbc.drivers()
        for name in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server", "SQL Server"):
            if name in drivers:
                return name
        raise RuntimeError(f"No SQL Server ODBC driver found. Available: {drivers}")
    
    def _get_hydra_engine(self):
        driver = self._pick_driver()
        conn_str = (
            f"DRIVER={{{driver}}};"
            f"SERVER={SERVER};"
            f"DATABASE={DATABASE};"
            "Trusted_Connection=yes;"
            "Encrypt=yes;"
            "TrustServerCertificate=yes;"
        )
        
        quoted_conn_str = urllib.parse.quote_plus(conn_str)
        # Tworzymy engine (SQLAlchemy zajmie się resztą)
        return create_engine(f"mssql+pyodbc:///?odbc_connect={quoted_conn_str}")
    
    def fetch_available_machines(self):
        sql = f"""
            SELECT DISTINCT masch_nr
            FROM {VIEW_FULLNAME}
            WHERE masch_nr IS NOT NULL AND LTRIM(RTRIM(masch_nr)) <> ''
            ORDER BY masch_nr
        """

        try:
            with self.engine.connect() as connection:
                df = pd.read_sql(text(sql), connection)
            
            # Twoja logika obróbki danych jest bardzo dobra
            return df["masch_nr"].astype("string").str.strip().dropna().tolist()
        except Exception as e:
            print(f"Błąd podczas pobierania maszyn: {e}")
            return [] # Zwracamy pustą listę, by GUI mogło to obsłużyć