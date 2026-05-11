import pandas as pd
import os
from docx import Document
from docx.shared import RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

class ReportEngine:
    def __init__(self, db_manager):
        self.db = db_manager

    def load_excel_data(self, file_path: str) -> pd.DataFrame:
        """Wczytuje plik z Hydry (.csv lub .xlsx), automatycznie dopasowując ustawienia."""
        _, ext = os.path.splitext(file_path.lower())
        df = pd.DataFrame()
        
        if ext in ['.xlsx', '.xls']:
            try:
                df = pd.read_excel(file_path, header=1)
            except Exception as e:
                print(f"Error reading Excel: {e}")
                return pd.DataFrame()
        elif ext == '.csv':
            configs = [
                {'encoding': 'cp1250', 'sep': ';'},
                {'encoding': 'cp1250', 'sep': ','},
                {'encoding': 'utf-8', 'sep': ';'},
                {'encoding': 'utf-8', 'sep': ','},
            ]
            for config in configs:
                try:
                    temp_df = pd.read_csv(file_path, header=1, encoding=config['encoding'], sep=config['sep'], on_bad_lines='skip')
                    if 'Artykuł' in temp_df.columns:
                        df = temp_df
                        break
                except Exception:
                    continue
        
        if not df.empty and 'Artykuł' in df.columns:
            df = df.dropna(subset=['Artykuł'])
        return df

    def get_bom_details(self, matnr_list: list) -> pd.DataFrame:
        """Pobiera BOM z Kronosa. Rozszerzono filtry, aby łapać wszystkie typy folii."""
        if not matnr_list:
            return pd.DataFrame()

        matnr_str = "', '".join([str(m) for m in matnr_list])
        
        # Pobieramy szerszy zakres, aby uwzględnić różne typy folii ochronnych (POSNR 0050, 0060)
        sql = f"""
            SELECT MATNR, KOLOR, IDNRK, POSNR
            FROM HANA_INDEKS_BOM_LINIA
            WHERE MATNR IN ('{matnr_str}')
              AND (IDNRK LIKE 'F%' OR POSNR IN ('0050', '0060'))
            ORDER BY MATNR, POSNR
        """
        
        try:
            with self.db.raporty_engine.connect() as connection:
                df = pd.read_sql(sql, connection)
            return df
        except Exception as e:
            print(f"SQL Error (Kronos): {e}")
            return pd.DataFrame()

    def _extract_width_and_type(self, idnrk: str):
        """Dynamicznie wyciąga typ i szerokość z IDNRK."""
        idnrk = str(idnrk).strip()
        parts = idnrk.split('.')
        
        foil_prefix = parts[0]
        try:
            width = int(parts[-1])
        except (ValueError, IndexError):
            width = 0
            
        return foil_prefix, width

    def aggregate_requirements(self, excel_df: pd.DataFrame, bom_df: pd.DataFrame, matnr_col="Artykuł", meters_col="Docelowa wartość (P)") -> dict:
        """Agreguje dane w oparciu o POSNR."""
        report_data = {
            'outer_side': [],
            'inner_side': [],
            'protective': {} # {(idnrk): total_meters}
        }

        for _, row in excel_df.iterrows():
            matnr = str(row[matnr_col]).strip()
            meters = float(row[meters_col])
            
            requirements = bom_df[bom_df['MATNR'] == matnr]
            
            for _, bom_row in requirements.iterrows():
                posnr = str(bom_row['POSNR']).strip()
                idnrk = str(bom_row['IDNRK']).strip()
                _, width = self._extract_width_and_type(idnrk)
                
                # Klasyfikacja na podstawie POSNR
                if posnr in ['0050', '0060']:
                    # Sumowanie zbiorcze wszystkich folii ochronnych
                    report_data['protective'][idnrk] = report_data['protective'].get(idnrk, 0) + meters
                
                elif posnr == '0030':
                    # Strona zewnętrzna (Dekor)
                    self._add_to_sequential_list(report_data['outer_side'], idnrk, width, meters, matnr)
                
                elif posnr == '0020':
                    # Strona wewnętrzna (Dekor)
                    self._add_to_sequential_list(report_data['inner_side'], idnrk, width, meters, matnr)
                        
        return report_data

    def _add_to_sequential_list(self, target_list, idnrk, width, meters, geometry):
        """Pomocnicza metoda do sumowania sekwencyjnego dekorów."""
        if target_list and target_list[-1]['idnrk'] == idnrk and target_list[-1]['geometry'] == geometry:
            target_list[-1]['meters'] += meters
        else:
            target_list.append({
                'idnrk': idnrk,
                'width': width,
                'meters': meters,
                'geometry': geometry
            })

    def generate_word_report(self, report_data: dict, machine_name: str, output_path: str):
        """Generuje raport z zachowaniem technicznego układu."""
        doc = Document()
        
        header = doc.add_heading(f'RAPORT CIĘCIA FOLII - {machine_name}', 0)
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 1. Folia dekoracyjna - STRONA ZEWNĘTRZNA
        doc.add_heading('1. STRONA ZEWNĘTRZNA)', level=1)
        self._fill_decor_section(doc, report_data['outer_side'])

        # 2. Folia dekoracyjna - STRONA WEWNĘTRZNA
        doc.add_heading('2. STRONA WEWNĘTRZNA)', level=1)
        self._fill_decor_section(doc, report_data['inner_side'])

        # 3. Folia ochronna - SUMA ZBIORCZA
        doc.add_heading('3. Folia ochronna (SUMA ZBIORCZA)', level=1)
        if not report_data['protective']:
            doc.add_paragraph("Brak folii ochronnych.")
        else:
            for symbol in sorted(report_data['protective'].keys()):
                meters_sum = report_data['protective'][symbol]
                p = doc.add_paragraph()
                run = p.add_run(f"SUMA {symbol}:")
                run.bold = True
                run_m = p.add_run(f" {int(meters_sum)} mb")
                run_m.bold = True
                run_m.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

        try:
            doc.save(output_path)
            return True
        except Exception as e:
            print(f"Save error: {e}")
            return False

    def _fill_decor_section(self, doc, data_list):
        if not data_list:
            doc.add_paragraph("Brak zleceń dla tej sekcji.")
        else:
            # Generujemy tabelę Word z 5 kolumnami i widocznym obramowaniem
            table = doc.add_table(rows=1, cols=5)
            table.style = 'Table Grid'
            
            # Definiowanie nagłówków
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Lp.'
            hdr_cells[1].text = 'Geometria (Artykuł)'
            hdr_cells[2].text = 'Indeks folii'
            hdr_cells[3].text = 'Długość [mb]'
            hdr_cells[4].text = 'Uwagi / Dodatkowe info'
            
            for idx, item in enumerate(data_list, start=1):
                row_cells = table.add_row().cells
                row_cells[0].text = str(idx)
                row_cells[1].text = str(item['geometry'])
                row_cells[2].text = str(item['idnrk'])
                # Metry dodajemy osobnym "runem", żeby dało się je pogrubić
                row_cells[3].paragraphs[0].add_run(str(int(item['meters']))).bold = True
                row_cells[4].text = ''  # Puste miejsce do wypełnienia ręcznego