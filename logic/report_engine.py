import pandas as pd
import os
from docx import Document
from docx.shared import RGBColor, Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

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
                    report_data['protective'][idnrk] = report_data['protective'].get(idnrk, 0.0) + meters
                
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
        """Generuje raport z zachowaniem technicznego układu i skompresowanym podsumowaniem dekorów."""
        doc = Document()
        
        # 1. Numeracja stron w stopce
        self._add_page_numbering(doc)
        
        header = doc.add_heading(f'RAPORT CIĘCIA FOLII - {machine_name}', 0)
        header.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Sekcja 1 i 2 (Strony zewn./wewn. - szczegółowe)
        doc.add_heading('1. STRONA ZEWNĘTRZNA', level=1)
        self._fill_decor_section(doc, report_data['outer_side'])

        doc.add_heading('2. STRONA WEWNĘTRZNA', level=1)
        self._fill_decor_section(doc, report_data['inner_side'])

        # 3. Folia ochronna - SUMA ZBIORCZA (Standardowa lista przed końcem sekcji głównej)
        doc.add_heading('3. Folia ochronna (SUMA ZBIORCZA)', level=1)
        if not report_data['protective']:
            doc.add_paragraph("Brak folii ochronnych.")
        else:
            for symbol in sorted(report_data['protective'].keys()):
                meters_sum = report_data['protective'][symbol]
                p = doc.add_paragraph()
                run = p.add_run(f"{symbol}:")
                run.bold = True
                # Formatowanie do 1 miejsca po przecinku z polskim przecinkiem
                val_str = f"{meters_sum:.1f}".replace('.', ',')
                run_m = p.add_run(f" {val_str} mb")
                run_m.bold = True
                run_m.font.color.rgb = RGBColor(0xCC, 0x00, 0x00) # Czerwony dla ochrony

        # --- SEKCJA PODSUMOWANIA DEKORÓW NA NOWEJ STRONIE (Dwukolumnowa) ---
        doc.add_page_break()
        doc.add_heading('4. Folia dekoracyjna (SUMA ZBIORCZA)', level=1)
        
        # Agregacja danych ze wszystkich stron
        decor_summary = {}
        for side in ['outer_side', 'inner_side']:
            for item in report_data.get(side, []):
                idnrk = item['idnrk']
                decor_summary[idnrk] = decor_summary.get(idnrk, 0.0) + item['meters']
                
        if not decor_summary:
            doc.add_paragraph("Brak folii dekoracyjnych.")
        else:
            # Sortujemy indeksy i obliczamy punkt podziału listy na pół
            sorted_keys = sorted(decor_summary.keys())
            num_items = len(sorted_keys)
            mid = (num_items + 1) // 2 
            
            # Tworzymy tabelę pomocniczą (2 kolumny)
            summary_table = doc.add_table(rows=1, cols=2)
            summary_table.autofit = False
            summary_table.columns[0].width = Cm(8.5)
            summary_table.columns[1].width = Cm(8.5)
            
            # Lewa kolumna (pierwsza połowa listy)
            cell_left = summary_table.rows[0].cells[0]
            for i in range(mid):
                symbol = sorted_keys[i]
                meters_sum = decor_summary[symbol]
                p = cell_left.add_paragraph(style='List Bullet')
                p.paragraph_format.space_after = Pt(0)
                run = p.add_run(f"{symbol}:")
                run.bold = True
                # Formatowanie
                val_str = f"{meters_sum:.1f}".replace('.', ',')
                run_m = p.add_run(f" {val_str} mb")
                run_m.bold = True
                run_m.font.color.rgb = RGBColor(0x00, 0x66, 0xCC) # Niebieski dla dekorów
                
            # Prawa kolumna (druga połowa listy)
            cell_right = summary_table.rows[0].cells[1]
            for i in range(mid, num_items):
                symbol = sorted_keys[i]
                meters_sum = decor_summary[symbol]
                p = cell_right.add_paragraph(style='List Bullet')
                p.paragraph_format.space_after = Pt(0)
                run = p.add_run(f"{symbol}:")
                run.bold = True
                # Formatowanie
                val_str = f"{meters_sum:.1f}".replace('.', ',')
                run_m = p.add_run(f" {val_str} mb")
                run_m.bold = True
                run_m.font.color.rgb = RGBColor(0x00, 0x66, 0xCC)

        try:
            doc.save(output_path)
            return True
        except Exception as e:
            print(f"Błąd zapisu: {e}")
            return False

    def _add_page_numbering(self, doc):
        """Dodaje numerację stron 'Strona X z Y' w stopce dokumentu."""
        footer = doc.sections[0].footer
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        run = p.add_run("Strona ")
        self._append_page_number_field(run, "PAGE")
        run = p.add_run(" z ")
        self._append_page_number_field(run, "NUMPAGES")

    def _append_page_number_field(self, run, field_name):
        """Pomocnicza metoda do wstawiania pól dynamicznych XML (Strona/Suma stron)."""
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        
        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = field_name
        
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')
        
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)

    def _fill_decor_section(self, doc, data_list):
        if not data_list:
            doc.add_paragraph("Brak zleceń dla tej sekcji.")
            return

        current_base_geometry = None
        table = None
        
        # Szerokości kolumn (razem ok. 17 cm, idealnie na A4)
        widths = (Cm(1.5), Cm(6.0), Cm(4.0), Cm(2.5), Cm(3.0))
        
        def style_row(row):
            for i, cell in enumerate(row.cells):
                cell.width = widths[i]
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(4)
                    p.paragraph_format.space_after = Pt(4)

        for idx, item in enumerate(data_list, start=1):
            # 1. ROZWIĄZANIE LOGICZNE: Odcinamy kolor od bazy profilu.
            full_article = str(item['geometry'])
            base_geometry = full_article.split('-')[0]

            # 2. ROZWIĄZANIE WIZUALNE: Kiedy zmienia się baza profilu, wstawiamy nagłówek i nową tabelę
            if base_geometry != current_base_geometry:
                current_base_geometry = base_geometry
                
                # Dodajemy prawdziwy nagłówek (Heading 2) poza tabelą
                doc.add_heading(f"Geometria: {base_geometry}", level=2)
                
                # Tworzymy nową, niezależną tabelę dla tej geometrii
                table = doc.add_table(rows=1, cols=5)
                table.style = 'Table Grid'
                table.autofit = False
                
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'Lp.'
                hdr_cells[1].text = 'Artykuł'
                hdr_cells[2].text = 'Indeks folii'
                hdr_cells[3].text = 'Dł. [mb]'
                hdr_cells[4].text = 'Uwagi'
                
                style_row(table.rows[0])

            # 3. Dodajemy wiersze ze wszystkimi wariantami folii do aktualnej tabeli
            if table is not None:
                row = table.add_row()
                row_cells = row.cells
                row_cells[0].text = str(idx)
                row_cells[1].text = full_article      # Wyświetlamy cały artykuł z kolorem
                row_cells[2].text = str(item['idnrk'])
                
                # Formatowanie do 1 miejsca po przecinku z polskim przecinkiem
                val_str = f"{item['meters']:.1f}".replace('.', ',')
                row_cells[3].paragraphs[0].add_run(val_str).bold = True
                
                row_cells[4].text = ''
                
                style_row(row)