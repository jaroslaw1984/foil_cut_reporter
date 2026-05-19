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

    # --- POMOCNICZA METODA DO WSKAZANIA ODPowiedniego STEROWNIKA ODBC ---
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

    # --- POMOCNICZA METODA DO OBLICZANIA POZOSTAŁYCH ILOŚCI (PCS) ---
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

    # --- POMOCNICZA METODA DO SPRAWDZENIA, CZY MASZYNA JEST OBUSTRONNA ---
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

    # --- GŁÓWNA METODA AGREGUJĄCA DANE W OPARCIU O POSNR ---
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

    # --- POMOCNICZA METODA DO SUMOWANIA SEKWENCYJNEGO DEKORÓW ---
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

    # --- POMOCNICZA METODA DO USTAWIANIA STATUSU RAPORTU W KOLEJCE ---
    def generate_word_report(self, report_data: dict, machine_name: str, output_path: str):
        doc = Document()
        
        # Pierwsza sekcja - standardowo zmniejszamy tylko górę i dół
        section = doc.sections[0]
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(1.0)
        
        self._add_page_numbering(doc)
        
        # Przygotowanie tekstu nagłówka głównego
        snap_date = report_data.get("snapshot_date", "")
        shift_info = report_data.get("shift_info", "")
        
        header_text = f"{machine_name}"
        if shift_info:
            header_text += f" - {shift_info}"
        if snap_date:
            header_text += f" - data wydruku: {snap_date}"
             
        # POMOCNICZA FUNKCJA DO RYSOWANIA GŁÓWNEGO NAGŁÓWKA NA KAŻDEJ STRONIE ZBIORCZEJ
        def draw_main_header(document):
            h = document.add_heading(header_text, 0)
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            h.paragraph_format.space_after = Pt(0) 
            for run in h.runs:
                run.font.size = Pt(14)

        # Rysujemy nagłówek na pierwszej, głównej stronie
        draw_main_header(doc)
        
        data_dict = report_data.get("data", report_data)
        
        combined = data_dict.get('combined_side', [])
        sequence = data_dict.get('production_sequence', [])
        
        # --- Sekcja GŁÓWNEJ LISTY ---
        if combined:
            doc.add_heading('1. STRONA ZEWNĘTRZNA, WEWNĘTRZNA I GÓRNA (KOMBAJN)', level=1)
            self._fill_decor_section(doc, combined)
            next_num = 2
        elif sequence:
            current_side = sequence[0].get('side_desc', 'Zewn.')
            chunk = []
            section_num = 1
            
            side_to_title = {
                'Zewn.': 'STRONA ZEWNĘTRZNA',
                'Wewn.': 'STRONA WEWNĘTRZNA',
                'Górna': 'STRONA GÓRNA'
            }
            
            for item in sequence:
                side_desc = item.get('side_desc', '')
                if side_desc == current_side:
                    chunk.append(item)
                else:
                    title = side_to_title.get(current_side, 'STRONA NIEZNANA')
                    doc.add_heading(f'{section_num}. {title}', level=1)
                    self._fill_decor_section(doc, chunk)
                    
                    section_num += 1
                    current_side = side_desc
                    chunk = [item]
                    
            if chunk:
                title = side_to_title.get(current_side, 'STRONA NIEZNANA')
                doc.add_heading(f'{section_num}. {title}', level=1)
                self._fill_decor_section(doc, chunk)
                section_num += 1
                
            next_num = section_num
        else:
            doc.add_heading('1. BRAK DANYCH', level=1)
            next_num = 2

        summary_widths = (Cm(1.5), Cm(6.5), Cm(3.5), Cm(5.5))
        def style_summary_row(row):
            for i, cell in enumerate(row.cells):
                cell.width = summary_widths[i]
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)

        # --- Sekcja FOLIA OCHRONNA ---
        protective = data_dict.get('protective', {})
        
        # ZMIANA 1: Ukrywamy całkowicie stronę, jeśli brak folii ochronnej!
        if protective:
            doc.add_page_break()
            
            # ZMIANA 2: Dopisujemy nagłówek "Maszyna X - zapotrzebowanie..."
            draw_main_header(doc)
            
            doc.add_heading(f'{next_num}. Folia ochronna (SUMA ZBIORCZA)', level=1)
            
            prot_table = doc.add_table(rows=1, cols=4)
            prot_table.style = 'Table Grid'
            hdr_cells = prot_table.rows[0].cells
            hdr_cells[0].text = 'Lp.'
            hdr_cells[1].text = 'Indeks folii'
            hdr_cells[2].text = 'Dł. [mb]'
            hdr_cells[3].text = 'Uwagi'
            style_summary_row(prot_table.rows[0])

            for idx, symbol in enumerate(sorted(protective.keys()), start=1):
                meters_sum = protective[symbol]
                row = prot_table.add_row()
                row.cells[0].text = str(idx)
                row.cells[1].text = symbol
                
                val_str = f"{meters_sum:.1f}".replace('.', ',')
                run_m = row.cells[2].paragraphs[0].add_run(val_str)
                run_m.bold = True
                run_m.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)
                
                row.cells[3].text = '' 
                style_summary_row(row)
                
            next_num += 1

        # --- Sekcja FOLIA DEKORACYJNA ---
        decor_summary = {}
        for side in ['production_sequence', 'combined_side']:
            for item in data_dict.get(side, []):
                idnrk = item['idnrk']
                decor_summary[idnrk] = decor_summary.get(idnrk, 0.0) + item['meters']
                
        if decor_summary:
            # ZMIANA 3: Tworzymy nową sekcję Worda specjalnie na folie dekoracyjne!
            # Pozwala to zredukować lewy i prawy margines niemal do zera na tej jednej stronie.
            new_section = doc.add_section()
            new_section.top_margin = Cm(0.5)
            new_section.bottom_margin = Cm(0.5)
            new_section.left_margin = Cm(1.0)
            new_section.right_margin = Cm(0.5)
            
            # Nagłówek główny na górze
            draw_main_header(doc)
            
            doc.add_heading(f'{next_num}. Folia dekoracyjna (SUMA ZBIORCZA)', level=1)
            
            # ZMIANA 3 cd.: Budujemy tabelę złożoną z 8 kolumn! (Lewa i Prawa strona w jednym)
            decor_table = doc.add_table(rows=1, cols=8)
            decor_table.style = 'Table Grid'
            hdr_cells = decor_table.rows[0].cells
            
            # Nagłówki lewej części
            hdr_cells[0].text = 'Lp.'
            hdr_cells[1].text = 'Indeks'
            hdr_cells[2].text = 'Dł.[mb]'
            hdr_cells[3].text = 'Uwagi'
            # Nagłówki prawej części
            hdr_cells[4].text = 'Lp.'
            hdr_cells[5].text = 'Indeks'
            hdr_cells[6].text = 'Dł.[mb]'
            hdr_cells[7].text = 'Uwagi'

            # Specjalne szerokości - suma daje nam równe 20 cm!
            decor_widths = (Cm(0.8), Cm(2.7), Cm(1.7), Cm(4.8), Cm(0.8), Cm(2.7), Cm(1.7), Cm(4.8))

            def style_decor_row(row):
                for i, cell in enumerate(row.cells):
                    cell.width = decor_widths[i]
                    for p in cell.paragraphs:
                        p.paragraph_format.space_before = Pt(2)
                        p.paragraph_format.space_after = Pt(2)

            style_decor_row(decor_table.rows[0])

            # Dzielimy listę folii na pół
            keys = sorted(decor_summary.keys())
            mid = (len(keys) + 1) // 2

            for i in range(mid):
                row = decor_table.add_row()

                # Lewa strona danych
                left_key = keys[i]
                row.cells[0].text = str(i + 1)
                row.cells[1].text = left_key
                val_str_l = f"{decor_summary[left_key]:.1f}".replace('.', ',')
                run_l = row.cells[2].paragraphs[0].add_run(val_str_l)
                run_l.bold = True
                run_l.font.color.rgb = RGBColor(0x00, 0x66, 0xCC)
                row.cells[3].text = ''

                # Prawa strona danych (jeśli występuje)
                if i + mid < len(keys):
                    right_key = keys[i + mid]
                    row.cells[4].text = str(i + mid + 1)
                    row.cells[5].text = right_key
                    val_str_r = f"{decor_summary[right_key]:.1f}".replace('.', ',')
                    run_r = row.cells[6].paragraphs[0].add_run(val_str_r)
                    run_r.bold = True
                    run_r.font.color.rgb = RGBColor(0x00, 0x66, 0xCC)
                    row.cells[7].text = ''

                style_decor_row(row)

        try:
            doc.save(output_path)
            return True
        except Exception as e:
            print(f"Błąd zapisu: {e}")
            return False

    # --- POMOCNICZA METODA DO DODAWANIA NUMERACJI STRON W STOPCE ---
    def _add_page_numbering(self, doc):
        """Dodaje numerację stron 'Strona X z Y' w stopce dokumentu."""
        footer = doc.sections[0].footer
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        run = p.add_run("Strona ")
        self._append_page_number_field(run, "PAGE")
        run = p.add_run(" z ")
        self._append_page_number_field(run, "NUMPAGES")

    # --- POMOCNICZA METODA DO DODAWANIA DYNAMICZNYCH PÓL NUMERACJI STRON ---
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

    # --- POMOCNICZA METODA DO WYPEŁNIANIA SEKCJI DEKORACYJNEJ ---
    def _fill_decor_section(self, doc, data_list):
        if not data_list:
            doc.add_paragraph("Brak zleceń dla tej sekcji.")
            return

        current_base_geometry = None
        table = None
        
        # Szerokości kolumn (razem ok. 17 cm, idealnie na A4)
        widths = (Cm(1.5), Cm(6.0), Cm(4.0), Cm(2.5), Cm(3.0))
        
        # --- POMOCNICZA FUNKCJA DO STYLOWANIA WIERSZY TABELI DEKORACYJNEJ ---
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