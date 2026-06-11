import pandas as pd
import os
import re
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

    # --- GŁÓWNA METODA GENERUJĄCA DOKUMENT WORDA ---
    def generate_word_report(self, report_data: dict, machine_name: str, output_path: str):
        doc = Document()
        
        section = doc.sections[0]
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(1.0)
        
        self._add_page_numbering(doc)
        
        snap_date = report_data.get("snapshot_date", "")
        shift_info = report_data.get("shift_info", "")
        
        # --- Usunięcie za pomocą REGEX z nagłówka '(zmiana 3)' na '(3)' ---
        if shift_info:
            shift_info = re.sub(r'\(zmiana\s+(\d+)\)', r'(\1)', shift_info)
        
        header_text = f"{machine_name}"
        if shift_info:
            header_text += f" - {shift_info}"
        if snap_date:
            header_text += f"                {snap_date}"
             
        def draw_main_header(document):
            h = document.add_heading(header_text, 0)
            h.alignment = WD_ALIGN_PARAGRAPH.CENTER
            h.paragraph_format.space_after = Pt(0) 
            for run in h.runs:
                run.font.size = Pt(20)
                run.bold = True
                # --- Wymuszenie 100% czerni zamiast domyślnego szarego/niebieskiego z szablonu Worda ---
                run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

        draw_main_header(doc)
        
        data_dict = report_data.get("data", report_data)
        
        combined = data_dict.get('combined_side', [])
        sequence = data_dict.get('production_sequence', [])
        
        if combined:
            self._fill_decor_section(doc, combined, side_title='ZEWNĘTRZNA, WEWNĘTRZNA (KOMBAJN)')
            next_num = 1
        elif sequence:
            current_side = sequence[0].get('side_desc', 'Zewn.')
            chunk = []
            
            side_to_title = {
                'Zewn.': 'ZEWNĘTRZNA',
                'Wewn.': 'WEWNĘTRZNA',
                'Górna': 'STRONA GÓRNA'
            }
            
            for item in sequence:
                side_desc = item.get('side_desc', '')
                if side_desc == current_side:
                    chunk.append(item)
                else:
                    title = side_to_title.get(current_side, 'STRONA NIEZNANA')
                    self._fill_decor_section(doc, chunk, side_title=title)
                    
                    current_side = side_desc
                    chunk = [item]
                    
            if chunk:
                title = side_to_title.get(current_side, 'STRONA NIEZNANA')
                self._fill_decor_section(doc, chunk, side_title=title)
                
            next_num = 1
        else:
            doc.add_heading('BRAK DANYCH', level=1)
            next_num = 1

        # --- Sekcja FOLIA OCHRONNA ---
        protective = data_dict.get('protective', {})
        
        if protective:
            doc.add_page_break()
            draw_main_header(doc)
            
            doc.add_heading(f'{next_num}. Folia ochronna (SUMA ZBIORCZA)', level=1)
            
            prot_table = doc.add_table(rows=0, cols=3)
            prot_table.style = 'Table Grid'
            
            prot_widths = (Cm(3.0), Cm(3.5), Cm(10.5)) 
            def style_prot_row(row):
                for i, cell in enumerate(row.cells):
                    cell.width = prot_widths[i]
                    for p in cell.paragraphs:
                        p.paragraph_format.space_before = Pt(5)
                        p.paragraph_format.space_after = Pt(5)

            for symbol in sorted(protective.keys()):
                meters_sum = protective[symbol]
                row = prot_table.add_row()
                
                run_sym = row.cells[0].paragraphs[0].add_run(symbol)
                run_sym.bold = True
                
                val_str = f"{meters_sum:.1f}".replace('.', ',')
                run_m = row.cells[1].paragraphs[0].add_run(val_str)
                run_m.bold = True
                
                row.cells[2].text = '' 
                style_prot_row(row)
                
            next_num += 1

        # --- Sekcja FOLIA DEKORACYJNA ---
        decor_summary = {}
        for side in ['production_sequence', 'combined_side']:
            for item in data_dict.get(side, []):
                idnrk = item['idnrk']
                decor_summary[idnrk] = decor_summary.get(idnrk, 0.0) + item['meters']
                
        if decor_summary:
            new_section = doc.add_section()
            new_section.top_margin = Cm(0.5)
            new_section.bottom_margin = Cm(0.5)
            new_section.left_margin = Cm(1.0)
            new_section.right_margin = Cm(0.5)
            
            draw_main_header(doc)
            
            doc.add_heading(f'{next_num}. Folia dekoracyjna (SUMA ZBIORCZA)', level=1)
            
            decor_table = doc.add_table(rows=0, cols=6)
            decor_table.style = 'Table Grid'
            
            decor_widths = (Cm(2.5), Cm(2.5), Cm(5.0), Cm(2.5), Cm(2.5), Cm(5.0))

            def style_decor_row(row):
                for i, cell in enumerate(row.cells):
                    cell.width = decor_widths[i]
                    for p in cell.paragraphs:
                        p.paragraph_format.space_before = Pt(5)
                        p.paragraph_format.space_after = Pt(5)

            keys = sorted(
                decor_summary.keys(), 
                key=lambda k: (self._extract_width_and_type(k)[1], self._extract_width_and_type(k)[0])
            )
            
            for i in range(0, len(keys), 2):
                row = decor_table.add_row()

                left_key = keys[i]
                run_l = row.cells[0].paragraphs[0].add_run(left_key)
                run_l.bold = True
                
                val_str_l = f"{decor_summary[left_key]:.1f}".replace('.', ',')
                run_lm = row.cells[1].paragraphs[0].add_run(val_str_l)
                run_lm.bold = True
                row.cells[2].text = ''

                if i + 1 < len(keys):
                    right_key = keys[i + 1]
                    run_r = row.cells[3].paragraphs[0].add_run(right_key)
                    run_r.bold = True
                    
                    val_str_r = f"{decor_summary[right_key]:.1f}".replace('.', ',')
                    run_rm = row.cells[4].paragraphs[0].add_run(val_str_r)
                    run_rm.bold = True
                    row.cells[5].text = ''

                style_decor_row(row)

        try:
            doc.save(output_path)
            return True
        except Exception as e:
            print(f"Błąd zapisu: {e}")
            return False

    # --- POMOCNICZA METODA DO WYPEŁNIANIA SEKCJI DEKORACYJNEJ ---
    def _fill_decor_section(self, doc, data_list, side_title=""):
        if not data_list:
            doc.add_paragraph("Brak zleceń dla tej sekcji.")
            return

        current_base_geometry = None
        table = None
        
        widths = (Cm(6.5), Cm(2.5), Cm(3.0), Cm(5.0))
        
        def style_row(row):
            for i, cell in enumerate(row.cells):
                cell.width = widths[i]
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(4)
                    p.paragraph_format.space_after = Pt(4)

        for idx, item in enumerate(data_list, start=1):
            full_article = str(item['geometry'])
            base_geometry = full_article.split('-')[0]

            if base_geometry != current_base_geometry:
                current_base_geometry = base_geometry
                
                # ZMIANA: Sklejamy string bazy geometrii z tytułem strony w jednym nagłówku
                heading_text = base_geometry
                if side_title:
                    heading_text += f" - {side_title}"
                    
                doc.add_heading(heading_text, level=2)
                
                table = doc.add_table(rows=0, cols=4)
                table.style = 'Table Grid'
                table.autofit = False

            if table is not None:
                row = table.add_row()
                row_cells = row.cells
                row_cells[0].text = full_article      
                
                run_idx = row_cells[1].paragraphs[0].add_run(str(item['idnrk']))
                run_idx.bold = True
                
                val_str = f"{item['meters']:.1f}".replace('.', ',')
                run_m = row_cells[2].paragraphs[0].add_run(val_str)
                run_m.bold = True
                
                row_cells[3].text = ''
                
                style_row(row)

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
