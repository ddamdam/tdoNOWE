import streamlit as st
import pandas as pd
import io
import re
import os
from typing import Dict, Optional, List, Tuple

# PyMuPDF — na Streamlit Cloud używaj pymupdf (import fitz bywa niedostępny)
import pymupdf as fitz

# --- Logika Ekstrakcji Danych (klasa OCR) ---

class ApartmentPlanOCR:
    def __init__(self):
        # Pełna, zaktualizowana lista kolumn
        self.columns = [
            "Nazwa pliku", "NR budynku", "NR Lokalu", "Nazwa robocza typu",
            "Etap realizacyjny", "Strony świata (okna)", "Liczba kondygnacji",
            "Poddasze Użytkowe/Nieużytkowe",
            "Główne wejscie do lokalu - kondygnacja",
            "Numer lokalu z dokumentacji budowlanej",
            "Powierzchnia z dokumentacji budowlanej dla GW (m2) (z garażami)\n*UWAGA: powierzchnia do zweryfikowania po otrzymaniu proj. budowlanego",
            "Powierzchnia z dokumentacji budowlanej dla GW (m2) + pow. schodów",
            "Powierzchnia użytkowa parter (m2)",
            "Powierzchnia użytkowa piętro (m2)",
            "Powierzchnia pod ściankami i schodów Parter (m2)",
            "Powierzchnia pod ściankami  i schodów  Piętro  (m2)",
            "Łączna powierzchnia lokalu poddasze użytkowe (m2)",
            "Powierzchnia łącznie  (m2)",
            "Powierzchnia ogródek (m2)",
            "Ilość pokoi",
            "Garderoba(0 brak, 1 jest)",
            "Kuchnia (0 brak, 1 jest)",
            "Aneks kuchenny (0 brak, 1 jest)",
            "Okno w kuchni lub aneksie (0 brak, 1 jest)",
            "Schowek (0 brak, 1 jest)",
            "Spiżarnia (0 brak, 1 jest)",
            "Pom. Gosp. (0 brak, 1 jest)",
            "Lokal połączony z garażem (0 brak, 1 jest)",
            "Ilość łazienek z prysznicem/wanną",
            "Gabinet (0 brak, 1 jest)",
            "Balkon (m2)",
            "Loggia (m2)",
            "Taras (m2)",
            "Oznaczenie miejsca postojowego na terenie/w garażu",
            "Oznaczenie dodatkowego miejsca postojowego",
            "Oznaczenie terenu do wyłącznego korzystania",
            "Wysokość parteru (m)",
            "Wysokość piętra (m)",
            "Piętro 3 (0 brak, 1 jest)",
            "Wysokość Piętra 3",
            "Wysokość piwnicy (m2)"
        ]

    def extract_text_from_pdf(self, pdf_file_obj: io.BytesIO) -> Tuple[str, int]:
        """Używa metody "blocks" z PyMuPDF dla maksymalnej dokładności ekstrakcji tekstu."""
        full_text = ""
        page_count = 0
        try:
            pdf_bytes = pdf_file_obj.getvalue()
            with fitz.open(stream=pdf_bytes, filetype="pdf") as doc:
                page_count = doc.page_count
                for i, page in enumerate(doc):
                    blocks = page.get_text("blocks", sort=True)
                    page_text = " ".join([block[4].strip().replace('\n', ' ') for block in blocks if block[4].strip()])
                    full_text += f"\n--- STRONA {i + 1} ---\n" + page_text
            return full_text, page_count
        except Exception as e:
            st.error(f"Błąd podczas odczytu pliku PDF za pomocą PyMuPDF: {e}")
            return "", 0

    def extract_height_in_cm(self, text: str) -> str:
        """Wydobywa wysokość pomieszczenia (np. H=280) z tekstu strony i zwraca w CM."""
        match = re.search(r'H\s*=\s*(\d+)', text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return ""

    def extract_info_from_filename(self, filename: str) -> Tuple[Optional[str], Optional[str]]:
        """Analizuje nazwę pliku, aby wyciągnąć numer budynku i lokalu."""
        try:
            base_name = os.path.splitext(filename)[0]
            # Normalizujemy separatory najwyższego poziomu, ale nie polegamy tylko na splitach,
            # bo kod lokalu może zawierać '-'
            normalized = base_name.replace('\u2013', '-').replace('\u2014', '-').strip()

            # 1) Wyciągnij kod lokalu w formacie np. "1A", "12B" (cyfry + litera/litery)
            #    Szukamy ostatniego wystąpienia w całej nazwie, aby obsłużyć różne warianty
            apt_matches = re.findall(r"\b(\d+\s*[A-Za-z]{1,2})\b", normalized)
            apartment_code: Optional[str] = None
            if apt_matches:
                # Ostatnie dopasowanie zazwyczaj jest właściwym kodem lokalu
                apartment_code = re.sub(r"\s+", "", apt_matches[-1])
            else:
                # Fallback: jeśli są wzorce typu "ZT-1A" albo "ZT_1A", wyciągnij część po '-' lub '_'
                # i spróbuj ponownie dopasować \d+[A-Za-z]
                tail = normalized.split('_')[-1]
                tail = tail.split('-')[-1]
                m = re.search(r"(\d+\s*[A-Za-z]{1,2})", tail)
                if m:
                    apartment_code = re.sub(r"\s+", "", m.group(1))

            # 2) Numer budynku: wyciągnięty z KODU LOKALU (cyfry na początku)
            building_num: Optional[str] = None
            if apartment_code:
                m2 = re.match(r"^(\d+)", apartment_code)
                if m2:
                    building_num = m2.group(1)

            if building_num or apartment_code:
                return building_num, apartment_code
        except Exception:
            return None, None
        return None, None

    def extract_apartment_code_from_text(self, text: str) -> Optional[str]:
        """Wyodrębnia kod lokalu Z TEKSTU, jeśli nie uda się z nazwy pliku."""
        patterns = [r"Oznaczenie lokalu[:\s]+([\w\d/.-]+)"]
        matches = re.findall(patterns[0], text, re.IGNORECASE)
        for match in matches:
            if not re.match(r'\d{2}-\d{3}', match):
                return match.strip()
        return None

    def extract_building_number_from_text(self, text: str, apartment_code: Optional[str]) -> Optional[str]:
        """Wyodrębnia numer budynku Z TEKSTU."""
        # Wyłącznie z kodu lokalu (np. 1 z 1A); nie używamy kodu pocztowego
        if apartment_code:
            match = re.search(r'^(\d+)', apartment_code)
            if match:
                return match.group(1)
        return None

    # --- POPRAWIONA FUNKCJA ---
    def _extract_value(self, text: str, patterns: List[str], default_value: str = "") -> str:
        """
        Znajduje wszystkie dopasowania i zwraca OSTATNIE z nich.
        Ulepszona wersja ignoruje teraz tekstowe prefiksy (np. "ok.") przed liczbą.
        """
        normalized_text = re.sub(r'\s+', ' ', text).replace(',', '.')
        for pattern in patterns:
            # Wzorzec dopasowuje pattern, potem ignoruje opcjonalne znaki i słowa (np. ':', 'ok.'), a na końcu wyciąga liczbę
            full_pattern = pattern + r'[:\s.a-zA-Z]*?(\d+\.?\d*)'
            matches = re.findall(full_pattern, normalized_text, re.IGNORECASE)
            if matches:
                return matches[-1].strip()
        return default_value

    def extract_surface_area(self, text: str, keywords: List[str]) -> str:
        """Generyczna funkcja do ekstrakcji powierzchni na podstawie słów kluczowych."""
        return self._extract_value(text, keywords)

    def extract_floor_areas(self, text: str) -> Dict[str, str]:
        """
        Poprawiona logika do ekstrakcji powierzchni z podziałem na strony (parter/piętro).
        """
        areas = {
            "parter_uzytkowa": "", "pietro_uzytkowa": "",
            "parter_scianki": "", "pietro_scianki": "",
            "laczna": ""
        }
        
        patterns_uzytkowa = [r"POWIERZCHNIA UŻYTKOWA"]
        patterns_scianki = [r"POWIERZCHNIA POD ŚCIANKAMI(?: I SCHODÓW)?"]
        # --- POPRAWIONY WZORZEC ---
        patterns_laczna = [r"Razem Łączna Powierzchnia Lokalu", r"Łączna Powierzchnia Lokalu"]

        sections = re.split(r'--- STRONA \d+ ---', text)

        if len(sections) > 1:
            page1_text = sections[1]
            areas["parter_uzytkowa"] = self._extract_value(page1_text, patterns_uzytkowa)
            areas["parter_scianki"] = self._extract_value(page1_text, patterns_scianki)

        if len(sections) > 2:
            page2_text = sections[2]
            areas["pietro_uzytkowa"] = self._extract_value(page2_text, patterns_uzytkowa)
            areas["pietro_scianki"] = self._extract_value(page2_text, patterns_scianki)
            
        # Szukamy łącznej powierzchni w całym tekście, używając poprawionych wzorców
        areas["laczna"] = self._extract_value(text, patterns_laczna)
        return areas


    def process_pdf(self, pdf_file_obj: io.BytesIO, filename: str) -> Tuple[Dict[str, str], str]:
        # Najpierw ekstraktuj tekst z PDF
        text, page_count = self.extract_text_from_pdf(pdf_file_obj)
        if not text:
            return {col: "" for col in self.columns}, ""

        # Priorytet: wyciągamy z TEKSTU, a dopiero potem fallback do nazwy pliku
        apartment_code = self.extract_apartment_code_from_text(text)
        building_num = self.extract_building_number_from_text(text, apartment_code)

        # Fallback do nazwy pliku, jeśli czegokolwiek brakuje
        if not apartment_code or not building_num:
            file_building_num, file_apartment_code = self.extract_info_from_filename(filename)
            if not apartment_code:
                apartment_code = file_apartment_code
            if not building_num:
                building_num = file_building_num

        floor_areas = self.extract_floor_areas(text)
        
        # Ekstrakcja powierzchni z użyciem poprawionej funkcji _extract_value
        garden_area = self.extract_surface_area(text, [r"Powierzchnia ogrodu", r"ogród"])
        balkon_area = self.extract_surface_area(text, [r"balkon"])
        loggia_area = self.extract_surface_area(text, [r"loggia"])
        taras_area = self.extract_surface_area(text, [r"taras"])

        parter_height_cm = ""
        pietro_height_cm = ""
        sections = re.split(r'--- STRONA \d+ ---', text)
        main_level = ""
        if len(sections) > 1:
            page1_text = sections[1]
            parter_height_cm = self.extract_height_in_cm(page1_text)
            # Ustal główne wejście na podstawie nagłówków/tekstów strony 1
            if re.search(r"\bparter\b", page1_text, re.IGNORECASE):
                main_level = "Parter"
            elif re.search(r"\bpi[eę]tro\b", page1_text, re.IGNORECASE):
                main_level = "Piętro"
        if len(sections) > 2:
            pietro_height_cm = self.extract_height_in_cm(sections[2])
        pietro3_height_cm = ""
        if len(sections) > 3:
            pietro3_height_cm = self.extract_height_in_cm(sections[3])
            
        parter_height_m = f"{float(parter_height_cm) / 100:.2f}" if parter_height_cm.isdigit() else ""
        pietro_height_m = f"{float(pietro_height_cm) / 100:.2f}" if pietro_height_cm.isdigit() else ""
        pietro3_height_m = f"{float(pietro3_height_cm) / 100:.2f}" if pietro3_height_cm.isdigit() else ""

        # Wykrywanie obecności słów kluczowych (0/1)
        def present(patterns: List[str]) -> str:
            for p in patterns:
                if re.search(p, text, re.IGNORECASE):
                    return "1"
            return "0"

        has_garderoba = present([r"\bgarderob\w*"])
        
        # Logika dla kuchni vs aneksu - jeśli jest "salon z aneksem" to aneks=1, kuchnia=0
        has_aneks_salon = bool(re.search(r"\bsalon\s+z\s+aneks\w*", text, re.IGNORECASE))
        if has_aneks_salon:
            has_kuchnia = "0"
            has_aneks = "1"
        else:
            # Jeśli nie ma "salon z aneksem", sprawdź czy jest sama kuchnia lub aneks
            has_aneks = present([r"\baneks\w*\s+kuchenn\w*"])
            if has_aneks == "1":
                has_kuchnia = "0"
            else:
                has_kuchnia = present([r"\bkuchni\w*"])
        
        has_schowek = present([r"\bschow\w*"])
        has_spizarnia = present([r"\bspi(?:ż|z)arni\w*"])
        has_pom_gosp = present([r"\bpom\.?\s*gosp\.?\b", r"\bpomieszczenie\s+gospodarcze\b"])
        has_gabinet = present([r"\bgabine\w*"])
        has_kominek = present([r"\bkomine\w*"])
        # Dodatkowo wykrywamy obecność balkon/loggia/taras (poza metrażem)
        has_balkon = present([r"\bbalkon\w*"])
        has_loggia = present([r"\bloggi\w*"])
        has_taras = present([r"\btaras\w*"])

        # Poddasze U/N - bazuje na liczbie stron (kondygnacji)
        # 1-2 strony = tylko parter/piętro = poddasze nieużytkowe (N)
        # 3+ strony = jest dodatkowe poddasze użytkowe = U
        attic_use = "U" if page_count >= 3 else "N"

        # Liczba łazienek - szukamy WSZYSTKICH tabel w dokumencie (może być wiele stron)
        bathrooms_count = 0
        rooms_matches = re.findall(r"L\.p\.[\s\S]{0,300}?Pomieszczenie([\s\S]*?)RAZEM", text, re.IGNORECASE)
        if rooms_matches:
            for rooms_block in rooms_matches:
                # Szukamy linii z numeracją typu "0.2", "1.4" itp. + słowo "łazienka"
                bathroom_lines = re.findall(r"\d+\.\d+\s+(?:ł|l)azienk\w*", rooms_block, re.IGNORECASE)
                bathrooms_count += len(bathroom_lines)
        else:
            # Fallback: szukamy w całym tekście z numeracją
            bathroom_lines = re.findall(r"\d+\.\d+\s+(?:ł|l)azienk\w*", text, re.IGNORECASE)
            bathrooms_count = len(bathroom_lines)

        # Ilość pokoi: salon + gabinet + sypialnie we WSZYSTKICH tabelach
        rooms_count = 0
        if rooms_matches:
            for rooms_block in rooms_matches:
                # Szukamy linii z numeracją + słowa: salon, sypialnia, gabinet, pokój
                room_lines = re.findall(r"\d+\.\d+\s+(?:salon|sypialn\w*|gabine\w*|pok[oó]j\w*)", rooms_block, re.IGNORECASE)
                rooms_count += len(room_lines)
        else:
            # Fallback
            room_lines = re.findall(r"\d+\.\d+\s+(?:salon|sypialn\w*|gabine\w*|pok[oó]j\w*)", text, re.IGNORECASE)
            rooms_count = len(room_lines)

        result = {
            "Nazwa pliku": filename, "NR budynku": building_num or "", "NR Lokalu": apartment_code or "",
            "Nazwa robocza typu": "", "Etap realizacyjny": "", "Strony świata (okna)": "",
            "Liczba kondygnacji": str(page_count) if page_count > 0 else "",
            "Poddasze Użytkowe/Nieużytkowe": attic_use,
            "Główne wejscie do lokalu - kondygnacja": main_level if main_level else ("Parter" if page_count > 0 else ""),
            "Numer lokalu z dokumentacji budowlanej": apartment_code or "",
            "Powierzchnia z dokumentacji budowlanej dla GW (m2) (z garażami)\n*UWAGA: powierzchnia do zweryfikowania po otrzymaniu proj. budowlanego": "",
            "Powierzchnia z dokumentacji budowlanej dla GW (m2) + pow. schodów": "",
            "Powierzchnia użytkowa parter (m2)": floor_areas.get("parter_uzytkowa", ""),
            "Powierzchnia użytkowa piętro (m2)": floor_areas.get("pietro_uzytkowa", ""),
            "Powierzchnia pod ściankami i schodów Parter (m2)": floor_areas.get("parter_scianki", ""),
            "Powierzchnia pod ściankami  i schodów  Piętro  (m2)": floor_areas.get("pietro_scianki", ""),
            "Łączna powierzchnia lokalu poddasze użytkowe (m2)": "",
            "Powierzchnia łącznie  (m2)": floor_areas.get("laczna", ""),
            "Powierzchnia ogródek (m2)": garden_area,
            "Ilość pokoi": str(rooms_count), "Garderoba(0 brak, 1 jest)": has_garderoba, "Kuchnia (0 brak, 1 jest)": has_kuchnia,
            "Aneks kuchenny (0 brak, 1 jest)": has_aneks, "Okno w kuchni lub aneksie (0 brak, 1 jest)": "",
            "Schowek (0 brak, 1 jest)": has_schowek, "Spiżarnia (0 brak, 1 jest)": has_spizarnia, "Pom. Gosp. (0 brak, 1 jest)": has_pom_gosp,
            "Lokal połączony z garażem (0 brak, 1 jest)": "", "Ilość łazienek z prysznicem/wanną": str(bathrooms_count),
            "Gabinet (0 brak, 1 jest)": has_gabinet,
            "Balkon (m2)": balkon_area if balkon_area else ("1" if has_balkon == "1" else "0"),
            "Loggia (m2)": loggia_area if loggia_area else ("1" if has_loggia == "1" else "0"),
            "Taras (m2)": taras_area if taras_area else ("1" if has_taras == "1" else "0"),
            "Oznaczenie miejsca postojowego na terenie/w garażu": "",
            "Oznaczenie dodatkowego miejsca postojowego": "", "Oznaczenie terenu do wyłącznego korzystania": "",
            "Wysokość parteru (m)": parter_height_m,
            "Wysokość piętra (m)": pietro_height_m,
            "Piętro 3 (0 brak, 1 jest)": "1" if page_count >= 3 else "0",
            "Wysokość Piętra 3": pietro3_height_m if page_count >= 3 else "",
            "Wysokość piwnicy (m2)": ""
        }
        return result, text

# --- Interfejs Użytkownika Streamlit ---

st.set_page_config(layout="wide")
st.title("🤖 OCR do ekstrakcji danych v4.8") # Zwiększona wersja
st.subheader("Silnik ekstrakcji: PyMuPDF (metoda 'blocks', logika 'ostatniego dopasowania')")
st.write("Wgraj jeden lub więcej plików PDF. Numery lokali i budynków będą odczytywane głównie z nazw plików.")

ocr = ApartmentPlanOCR()

uploaded_files = st.file_uploader(
    "Wybierz pliki PDF (np. 'Projekt_Budynku_5A.pdf')",
    type="pdf",
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("🚀 Przetwórz pliki"):
        all_results = []
        progress_bar = st.progress(0, text="Oczekiwanie...")

        with st.spinner("Analizuję dokumenty..."):
            for i, uploaded_file in enumerate(uploaded_files):
                progress_bar.progress((i + 1) / len(uploaded_files), text=f"Przetwarzanie: {uploaded_file.name}")

                result, full_text = ocr.process_pdf(uploaded_file, uploaded_file.name)
                all_results.append(result)

                if full_text:
                    with st.expander(f"Pokaż/Ukryj pełny tekst odczytany z '{uploaded_file.name}'"):
                        st.text_area("Surowy tekst OCR:", full_text, height=300, key=f"text_area_{i}")
                else:
                    st.warning(f"Nie udało się odczytać tekstu z pliku '{uploaded_file.name}'.")

        if all_results:
            st.markdown("---")
            st.header("📊 Zbiorcze wyniki")
            st.success(f"Zakończono! Przetworzono {len(all_results)} plików.")
            df_results = pd.DataFrame(all_results, columns=ocr.columns)

            st.dataframe(df_results)

            st.session_state['df_results'] = df_results
        else:
            st.warning("Nie udało się wydobyć danych z żadnego z plików.")

if 'df_results' in st.session_state:
    df_to_download = st.session_state['df_results']

    @st.cache_data
    def to_excel(df: pd.DataFrame):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Wyniki_OCR')
        return output.getvalue()

    excel_data = to_excel(df_to_download)

    st.download_button(
        label="📥 Pobierz wyniki jako plik Excel",
        data=excel_data,
        file_name="wyniki_ocr.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheet.sheet"
    )