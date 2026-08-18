import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="Centralizator Plan Școlarizare", layout="wide")
st.title("📊 Centralizator Automat: Plan de Școlarizare")
st.write("Aplicația caută automat tabelul cu clase pe **toate paginile** fișierelor PDF. Documentele scanate sau scrise de mână vor fi semnalate separat pentru introducere manuală.")

uploaded_files = st.file_uploader(
    "Trageți fișierele PDF aici (Drag & Drop)", 
    type="pdf", 
    accept_multiple_files=True
)

if uploaded_files:
    all_data = []
    manual_review_files = [] # Listă pentru fișierele cu probleme (scanate/de mână)
    
    with st.spinner('Se analizează toate paginile... Vă rugăm așteptați.'):
        for file in uploaded_files:
            school_name = file.name.replace(".pdf", "").replace(".PDF", "")
            table_found_in_file = False
            
            try:
                with pdfplumber.open(file) as pdf:
                    # Scanăm FIECARE pagină din document
                    for page in pdf.pages:
                        tables = page.extract_tables()
                        
                        for table in tables:
                            rows_extracted = 0
                            temp_data = []
                            
                            for row in table:
                                clean_row = [str(cell).strip() if cell is not None else "" for cell in row]
                                
                                # Căutăm rândurile care încep cu un număr
                                if clean_row and clean_row[0].isdigit():
                                    while len(clean_row) < 9:
                                        clean_row.append("")
                                    clean_row = clean_row[:9]
                                    
                                    final_row = [school_name] + clean_row
                                    temp_data.append(final_row)
                                    rows_extracted += 1
                            
                            # Dacă am găsit rânduri valide în acest tabel, le adăugăm
                            if rows_extracted > 0:
                                all_data.extend(temp_data)
                                table_found_in_file = True
                
                # Dacă am terminat de citit tot PDF-ul și nu am găsit niciun tabel digital:
                if not table_found_in_file:
                    manual_review_files.append(file.name)
                    
            except Exception as e:
                manual_review_files.append(f"{file.name} (Eroare la citire)")

    # Afișăm rezultatele
    st.markdown("---")
    
    # 1. Secțiunea de succes (Datele extrase)
    if all_data:
        columns = [
            "Unitate Învățământ", "Nr.", "Structura / locația", 
            "Nivel, grupa / clasa", "Aprobat prin plan", 
            "Înscriși la dată", "Efectiv minim legal", 
            "Situația față de minim", "Ore / norme afectate", "Observații"
        ]
        
        df = pd.DataFrame(all_data, columns=columns)
        
        st.success(f"✅ Au fost extrase automat date din **{len(uploaded_files) - len(manual_review_files)}** fișiere!")
        
        st.subheader("Previzualizare Date Centralizate")
        st.dataframe(df, use_container_width=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Centralizator_Clase')
        excel_data = output.getvalue()
        
        st.download_button(
            label="📥 Descarcă Centralizatorul Excel",
            data=excel_data,
            file_name="Centralizator_Automat.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    
    # 2. Secțiunea de alertă (Atenție la fișierele scanate)
    if manual_review_files:
        st.error(f"⚠️ Atenție! Următoarele {len(manual_review_files)} fișiere nu conțin tabele digitale detectabile (probabil sunt scanate ca poze sau completate de mână). Acestea trebuie verificate manual:")
        for bad_file in manual_review_files:
            st.write(f"- 📄 **{bad_file}**")
