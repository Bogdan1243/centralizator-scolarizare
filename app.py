import streamlit as st
import pdfplumber
import pandas as pd
import io

# Configurare pagină
st.set_page_config(page_title="Centralizator Plan Școlarizare", layout="wide")
st.title("📊 Centralizator Automat: Plan de Școlarizare")
st.write("Încărcați fișierele PDF primite de la unitățile de învățământ. Aplicația va extrage tabelul de pe **Pagina 2** și va genera un fișier Excel centralizat.")

# Zona de Drag & Drop
uploaded_files = st.file_uploader(
    "Trageți fișierele PDF aici sau dați click pentru a selecta (Drag & Drop)", 
    type="pdf", 
    accept_multiple_files=True
)

if uploaded_files:
    all_data = []
    
    with st.spinner('Se procesează fișierele... Vă rugăm așteptați.'):
        for file in uploaded_files:
            # Extragem numele școlii din denumirea fișierului
            school_name = file.name.replace(".pdf", "").replace(".PDF", "")
            
            try:
                with pdfplumber.open(file) as pdf:
                    # Verificăm dacă documentul are cel puțin 2 pagini
                    if len(pdf.pages) >= 2:
                        page = pdf.pages[1] # Indexul 1 reprezintă Pagina 2
                        table = page.extract_table()
                        
                        if table:
                            for row in table:
                                # Curățăm celulele (eliminăm textul gol sau caracterele invizibile)
                                clean_row = [str(cell).strip() if cell is not None else "" for cell in row]
                                
                                # Condiția de păstrare: primul element din rând trebuie să fie un număr (Nr. crt.)
                                if clean_row and clean_row[0].isdigit():
                                    
                                    # Ne asigurăm că rândul extras are fix 9 coloane (completăm cu spații goale dacă lipsesc)
                                    while len(clean_row) < 9:
                                        clean_row.append("")
                                    clean_row = clean_row[:9]
                                    
                                    # Adăugăm Unitatea de Învățământ ca primă coloană
                                    final_row = [school_name] + clean_row
                                    all_data.append(final_row)
            except Exception as e:
                st.error(f"Eroare la procesarea fișierului {file.name}: {e}")

    # Dacă am găsit date valide, construim tabelul final
    if all_data:
        columns = [
            "Unitate Învățământ", "Nr.", "Structura / locația", 
            "Nivel, grupa / clasa", "Aprobat prin plan", 
            "Înscriși la dată", "Efectiv minim legal", 
            "Situația față de minim", "Ore / norme afectate", "Observații"
        ]
        
        df = pd.DataFrame(all_data, columns=columns)
        
        st.success(f"✅ Au fost procesate cu succes datele din {len(uploaded_files)} fișiere!")
        
        # Afișare Preview
        st.subheader("Previzualizare Date Centralizate")
        st.dataframe(df, use_container_width=True)
        
        # Generare fișier Excel pentru descărcare
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Centralizator_Clase')
        excel_data = output.getvalue()
        
        st.markdown("---")
        st.download_button(
            label="📥 Descarcă Fișierul Excel",
            data=excel_data,
            file_name="Centralizator_Plan_Scolarizare.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Nu au putut fi identificate rânduri valide (care să înceapă cu un număr) pe pagina 2 a fișierelor încărcate.")
