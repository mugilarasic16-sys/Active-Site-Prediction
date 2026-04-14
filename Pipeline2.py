import streamlit as st
import os
import io
from Bio.PDB import PDBList, PDBParser
from docx import Document
from docx.shared import RGBColor, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from streamlit_molstar import st_molstar

# 1. Set page config MUST be the first Streamlit command
st.set_page_config(page_title="Protein Analyzer", layout="wide")

st.title("🧬 Protein Active Site Analyzer")

# 2. Sidebar Inputs
uploaded_file = st.sidebar.file_uploader("Upload PDB File", type=['pdb', 'ent'])
pdb_id = st.sidebar.text_input("OR Enter PDB ID (e.g., 4NOS)").strip().upper()

file_path = None
name_tag = "protein"

# 3. File Handling
if uploaded_file:
    file_path = "temp_input.pdb"
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    name_tag = uploaded_file.name.split('.')[0]
elif pdb_id:
    pdbl = PDBList()
    downloaded_file = pdbl.retrieve_pdb_file(pdb_id, pdir='.', file_format='pdb')
    if os.path.exists(downloaded_file):
        file_path = downloaded_file
    elif os.path.exists(f"pdb{pdb_id.lower()}.ent"):
        file_path = f"pdb{pdb_id.lower()}.ent"
    name_tag = pdb_id

# 4. Processing & UI
if file_path and os.path.exists(file_path):
    try:
        parser = PDBParser(QUIET=True)
        structure = parser.get_structure(name_tag, file_path)

        col1, col2 = st.columns([1, 1])

        with col1:
            st.subheader("3D Visualization")
            st_molstar(file_path, height=500)
            st.caption("Use the 'Selection' panel in Molstar to highlight residues.")

        # Data extraction
        res_map = {'HIS': [], 'SER': [], 'ASP': []}
        for model in structure:
            for chain in model:
                for res in chain:
                    if res.resname in res_map and res.id[0] == ' ':
                        res_map[res.resname].append(f"{res.resname}{res.id[1]}({chain.id})")

        with col2:
            st.subheader("Analysis & Download")
            
            def set_cell_shading(cell, color):
                """Helper to set background color of a cell"""
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement('w:shd')
                shd.set(qn('w:fill'), color)
                tcPr.append(shd)

            def generate_docx():
                doc = Document()
                doc.add_heading(f'Active Site Report: {name_tag.upper()}', 0)
                doc.add_paragraph("Identified catalytic residues. Blocking these positions will DECREASE activity.")
                
                # Table configuration
                table = doc.add_table(rows=1, cols=3)
                table.style = 'Table Grid'
                
                # Headers
                hdr = table.rows[0].cells
                hdr[0].text, hdr[1].text, hdr[2].text = 'HIS (Catalytic)', 'SER (Nucleophilic)', 'ASP (Stabilizer)'
                
                # Colors: Red, Sky Blue, Lime Green
                colors = {
                    'HIS': RGBColor(200, 0, 0),    
                    'SER': RGBColor(0, 150, 255),  
                    'ASP': RGBColor(50, 205, 50)   
                }
                
                table_bg = "F2F2F2" # Light Gray background
                max_rows = max(len(res_map['HIS']), len(res_map['SER']), len(res_map['ASP']), 1)
                
                for i in range(max_rows):
                    row_cells = table.add_row().cells
                    for idx, key in enumerate(['HIS', 'SER', 'ASP']):
                        cell = row_cells[idx]
                        set_cell_shading(cell, table_bg) # Apply light color to table
                        
                        if i < len(res_map[key]):
                            paragraph = cell.paragraphs[0]
                            run = paragraph.add_run(res_map[key][i])
                            run.font.color.rgb = colors[key]
                            run.bold = True
                
                bio = io.BytesIO()
                doc.save(bio)
                return bio.getvalue()

            st.download_button(
                label="💾 Download Word Report",
                data=generate_docx(),
                file_name=f"{name_tag}_Analysis.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            st.write("### Found Residues")
            st.json(res_map)

    except Exception as e:
        st.error(f"Error analyzing protein: {e}")
else:
    st.info("Waiting for protein input from the sidebar...")
