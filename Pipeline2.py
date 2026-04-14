import streamlit as st
import os
import io
from Bio.PDB import PDBList, PDBParser
from docx import Document
from docx.shared import RGBColor
import py3Dmol
from st_py3dmol import show_st_py3dmol

# Set page config at the very top
st.set_page_config(page_title="Protein Analyzer", layout="wide")

st.title("🧬 Protein Active Site Analyzer")

# 1. Sidebar Inputs
uploaded_file = st.sidebar.file_uploader("Upload PDB File", type=['pdb', 'ent'])
pdb_id = st.sidebar.text_input("OR Enter PDB ID (e.g., 4NOS)").strip().upper()

file_path = None
name_tag = "protein"

# 2. File Handling
if uploaded_file:
    # Use a generic name to avoid path issues
    file_path = "temp_input.pdb"
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    name_tag = uploaded_file.name.split('.')[0]
elif pdb_id:
    pdbl = PDBList()
    # Download to current directory
    file_path = pdbl.retrieve_pdb_file(pdb_id, pdir='.', file_format='pdb')
    # Biopython might name the file 'pdb4nos.ent', we need to find it
    if not os.path.exists(file_path) and os.path.exists(f"pdb{pdb_id.lower()}.ent"):
        file_path = f"pdb{pdb_id.lower()}.ent"
    name_tag = pdb_id

# 3. Processing & UI
if file_path and os.path.exists(file_path):
    try:
        parser = PDBParser(QUIET=True)
        structure = parser.get_structure(name_tag, file_path)

        col1, col2 = st.columns([1, 1])

        with col1:
            st.subheader("3D Visualization")
            # Create the 3D view
            view = py3Dmol.view(width=500, height=500)
            with open(file_path, 'r') as f:
                view.addModel(f.read(), 'pdb')
            
            view.setStyle({'cartoon': {'color': 'spectrum'}})
            # Highlight active sites
            view.addStyle({'resn': 'HIS'}, {'stick': {'color': 'red'}})
            view.addStyle({'resn': 'SER'}, {'stick': {'color': 'blue'}})
            view.addStyle({'resn': 'ASP'}, {'stick': {'color': 'green'}})
            view.zoomTo()
            
            # Show the viewer
            show_st_py3dmol(view)
            st.caption("🔴 HIS | 🔵 SER | 🟢 ASP")

        # Data extraction
        res_map = {'HIS': [], 'SER': [], 'ASP': []}
        for model in structure:
            for chain in model:
                for res in chain:
                    if res.resname in res_map and res.id[0] == ' ':
                        res_map[res.resname].append(f"{res.resname}{res.id[1]}({chain.id})")

        with col2:
            st.subheader("Analysis & Download")
            
            # Word Report Generator
            def generate_docx():
                doc = Document()
                doc.add_heading(f'Active Site Report: {name_tag.upper()}', 0)
                doc.add_paragraph("Identified catalytic residues. Blocking these positions will DECREASE activity.")
                
                table = doc.add_table(rows=1, cols=3)
                table.style = 'Table Grid'
                hdr = table.rows[0].cells
                hdr[0].text, hdr[1].text, hdr[2].text = 'HIS (Catalytic)', 'SER (Nucleophilic)', 'ASP (Stabilizer)'
                
                colors = {'HIS': RGBColor(200, 0, 0), 'SER': RGBColor(0, 0, 200), 'ASP': RGBColor(0, 120, 0)}
                
                max_rows = max(len(res_map['HIS']), len(res_map['SER']), len(res_map['ASP']), 1)
                for i in range(max_rows):
                    row = table.add_row().cells
                    for idx, key in enumerate(['HIS', 'SER', 'ASP']):
                        if i < len(res_map[key]):
                            run = row[idx].paragraphs[0].add_run(res_map[key][i])
                            run.font.color.rgb = colors[key]
                
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
