import streamlit as st
import os
from Bio.PDB import PDBList, PDBParser
from docx import Document
from docx.shared import Inches, RGBColor
from st_py3dmol import show_st_py3dmol
import py3Dmol
import io

st.set_page_config(page_title="Protein Analyzer", layout="wide")

st.title("🧬 Protein Active Site Analyzer")

# Use a temporary directory for PDB downloads
TEMP_DIR = "temp_pdb"
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)

uploaded_file = st.sidebar.file_uploader("Upload PDB", type=['pdb', 'ent'])
pdb_id = st.sidebar.text_input("OR Enter PDB ID").strip().upper()

file_path = None
name_tag = "protein"

if uploaded_file:
    file_path = os.path.join(TEMP_DIR, uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    name_tag = uploaded_file.name.split('.')[0]
elif pdb_id:
    pdbl = PDBList()
    # Streamlit Cloud requires specific paths for downloaded files
    file_path = pdbl.retrieve_pdb_file(pdb_id, pdir=TEMP_DIR, file_format='pdb')
    name_tag = pdb_id

if file_path:
    # ... rest of your analysis logic ...
    st.success(f"Loaded: {name_tag}")
