import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

ARIAL = "Arial"
TITLE_FONT = Font(name=ARIAL, size=16, bold=True, color="1F4E5F")
SUB_FONT = Font(name=ARIAL, size=11, italic=True, color="555555")
SECTION_FONT = Font(name=ARIAL, size=12, bold=True, color="FFFFFF")
SECTION_FILL = PatternFill("solid", fgColor="1F4E5F")
HEADER_FONT = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="4472A8")
BODY_FONT = Font(name=ARIAL, size=10)
BOLD_BODY = Font(name=ARIAL, size=10, bold=True)
NOTE_FONT = Font(name=ARIAL, size=9, italic=True, color="806000")
thin = Side(style="thin", color="CCCCCC")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def section(ws, row, text, span=6):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = SECTION_FONT
    c.fill = SECTION_FILL
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 20
    return row + 1

def field(ws, row, label, value, note=None):
    ws.cell(row=row, column=1, value=label).font = BOLD_BODY
    vc = ws.cell(row=row, column=2, value=value)
    vc.font = BODY_FONT
    if note:
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=6)
        nc = ws.cell(row=row, column=4, value=note)
        nc.font = NOTE_FONT
    return row + 1

def table_header(ws, row, headers, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col+i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return row + 1

def table_row(ws, row, values, start_col=1, italic_cols=()):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col+i, value=v)
        c.font = Font(name=ARIAL, size=10, italic=(i in italic_cols))
        c.border = BORDER
        c.alignment = Alignment(vertical="center", wrap_text=True)
    return row + 1

def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

# ---------------------------------------------------------------
# SHEET 1: Overview
# ---------------------------------------------------------------
ov = wb.active
ov.title = "Overview"
set_widths(ov, [22, 30, 14, 45, 8, 8])

ov.merge_cells("A1:F1")
ov["A1"] = "Computational Investigation of IsPETase"
ov["A1"].font = TITLE_FONT
ov.merge_cells("A2:F2")
ov["A2"] = "Sequence, Comparative, and Structural Analysis"
ov["A2"].font = SUB_FONT

r = 4
ov.cell(row=r, column=1, value="Research Question:").font = BOLD_BODY
r += 1
ov.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
ov.cell(row=r, column=1, value=("How do sequence features and structural organization of IsPETase "
    "relate to molecular features associated with PET degradation?")).font = Font(name=ARIAL, size=10, italic=True)
ov.row_dimensions[r].height = 28
r += 2

r = table_header(ov, r, ["Module", "Focus", "Status", "Key Finding"])
rows = [
    ["Module 1", "Sequence & Functional Characterization", "Complete",
     "Catalytic triad Ser160-Asp206-His237 identified directly from the raw sequence (G-x-S-x-G nucleophile elbow motif)."],
    ["Module 2", "Comparative Sequence Analysis", "Complete",
     "Triad 100% conserved across all 6 comparison sequences despite only 45-52% overall identity."],
    ["Module 3", "Structure-Function Analysis", "Complete",
     "Triad independently confirmed in PDB 5XJH; catalytic distances measured at 2.7-3.1 A (hydrogen-bonding range)."],
]
for row_vals in rows:
    r = table_row(ov, r, row_vals)
r += 1

ov.cell(row=r, column=1, value="Sources used:").font = BOLD_BODY
r += 1
ov.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
ov.cell(row=r, column=1, value="UniProt - InterPro - NCBI BLASTp - EBI Clustal Omega - RCSB PDB - PyMOL").font = BODY_FONT

# ---------------------------------------------------------------
# SHEET 2: Module 1
# ---------------------------------------------------------------
m1 = wb.create_sheet("Module1_Sequence")
set_widths(m1, [26, 24, 14, 45, 8, 8])
m1.merge_cells("A1:F1")
m1["A1"] = "MODULE 1 -- Sequence & Functional Characterization"
m1["A1"].font = TITLE_FONT
r = 3

r = section(m1, r, "A. Identity")
r = field(m1, r, "UniProt Accession", "A0A0K8P6T7")
r = field(m1, r, "Entry Name", "PETH_PISS1")
r = field(m1, r, "Review Status", "Reviewed (Swiss-Prot)")
r = field(m1, r, "Gene", "ISF6_4831")
r = field(m1, r, "EC Number", "3.1.1.101")
r = field(m1, r, "Organism (2016 name)", "Ideonella sakaiensis")
r = field(m1, r, "Organism (2023 reclassification)", "Piscinibacter sakaiensis", "Liu et al. 2023")
r = field(m1, r, "Organism (2026 reclassification)", "Pseudideonella sakaiensis", "Lu & Chen 2026; databases lag each other on this")
r = field(m1, r, "PDB structures available", 57)
r = field(m1, r, "UniProt homology cluster size", "~4,000")
r += 1

r = section(m1, r, "B. Sequence Properties")
length_row = r
r = field(m1, r, "Length (aa)", 290)
mw_row = r
r = field(m1, r, "Molecular Weight (Da)", 30246.87, "Confirmed via ExPASy ProtParam - exact match to computed value")
r = field(m1, r, "Theoretical pI", 9.65, "Confirmed via ExPASy ProtParam")
r = field(m1, r, "Aliphatic Index", 65.07, "Confirmed via ExPASy ProtParam")
r = field(m1, r, "Instability Index", 39.51, "Confirmed via ExPASy ProtParam; II < 40 classifies the protein as STABLE")
r = field(m1, r, "GRAVY (hydropathicity)", -0.167, "Confirmed via ExPASy ProtParam; negative = overall hydrophilic, consistent with a soluble secreted enzyme")
r = field(m1, r, "Charged residues", "13 negative (Asp+Glu), 22 positive (Arg+Lys)", "Net positive charge is consistent with the high theoretical pI")
r += 1

r = section(m1, r, "C. Amino Acid Composition")
aa_header_row = r
r = table_header(m1, r, ["Residue", "Count", "% of Total"])
aa_counts = [("Ala (A)",38),("Cys (C)",4),("Asp (D)",9),("Glu (E)",4),("Phe (F)",10),
             ("Gly (G)",26),("His (H)",2),("Ile (I)",12),("Lys (K)",7),("Leu (L)",14),
             ("Met (M)",10),("Asn (N)",20),("Pro (P)",18),("Gln (Q)",10),("Arg (R)",15),
             ("Ser (S)",40),("Thr (T)",21),("Val (V)",17),("Trp (W)",5),("Tyr (Y)",8)]
first_data_row = r
for name, cnt in aa_counts:
    m1.cell(row=r, column=1, value=name).font = BODY_FONT
    m1.cell(row=r, column=1).border = BORDER
    cc = m1.cell(row=r, column=2, value=cnt)
    cc.font = BODY_FONT; cc.border = BORDER; cc.alignment = Alignment(horizontal="center")
    pc = m1.cell(row=r, column=3, value=f"=B{r}/$B${length_row}")
    pc.font = BODY_FONT; pc.border = BORDER; pc.number_format = "0.0%"
    r += 1
last_data_row = r - 1
m1.cell(row=r, column=1, value="Total / check").font = BOLD_BODY
m1.cell(row=r, column=1).border = BORDER
tc = m1.cell(row=r, column=2, value=f"=SUM(B{first_data_row}:B{last_data_row})")
tc.font = BOLD_BODY; tc.border = BORDER; tc.alignment = Alignment(horizontal="center")
pc = m1.cell(row=r, column=3, value=f"=SUM(C{first_data_row}:C{last_data_row})")
pc.font = BOLD_BODY; pc.border = BORDER; pc.number_format = "0.0%"
r += 2

r = section(m1, r, "D. Domain / Family Classification (InterPro)")
r = table_header(m1, r, ["InterPro ID", "Name", "Level"])
for row_vals in [
    ["IPR029058", "Alpha/Beta hydrolase fold", "Superfamily"],
    ["IPR041127", "PET hydrolase/cutinase-like", "Domain"],
    ["IPR050261", "FrsA/Cutinase/Hydrolase-like", "Family"],
]:
    r = table_row(m1, r, row_vals)
r += 1

r = section(m1, r, "E. Catalytic & Key Residues")
r = table_header(m1, r, ["Residue", "Auth Position", "Role", "Evidence"])
for row_vals in [
    ["Ser160", 160, "Nucleophile", "G-x-S-x-G sequence motif; Joo et al. 2018; confirmed in PDB 5XJH"],
    ["Asp206", 206, "Acid", "Joo et al. 2018; confirmed in PDB 5XJH"],
    ["His237", 237, "Base", "Joo et al. 2018; confirmed in PDB 5XJH"],
    ["Cys203", 203, "Structural (confirmed disulfide with Cys239)", "Unique to IsPETase among known PET hydrolases -- Ala at this position in related cutinases (Joo et al. 2018; Chek et al. 2026); matches this project's Module 2 alignment"],
    ["Trp185", 185, "Substrate binding / product release", "Wobbling side chain within flexible beta6-beta7 loop, enabled by S214/I218 (Chen et al. 2021; Fecker et al. 2018; Crnjar et al. 2023)"],
]:
    r = table_row(m1, r, row_vals)
r += 1

r = section(m1, r, "F. Signal Peptide")
r = field(m1, r, "Predicted region", "~residues 1-30")
r = field(m1, r, "Evidence", "Hydrophobic N-terminal stretch in raw sequence")
r = field(m1, r, "Status", "PENDING", "Not yet confirmed via InterPro's signal-peptide caller")

# ---------------------------------------------------------------
# SHEET 3: Module 2
# ---------------------------------------------------------------
m2 = wb.create_sheet("Module2_Comparative")
set_widths(m2, [16, 34, 24, 12, 10, 12, 14])
m2.merge_cells("A1:G1")
m2["A1"] = "MODULE 2 -- Comparative Sequence Analysis"
m2["A1"].font = TITLE_FONT
r = 3

r = section(m2, r, "A. Search Parameters", span=7)
r = field(m2, r, "Query", "A0A0K8P6T7 (IsPETase)")
r = field(m2, r, "Tool", "NCBI BLASTp")
r = field(m2, r, "Database", "UniProtKB/Swiss-Prot (reviewed)")
r += 1

r = section(m2, r, "B. Raw BLASTp Hits (top 10, self-hit excluded)", span=7)
r = table_header(m2, r, ["Accession", "Description", "Organism", "% Identity", "E-value", "Query Cover", "In Final Set?"])
raw_hits = [
    ["Q47RJ7.1", "Cutinase / PET hydrolase", "Thermobifida fusca", 0.5111, "2e-83", 0.91, "Yes"],
    ["E9LVH8.1", "Cutinase 1", "Thermobifida cellulosilytica", 0.5171, "3e-83", 0.89, "Yes"],
    ["E9LVH9.1", "Cutinase 2", "Thermobifida cellulosilytica", 0.5019, "9e-81", 0.89, "No - paralog of E9LVH8.1"],
    ["G8GER6.1", "Cutinase cut1", "Thermobifida fusca", 0.4963, "2e-80", 0.91, "No - paralog of Q47RJ7.1"],
    ["(confirm from your BLAST output)", "Bis(hydroxyethyl) terephthalate hydrolase (BHETase)", "Streptomyces sp.", 0.4778, "1e-78", 0.91, "No - optional extra"],
    ["D4Q9N1.2", "Cutinase est1", "Thermobifida alba", 0.5019, "2e-78", 0.88, "Yes"],
    ["F7IX06.1", "Cutinase est2", "Thermobifida alba", 0.5019, "2e-77", 0.88, "No - paralog of D4Q9N1.2"],
    ["P0DX29.1", "Cutinase / AML / Lipase", "Amycolatopsis sp.", 0.4755, "1e-75", 0.90, "No - optional extra"],
    ["G9BY57.1", "Leaf-branch compost cutinase (LCC)", "unidentified prokaryote", 0.4889, "4e-75", 0.92, "Yes"],
    ["P19833.1", "Lipase 1 / Triacylglycerol lipase", "Moraxella sp. TA144", 0.4524, "4e-66", 0.83, "Yes - outgroup"],
]
for row_vals in raw_hits:
    r = table_row(m2, r, row_vals)
    m2.cell(row=r-1, column=4).number_format = "0.00%"
    m2.cell(row=r-1, column=6).number_format = "0%"
r += 1

r = section(m2, r, "C. Final Comparison Set (6 sequences; redundant paralogs dropped)", span=7)
r = table_header(m2, r, ["Accession", "Organism", "Role", "% Identity to IsPETase"], )
final_set_start = r
final_set = [
    ["A0A0K8P6T7", "Piscinibacter / Pseudideonella sakaiensis", "Query / reference", 1.0],
    ["Q47RJ7.1", "Thermobifida fusca", "Comparison", 0.5111],
    ["E9LVH8.1", "Thermobifida cellulosilytica", "Comparison", 0.5171],
    ["D4Q9N1.2", "Thermobifida alba", "Comparison", 0.5019],
    ["G9BY57.1", "LCC (unidentified prokaryote)", "Comparison", 0.4889],
    ["P19833.1", "Moraxella sp. TA144", "Outgroup", 0.4524],
]
for row_vals in final_set:
    r = table_row(m2, r, row_vals)
    m2.cell(row=r-1, column=4).number_format = "0.00%"
final_set_end = r - 1
m2.cell(row=r, column=1, value="Average identity (excl. query)").font = BOLD_BODY
m2.cell(row=r, column=1).border = BORDER
ac = m2.cell(row=r, column=4, value=f"=AVERAGE(D{final_set_start+1}:D{final_set_end})")
ac.font = BOLD_BODY; ac.border = BORDER; ac.number_format = "0.00%"
r += 2

r = section(m2, r, "D. Catalytic Triad Conservation", span=7)
r = table_header(m2, r, ["Sequence", "Ser160 site", "Asp206 site", "His237 site"])
triad_start = r
triad_rows = [
    ["IsPETase (query)", "S", "D", "H"],
    ["T. fusca cutinase", "S", "D", "H"],
    ["T. cellulosilytica Cutinase 1", "S", "D", "H"],
    ["T. alba cutinase", "S", "D", "H"],
    ["LCC", "S", "D", "H"],
    ["Moraxella lipase (outgroup)", "S", "D", "H"],
]
for row_vals in triad_rows:
    r = table_row(m2, r, row_vals)
triad_end = r - 1
r += 1
m2.cell(row=r, column=1, value="Conserved (Ser / Asp / His count out of 6):").font = BOLD_BODY
sc = m2.cell(row=r, column=2, value=f'=COUNTIF(B{triad_start}:B{triad_end},"S")')
dc = m2.cell(row=r, column=3, value=f'=COUNTIF(C{triad_start}:C{triad_end},"D")')
hc = m2.cell(row=r, column=4, value=f'=COUNTIF(D{triad_start}:D{triad_end},"H")')
for cc in (sc, dc, hc):
    cc.font = BOLD_BODY
r += 2

r = section(m2, r, "E. Phylogenetic Tree Notes", span=7)
r = field(m2, r, "Guide Tree topology", "(T.fusca,T.cellulosilytica) closest pair -> T.alba -> P0DX29 -> LCC -> IsPETase -> Moraxella (outermost)")
r = field(m2, r, "Phylogenetic Tree topology", "Differs at root: IsPETase groups with Moraxella lipase; negative branch length on E9LVH8.1 (-0.014)")
r = field(m2, r, "Robust findings", "Thermobifida pair near-identical; Moraxella consistently most divergent")
r = field(m2, r, "Not robust", "Exact root-level branching order -- small dataset, no bootstrap support (stated as a limitation)")

# ---------------------------------------------------------------
# SHEET 4: Module 3
# ---------------------------------------------------------------
m3 = wb.create_sheet("Module3_Structure")
set_widths(m3, [30, 20, 22, 45, 8, 8])
m3.merge_cells("A1:F1")
m3["A1"] = "MODULE 3 -- Structure-Function Analysis"
m3["A1"].font = TITLE_FONT
r = 3

r = section(m3, r, "A. Structure Information")
r = field(m3, r, "PDB ID", "5XJH")
r = field(m3, r, "Title", "Crystal structure of PETase from Ideonella sakaiensis")
r = field(m3, r, "Resolution", "~1.5 A")
r = field(m3, r, "Chain analyzed", "A")
r += 1

r = section(m3, r, "B. Fold Confirmation")
r = field(m3, r, "Observation", "Alternating beta-strand / alpha-helix secondary structure pattern")
r = field(m3, r, "Interpretation", "Consistent with alpha/beta-hydrolase fold (matches InterPro IPR029058)")
r += 1

r = section(m3, r, "C. Disulfide Bridges")
r = table_header(m3, r, ["Bridge", "Approx. Residue Range", "Notes"])
for row_vals in [
    ["1", "~195-230", "Confirmed as the active-site-proximal Cys203-Cys239 disulfide, unique to IsPETase among characterized PET hydrolases (Joo et al. 2018; Chek et al. 2026)"],
    ["2", "~255-280", "Not yet cross-checked against literature"],
]:
    r = table_row(m3, r, row_vals)
r += 1

r = section(m3, r, "D. Catalytic Triad -- Structural Confirmation")
r = table_header(m3, r, ["Residue", "Auth Position", "Viewer Label Position", "Confirmed in 5XJH?"])
for row_vals in [
    ["Ser", 160, 148, "Yes"],
    ["Asp", 206, 194, "Yes"],
    ["His", 237, 225, "Yes"],
]:
    r = table_row(m3, r, row_vals)
r += 1

r = section(m3, r, "E. Measured Catalytic Distances (PyMOL, 5XJH)")
r = table_header(m3, r, ["Pair", "Distance (A)", "In H-bond range (2.5-3.5 A)?"])
for row_vals in [
    ["Ser160(OG) -- His237(NE2)", 2.9, "Yes"],
    ["His237(ND1) -- Asp206(OD1)", 2.7, "Yes"],
    ["His237(ND1) -- Asp206(OD2)", 3.1, "Yes"],
]:
    r = table_row(m3, r, row_vals)
r += 1

r = section(m3, r, "F. Other Structural Observations")
r = field(m3, r, "Flexibility signal", "Disorder increase near residues ~185-195 = the Trp185 'wobbling' loop (beta6-beta7)", "Confirmed: Fecker et al. 2018; Chen et al. 2021; Crnjar et al. 2023; Burgin et al. 2024")
r = field(m3, r, "Termini", "N/C-termini unmodeled in crystal", "Flagged as possible expression-construct artifact, not confirmed as native signal peptide")
r = field(m3, r, "UniProt Active Site annotation", "Independent curated annotation observed near triad region", "Exact residue-level confirmation pending")

# ---------------------------------------------------------------
# SHEET 5: References
# ---------------------------------------------------------------
rf = wb.create_sheet("References")
set_widths(rf, [4, 55, 30])
rf.merge_cells("A1:C1")
rf["A1"] = "References & Sources"
rf["A1"].font = TITLE_FONT
r = 3
r = table_header(rf, r, ["#", "Citation", "Used For"])
refs = [
    [1, "Yoshida S., Hiraga K., Takehana T. et al. (2016). Science 351(6278), 1196-1199. \"A bacterium that degrades and assimilates poly(ethylene terephthalate).\"", "Discovery background"],
    [2, "Han X., Liu W., Huang J.W. et al. (2017). Nature Communications 8, 2106. DOI: 10.1038/s41467-017-02255-z", "First IsPETase crystal structure"],
    [3, "Austin H.P., Allen M.D., Donohoe B. et al. (2018). PNAS 115(19), E4350-E4357. DOI: 10.1073/pnas.1718804115", "0.92A ultra-high-res structure"],
    [4, "Joo S., Cho I.J., Seo H. et al. (2018). Nature Communications 9, 382. DOI: 10.1038/s41467-018-02881-1", "Catalytic triad Ser160/Asp206/His237"],
    [5, "Fecker T., Galaz-Davison P., Engelberger F. et al. (2018). Biophysical Journal 114(6), 1302-1312. DOI: 10.1016/j.bpj.2018.02.005", "Active-site (Trp185) flexibility"],
    [6, "Chen C.C., Han X., Li X. et al. (2021). Nature Catalysis 4, 425-430. DOI: 10.1038/s41929-021-00616-y", "W185 wobbling loop, S214/I218"],
    [7, "Crnjar A., Grinen A., Kamerlin S.C.L. & Ramirez-Sarmiento C. (2023). ACS Org. Inorg. Au. DOI: 10.1021/acsorginorgau.2c00054", "Trp185 conformational selection"],
    [8, "Burgin T., Pollard B.C., Knott B. et al. (2024). Communications Chemistry 7, 65. DOI: 10.1038/s42004-024-01154-x", "Reaction mechanism, Trp185"],
    [9, "Chek M.F., Kawano E., Sanuki R. et al. (2026). Int. J. Biological Macromolecules 338, 149745. DOI: 10.1016/j.ijbiomac.2025.149745", "Cys203-Cys239 disulfide, unique to IsPETase"],
    [10, "Liu et al. (2023). Reclassification of Ideonella sakaiensis as Piscinibacter sakaiensis.", "Organism naming (2023)"],
    [11, "Lu H. & Chen L. (2026). Int. J. Syst. Evol. Microbiol. \"Emended description of the genus Rhizobacter and reclassification of Ideonella sakaiensis as Pseudideonella sakaiensis.\"", "Organism naming (2026)"],
    [12, "UniProt entry A0A0K8P6T7 (PETH_PISS1)", "Sequence, identity, EC number, active-site annotation"],
    [13, "InterPro entries IPR029058, IPR041127, IPR050261", "Domain / family classification"],
    [14, "RCSB PDB 5XJH", "Structure"],
]
for row_vals in refs:
    r = table_row(rf, r, row_vals)
r += 1
rf.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
rf.cell(row=r, column=1, value="Note: references 2-9 come from the project's literature-review dataset (or were independently verified) and include DOIs. Refs 1 and 10 were recalled rather than re-fetched -- spot-check before final submission.").font = NOTE_FONT

wb.save("/home/claude/IsPETase_Project_Database.xlsx")
print("Workbook written.")
