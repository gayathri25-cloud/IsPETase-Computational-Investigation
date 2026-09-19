# Computational Investigation of IsPETase

Sequence, comparative, and structural analysis of a PET-degrading enzyme, using public biological databases and molecular visualization tools.

## Overview

This project investigates how the sequence and structural organization of IsPETase — the PET-degrading enzyme from *Ideonella sakaiensis* (now *Pseudideonella sakaiensis*) — relate to its plastic-degrading function. The work is organized into three connected modules:

1. **Sequence & Functional Characterization** — retrieval and physicochemical/domain analysis of the IsPETase sequence (UniProt, ExPASy, InterPro)
2. **Comparative Analysis** — BLASTp search and multiple sequence alignment against five taxonomically diverse PET hydrolases/cutinases (NCBI BLAST, EBI Clustal Omega)
3. **Structure–Function Analysis** — mapping catalytic residues onto the experimentally solved 5XJH crystal structure and measuring their spatial geometry (RCSB PDB, PyMOL)

## Key finding

The catalytic triad (Ser160–Asp206–His237) is:
- Directly identifiable from the raw sequence (G-x-S-x-G nucleophile elbow motif)
- 100% conserved across all 6 comparison sequences despite only 45–52% overall sequence identity
- Spatially clustered in the 5XJH structure at hydrogen-bonding distances (2.7–3.1 Å)

Three independent lines of evidence — sequence, cross-species conservation, and 3D structure — converge on the same three residues.

## Repository structure

```
├── report/       Full project report (Word) — Intro, Methodology, Results, Discussion, Conclusion
├── database/     Excel workbook — all data/results across all 3 modules, with live formulas
├── data/         Raw sequence data (FASTA) and the final comparison-sequence set
├── alignment/    Multiple sequence alignment output (Clustal Omega, ClustalW numbered format)
├── figures/      Standalone copies of all 4 report figures
└── scripts/      Scripts used for sequence characterization and to generate the report/database
```

## Tools used
UniProt · ExPASy ProtParam · InterPro · NCBI BLASTp · EBI Clustal Omega · RCSB PDB · PyMOL

## Notes on the scripts
`scripts/build_report.js` and `scripts/build_workbook.py` were run in a sandboxed environment and reference
absolute paths from that environment — adjust paths before rerunning locally. `build_report.js` requires the
`docx` npm package; `build_workbook.py` requires `openpyxl`.

## Author
[Your Name]
