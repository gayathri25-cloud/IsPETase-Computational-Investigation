const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  ImageRun, PageBreak, Header, Footer, PageNumber, VerticalAlign
} = require("docx");

const PAGE_W = 12240, PAGE_H = 15840, MARGIN = 1440;
const USABLE = PAGE_W - 2*MARGIN; // 9360

const NAVY = "1F4E5F";
const ACCENT = "4472A8";
const LIGHTFILL = "EAF0F6";
const GREY = "666666";

function h1(text) {
  return new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing: { before: 320, after: 160 } });
}
function h2(text) {
  return new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing: { before: 240, after: 120 } });
}
function body(text, opts={}) {
  return new Paragraph({
    children: [new TextRun({ text, size: 21, ...opts })],
    spacing: { after: 160, line: 276 },
    alignment: AlignmentType.JUSTIFIED
  });
}
function caption(text) {
  return new Paragraph({
    children: [new TextRun({ text, italics: true, size: 19, color: GREY })],
    spacing: { after: 240 },
    alignment: AlignmentType.CENTER
  });
}
function note(text) {
  return new Paragraph({
    children: [new TextRun({ text, italics: true, size: 18, color: "806000" })],
    spacing: { after: 160 }
  });
}
function titleP(text, size, color, bold=true) {
  return new Paragraph({
    children: [new TextRun({ text, size, bold, color })],
    alignment: AlignmentType.CENTER,
    spacing: { after: 120 }
  });
}

function cell(text, {header=false, width, bold=false, italics=false} = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: header ? { type: ShadingType.CLEAR, fill: ACCENT } : { type: ShadingType.CLEAR, fill: "FFFFFF" },
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    children: [new Paragraph({
      children: [new TextRun({ text: String(text), size: 19, bold: header || bold, italics,
        color: header ? "FFFFFF" : "000000" })]
    })]
  });
}

function mkTable(headers, rows, colWidths) {
  const total = colWidths.reduce((a,b)=>a+b,0);
  const headerRow = new TableRow({
    children: headers.map((hd,i)=>cell(hd, {header:true, width: colWidths[i]})),
    tableHeader: true
  });
  const bodyRows = rows.map(r => new TableRow({
    children: r.map((v,i)=>cell(v, {width: colWidths[i]}))
  }));
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...bodyRows],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2, color: "BFBFBF" },
      bottom: { style: BorderStyle.SINGLE, size: 2, color: "BFBFBF" },
      left: { style: BorderStyle.SINGLE, size: 2, color: "BFBFBF" },
      right: { style: BorderStyle.SINGLE, size: 2, color: "BFBFBF" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: "D9D9D9" },
      insideVertical: { style: BorderStyle.SINGLE, size: 2, color: "D9D9D9" }
    }
  });
}

function figure(path, dispWidth, naturalW, naturalH, captionText) {
  const dispHeight = Math.round(dispWidth * naturalH / naturalW);
  return [
    new Paragraph({
      children: [new ImageRun({ type: "png", data: fs.readFileSync(path), transformation: { width: dispWidth, height: dispHeight } })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 60 }
    }),
    caption(captionText)
  ];
}

const children = [];

// ---------------- TITLE PAGE ----------------
children.push(new Paragraph({ text: "", spacing: { before: 1800 } }));
children.push(titleP("Computational Investigation of IsPETase", 40, NAVY));
children.push(titleP("Sequence, Comparative, and Structural Analysis of a PET-Degrading Enzyme", 24, "444444", false));
children.push(new Paragraph({ text: "", spacing: { before: 800 } }));
children.push(titleP("Prepared by: [Your Name]", 22, "000000", false));
children.push(titleP("September 2026", 22, "000000", false));
children.push(new Paragraph({ children: [new PageBreak()] }));

// ---------------- 1. INTRODUCTION ----------------
children.push(h1("1. Introduction"));

children.push(h2("1.1 Background"));
children.push(body(
"Poly(ethylene terephthalate) (PET) is one of the most widely produced synthetic polyesters, used extensively in beverage bottles and textiles, and its chemical stability allows it to persist in the environment for decades. In 2016, Yoshida et al. reported the discovery of a bacterium, subsequently named Ideonella sakaiensis, isolated from a PET-contaminated environment, that had evolved the ability to use PET as a primary carbon and energy source. Central to this ability is a secreted enzyme, PETase (IsPETase), which hydrolyzes the ester bonds of the PET polymer."
));
children.push(body(
"Since its discovery, IsPETase has become one of the most extensively studied PET-degrading enzymes, with published work covering its crystal structure, catalytic mechanism, active-site flexibility, and protein engineering to improve degradation efficiency (for example, the stability-enhanced variants ThermoPETase, DuraPETase, and FAST-PETase). This breadth of existing literature makes IsPETase well suited to a focused, evidence-grounded computational investigation."
));

children.push(h2("1.2 Rationale"));
children.push(body(
"Rather than attempting to discover a new PET-degrading enzyme, this project uses IsPETase as a case study to investigate how sequence-level features relate to structural organization and catalytic function, using existing biological data and public databases rather than new experimental work."
));

children.push(h2("1.3 Research Question"));
children.push(new Paragraph({
  children: [new TextRun({ text: "How do sequence features and structural organization of IsPETase relate to molecular features associated with PET degradation?", italics: true, size: 21 })],
  spacing: { after: 200 }
}));

children.push(h2("1.4 Scope"));
children.push(body(
"This project was conducted entirely computationally, using a laptop and publicly available online databases and software: UniProt, ExPASy, InterPro, NCBI BLAST, EBI Clustal Omega, RCSB PDB, and PyMOL. Wet-lab experimentation, molecular dynamics simulation, molecular docking, and machine-learning model training were outside the scope of this project, which was intentionally scoped to be small, reproducible, and achievable within a short timeframe."
));

// ---------------- 2. METHODOLOGY ----------------
children.push(h1("2. Methodology"));

children.push(h2("2.1 Module 1 -- Sequence and Functional Characterization"));
children.push(body(
"The IsPETase sequence was retrieved from UniProt (accession A0A0K8P6T7, entry PETH_PISS1). Basic physicochemical properties -- length, molecular weight, theoretical isoelectric point, amino acid composition, instability index, aliphatic index, and grand average of hydropathicity (GRAVY) -- were obtained using ExPASy ProtParam and cross-validated against an independent calculation performed directly on the retrieved sequence. Protein family and domain annotations were obtained from InterPro. Catalytic residues were identified from the primary literature (Joo et al., 2018) and confirmed by directly indexing the raw retrieved sequence at the reported residue positions."
));

children.push(h2("2.2 Module 2 -- Comparative Analysis"));
children.push(body(
"A protein BLAST (BLASTp) search was performed against the UniProtKB/Swiss-Prot database using the IsPETase sequence as the query. From the resulting hits, a set of six sequences (IsPETase plus five comparison sequences) was selected to represent taxonomic diversity while excluding redundant, near-identical paralogs originating from the same source organism. One distantly related lipase (Moraxella sp. TA144, ~45% identity to IsPETase and the only hit not independently annotated as a cutinase or PET hydrolase) was deliberately retained as an outgroup. Multiple sequence alignment of the six sequences was performed using EBI Clustal Omega, and both the pre-alignment guide tree and the post-alignment neighbor-joining phylogenetic tree it produces were examined."
));

children.push(h2("2.3 Module 3 -- Structure-Function Analysis"));
children.push(body(
"An experimentally determined IsPETase structure (PDB 5XJH, ~1.5 A resolution) was retrieved from the RCSB Protein Data Bank. The structure was first examined using RCSB's browser-based sequence-linked 3D viewer, then in PyMOL (open-source), which was used to generate the final structural figure and to measure interatomic distances between catalytic residues."
));

children.push(h2("2.4 A Note on Residue Numbering"));
children.push(body(
"During this project it was observed that PDB-linked structure viewers can display two different residue-numbering schemes for the same structure: a sequential 'label' number assigned by the viewer itself, starting from the first modeled residue, and the 'author' number, which is the original numbering used in the deposited structure file and in the primary literature. All residue numbers reported in this document use author numbering, consistent with the cited literature and with UniProt."
));

children.push(h2("2.5 Limitations"));
children.push(body(
"The phylogenetic tree generated in Module 2 used Clustal Omega's built-in neighbor-joining method without bootstrap resampling; branch order near the root of that tree should therefore be interpreted with caution (Section 5.2)."
));

// ---------------- 3. RESULTS ----------------
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(h1("3. Results"));

// --- 3.1 Module 1 ---
children.push(h2("3.1 Module 1 -- Sequence and Functional Profile"));

children.push(new Paragraph({ children:[new TextRun({text:"Table 1. Identity", bold:true, size:20})], spacing:{before:120, after:80}}));
children.push(mkTable(
  ["Field","Value"],
  [
    ["UniProt accession","A0A0K8P6T7 (PETH_PISS1), reviewed Swiss-Prot entry"],
    ["Gene","ISF6_4831"],
    ["EC number","3.1.1.101"],
    ["Organism (2016 name)","Ideonella sakaiensis"],
    ["Organism (2023 reclassification)","Piscinibacter sakaiensis (Liu et al., 2023)"],
    ["Organism (2026 reclassification)","Pseudideonella sakaiensis (Lu & Chen, 2026)"],
    ["Experimentally determined structures","57 (RCSB PDB, as of this project)"],
  ],
  [3200, 6160]
));
children.push(note("Note: UniProt and NCBI displayed different organism names during this project (2023 vs 2026 reclassification) -- a reminder that reference-database taxonomy can lag active nomenclature revisions."));

children.push(new Paragraph({ children:[new TextRun({text:"Table 2. Sequence properties", bold:true, size:20})], spacing:{before:200, after:80}}));
children.push(mkTable(
  ["Property","Value"],
  [
    ["Length","290 amino acids"],
    ["Molecular weight","30,246.87 Da"],
    ["Theoretical pI","9.65"],
    ["Instability index","39.51 (classified stable; threshold 40)"],
    ["Aliphatic index","65.07"],
    ["GRAVY","-0.167 (net hydrophilic, consistent with a secreted enzyme)"],
    ["Charged residues","22 positive (Arg+Lys), 13 negative (Asp+Glu)"],
  ],
  [3200, 6160]
));
children.push(note("All values obtained from ExPASy ProtParam; the molecular weight and pI were independently cross-checked by direct calculation from the retrieved sequence and matched exactly."));

children.push(...figure("/home/claude/fig_aa_composition.png", 500, 1600, 840,
  "Figure 1. Amino acid composition of IsPETase. Catalytic triad residues (Ser, Asp, His) highlighted in red."));

children.push(new Paragraph({ children:[new TextRun({text:"Table 3. Domain and family classification (InterPro)", bold:true, size:20})], spacing:{before:200, after:80}}));
children.push(mkTable(
  ["InterPro ID","Name","Classification level"],
  [
    ["IPR029058","Alpha/Beta hydrolase fold","Homologous superfamily"],
    ["IPR041127","PET hydrolase/cutinase-like","Domain"],
    ["IPR050261","FrsA/Cutinase/Hydrolase-like","Family"],
  ],
  [2200, 4200, 2960]
));

children.push(new Paragraph({ children:[new TextRun({text:"Table 4. Catalytic and key residues", bold:true, size:20})], spacing:{before:200, after:80}}));
children.push(mkTable(
  ["Residue","Role","Evidence"],
  [
    ["Ser160","Nucleophile","G-x-S-x-G sequence motif; Joo et al. (2018); confirmed in PDB 5XJH"],
    ["Asp206","Acid","Joo et al. (2018); confirmed in PDB 5XJH"],
    ["His237","Base","Joo et al. (2018); confirmed in PDB 5XJH"],
    ["Cys203","Forms active-site-proximal disulfide with Cys239","Confirmed unique to IsPETase among known PET hydrolases -- Ala at this position in related cutinases (Joo et al., 2018; Chek et al., 2026)"],
    ["Trp185","Substrate binding / product release","Adopts a 'wobbling' multi-conformation state within a flexible beta6-beta7 loop, enabled by IsPETase-specific residues S214 and I218 (Chen et al., 2021; Fecker et al., 2018; Crnjar et al., 2023)"],
  ],
  [1600, 2600, 5160]
));

// --- 3.2 Module 2 ---
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(h2("3.2 Module 2 -- Comparative Sequence Analysis"));

children.push(...figure("/home/claude/fig_blast_table.png", 500, 1557, 666,
  "Figure 2. BLASTp results for IsPETase against UniProtKB/Swiss-Prot (top 10 hits, self-hit excluded from analysis)."));

children.push(new Paragraph({ children:[new TextRun({text:"Table 5. Final comparison set", bold:true, size:20})], spacing:{before:200, after:80}}));
children.push(mkTable(
  ["Accession","Organism","Role","% identity to IsPETase"],
  [
    ["A0A0K8P6T7","Piscinibacter / Pseudideonella sakaiensis","Query","100%"],
    ["Q47RJ7.1","Thermobifida fusca","Comparison","51.11%"],
    ["E9LVH8.1","Thermobifida cellulosilytica","Comparison","51.71%"],
    ["D4Q9N1.2","Thermobifida alba","Comparison","50.19%"],
    ["G9BY57.1","LCC (leaf-branch compost cutinase; unidentified prokaryote)","Comparison","48.89%"],
    ["P19833.1","Moraxella sp. TA144","Outgroup","45.24%"],
  ],
  [1800, 3560, 1800, 2200]
));
children.push(note("Average identity across the five comparison sequences (excluding the query): 49.43%. Redundant paralogs from the same source organism (e.g. a second Thermobifida fusca or T. alba cutinase) were excluded as non-independent data points."));

children.push(new Paragraph({ children:[new TextRun({text:"Table 6. Catalytic triad conservation across the comparison set", bold:true, size:20})], spacing:{before:200, after:80}}));
children.push(mkTable(
  ["Sequence","Ser160 site","Asp206 site","His237 site"],
  [
    ["IsPETase (query)","Ser","Asp","His"],
    ["T. fusca cutinase","Ser","Asp","His"],
    ["T. cellulosilytica Cutinase 1","Ser","Asp","His"],
    ["T. alba cutinase","Ser","Asp","His"],
    ["LCC","Ser","Asp","His"],
    ["Moraxella lipase (outgroup)","Ser","Asp","His"],
  ],
  [3560, 1933, 1933, 1934]
));
children.push(body(
"All three catalytic-triad positions are identical (Ser / Asp / His) across all six sequences -- a 100% conservation rate -- despite overall pairwise sequence identity to IsPETase ranging from only 45.24% to 51.71% (Table 5). This holds even in the outgroup lipase, which is not independently annotated as a PET hydrolase or cutinase."
));

children.push(...figure("/home/claude/fig_phylo_tree.png", 340, 832, 620,
  "Figure 3. Post-alignment phylogenetic tree of the comparison set (EBI Clustal Omega)."));
children.push(body(
"The Thermobifida fusca and T. cellulosilytica sequences form the most closely related pair in both the guide tree and the post-alignment phylogenetic tree, and the Moraxella lipase is consistently the most divergent member. The two trees disagree, however, on the branching position of IsPETase relative to the LCC and Moraxella sequences near the root, and the post-alignment tree contains a negative branch length (T. cellulosilytica: -0.014) -- a recognized artifact of neighbor-joining on small, closely related datasets rather than evidence of a modeling error. Root-level branching order is therefore not treated as a reliable finding of this analysis (see Section 2.5)."
));

// --- 3.3 Module 3 ---
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(h2("3.3 Module 3 -- Structure-Function Analysis"));

children.push(new Paragraph({ children:[new TextRun({text:"Table 7. Structure information", bold:true, size:20})], spacing:{before:120, after:80}}));
children.push(mkTable(
  ["Field","Value"],
  [
    ["PDB ID","5XJH"],
    ["Title","Crystal structure of PETase from Ideonella sakaiensis"],
    ["Resolution","~1.5 A"],
    ["Chain analyzed","A"],
  ],
  [3200,6160]
));
children.push(body(
"The structure's secondary-structure pattern -- alternating beta-strand and alpha-helix segments -- is consistent with the alpha/beta-hydrolase fold identified computationally in Module 1 (InterPro IPR029058), providing an independent, structure-based confirmation of the sequence-based domain assignment. Two disulfide bridges are present in the modeled structure. The first, proximal to the active site, connects Cys203 and Cys239; this bond is reported in the literature as unique to IsPETase among characterized PET hydrolases, with related cutinases carrying alanine at both positions instead, and is described as contributing to active-site flexibility and catalytic activity (Joo et al., 2018; Chek et al., 2026) -- directly confirming the pattern this project's own Module 2 alignment had already found empirically. The second, more C-terminal bridge (Cys273-Cys289) is conserved across essentially all known PET-degrading hydrolases and plays a structural rather than catalytic role, located away from the active site."
));

children.push(new Paragraph({ children:[new TextRun({text:"Table 8. Catalytic triad -- structural confirmation", bold:true, size:20})], spacing:{before:200, after:80}}));
children.push(mkTable(
  ["Residue","Author position","Confirmed in PDB 5XJH"],
  [
    ["Ser","160","Yes"],
    ["Asp","206","Yes"],
    ["His","237","Yes"],
  ],
  [2200,3730,3430]
));

children.push(new Paragraph({ children:[new TextRun({text:"Table 9. Measured catalytic distances (PyMOL, PDB 5XJH)", bold:true, size:20})], spacing:{before:200, after:80}}));
children.push(mkTable(
  ["Atom pair","Distance (A)","Within hydrogen-bonding range (2.5-3.5 A)?"],
  [
    ["Ser160 (O-gamma) -- His237 (N-epsilon-2)","2.9","Yes"],
    ["His237 (N-delta-1) -- Asp206 (O-delta-1)","2.7","Yes"],
    ["His237 (N-delta-1) -- Asp206 (O-delta-2)","3.1","Yes"],
  ],
  [4200,2200,2960]
));

children.push(...figure("/home/claude/fig_triad_pymol.png", 340, 620, 530,
  "Figure 4. Ser160 (yellow), Asp206 (orange), and His237 (magenta) in the PDB 5XJH structure, with measured interatomic distances (PyMOL)."));
children.push(body(
"Both measured distance pairs fall within the 2.5-3.5 A range typical of a hydrogen bond, consistent with a functioning serine-hydrolase proton-relay mechanism in which His237 shuttles a proton between Ser160 and Asp206 during catalysis. This provides a quantitative, structure-based complement to the sequence-conservation evidence in Module 2."
));
children.push(body(
"One further residue is worth noting alongside the triad, though it lies outside this project's core evidence chain: Trp185. The literature describes this residue as adopting a 'wobbling' set of alternate side-chain conformations within a flexible loop connecting two beta-strands (the beta6-beta7 loop), a property enabled by two IsPETase-specific substitutions -- Ser214 and Ile218 -- at positions that are a conserved histidine and phenylalanine in related cutinases (Chen et al., 2021; Fecker et al., 2018; Crnjar et al., 2023; Burgin et al., 2024). This matches a small increase in modeled disorder observed near residues 185-195 during this project's own structural exploration of 5XJH, and is reported to assist substrate positioning and product release -- a dynamic complement to the fixed geometry of the catalytic triad."
));

// ---------------- 4. DISCUSSION ----------------
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(h1("4. Discussion"));

children.push(body(
"The central aim of this project was to test whether IsPETase's PET-degrading capability could be traced through three independent, increasingly direct lines of computational evidence: sequence, comparative conservation, and three-dimensional structure. The results converge cleanly. Direct indexing of the retrieved sequence located a Ser-Asp-His triad at positions 160, 206, and 237, embedded in a G-x-S-x-G nucleophile-elbow motif characteristic of serine hydrolases (Module 1). Multiple sequence alignment against five taxonomically diverse homologs -- spanning three Thermobifida species, a metagenome-derived cutinase (LCC), and a distantly related bacterial lipase used as an outgroup -- showed this same triad perfectly conserved in every sequence, despite overall identity to IsPETase dropping as low as 45% (Module 2). Finally, mapping these residues onto the experimentally solved 5XJH structure confirmed that they cluster spatially, at interatomic distances (2.7-3.1 A) consistent with active hydrogen bonding (Module 3)."
));
children.push(body(
"No single line of evidence here would be fully convincing alone: sequence conservation can occur for reasons unrelated to a specific catalytic role, and spatial proximity in a static crystal structure does not by itself prove a residue is catalytically active. Together, however, sequence conservation across divergent taxa, direct literature confirmation, and measured structural geometry constitute a convergent case that is considerably stronger than any one method in isolation -- which was the methodological premise of structuring this project as three connected modules rather than three independent exercises."
));
children.push(body(
"Two further findings, both confirmed once this project's own observations were checked against primary literature, extend this picture without changing its core argument. First, the cysteine at position 203 -- present in IsPETase and the outgroup lipase but replaced by alanine in all three Thermobifida cutinases and in LCC in this project's own Module 2 alignment -- forms a real disulfide bond with Cys239 in the 5XJH structure. This bond is independently reported in the literature as unique to IsPETase among characterized PET hydrolases and as a contributor to active-site flexibility and catalytic activity (Joo et al., 2018; Chek et al., 2026). Second, the small increase in modeled disorder observed near residues 185-195 in Module 3 corresponds to a well-documented mobile loop containing Trp185, whose wobbling side-chain conformations are reported to assist substrate binding and product release (Fecker et al., 2018; Chen et al., 2021; Crnjar et al., 2023; Burgin et al., 2024). Together with the conserved catalytic triad, these two additional features -- one rigidifying, one mobile -- sketch a more complete picture of how IsPETase's active site is organized to accommodate a bulky polymeric substrate that its cutinase relatives do not efficiently degrade."
));
children.push(body(
"A minor but genuinely instructive observation from Module 1 is the source organism's taxonomic history: named Ideonella sakaiensis in 2016, reclassified as Piscinibacter sakaiensis in 2023, and reclassified again as Pseudideonella sakaiensis in 2026. UniProt and NCBI displayed different names for the same organism during this project simply because their taxonomy databases had synchronized with the literature at different times -- a small but concrete reminder that reference databases are not static ground truth, and that reproducible bioinformatics work benefits from citing the specific accession and retrieval date rather than relying on a name alone."
));
children.push(body(
"The main limitation of this project is the phylogenetic analysis in Module 2, which used Clustal Omega's default neighbor-joining tree without bootstrap resampling; the two trees produced (pre- and post-alignment) disagreed on root-level branching order, and a negative branch length was observed. This is flagged transparently rather than resolved, since a dedicated phylogenetics tool with bootstrap support was outside the intended scope of a short computational project."
));

// ---------------- 5. CONCLUSION ----------------
children.push(h1("5. Conclusion"));
children.push(body(
"This project set out to ask how the sequence and structural organization of IsPETase relate to its PET-degrading function. Across three connected modules, the same three residues -- Ser160, Asp206, and His237 -- were independently identified from the raw sequence, found to be perfectly conserved across six taxonomically diverse comparison sequences despite substantial overall sequence divergence, and confirmed to sit in close spatial proximity within the solved crystal structure at distances consistent with active hydrogen bonding. Taken together, this constitutes a computationally grounded, multi-method case that IsPETase's catalytic capability rests on a classical serine-hydrolase active site, inherited and conserved within the broader alpha/beta-hydrolase superfamily, rather than on chemistry unique to this one organism."
));

// ---------------- REFERENCES ----------------
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(h1("References"));
const refs = [
"Yoshida, S., Hiraga, K., Takehana, T., et al. (2016). A bacterium that degrades and assimilates poly(ethylene terephthalate). Science, 351(6278), 1196-1199.",
"Han, X., Liu, W., Huang, J.W., et al. (2017). Structural insight into catalytic mechanism of PET hydrolase. Nature Communications, 8, 2106. DOI: 10.1038/s41467-017-02255-z",
"Austin, H.P., Allen, M.D., Donohoe, B., et al. (2018). Characterization and engineering of a plastic-degrading aromatic polyesterase. Proceedings of the National Academy of Sciences, 115(19), E4350-E4357. DOI: 10.1073/pnas.1718804115",
"Joo, S., Cho, I.J., Seo, H., et al. (2018). Structural insight into molecular mechanism of poly(ethylene terephthalate) degradation. Nature Communications, 9, 382. DOI: 10.1038/s41467-018-02881-1",
"Fecker, T., Galaz-Davison, P., Engelberger, F., et al. (2018). Active site flexibility as a hallmark for efficient PET degradation by I. sakaiensis PETase. Biophysical Journal, 114(6), 1302-1312. DOI: 10.1016/j.bpj.2018.02.005",
"Chen, C.C., Han, X., Li, X., et al. (2021). General features to enhance enzymatic activity of poly(ethylene terephthalate) hydrolysis. Nature Catalysis, 4, 425-430. DOI: 10.1038/s41929-021-00616-y",
"Crnjar, A., Griñen, A., Kamerlin, S.C.L., & Ramirez-Sarmiento, C. (2023). Conformational selection of a tryptophan side chain drives the generalized increase in activity of PET hydrolases through a Ser/Ile double mutation. ACS Organic & Inorganic Au. DOI: 10.1021/acsorginorgau.2c00054",
"Burgin, T., Pollard, B.C., Knott, B., et al. (2024). The reaction mechanism of the Ideonella sakaiensis PETase enzyme. Communications Chemistry, 7, 65. DOI: 10.1038/s42004-024-01154-x",
"Chek, M.F., Kawano, E., Sanuki, R., et al. (2026). Enhancing the specificity of a thermostable PET hydrolase toward aromatic polyesters via Piscinibacter sakaiensis PETase-inspired mutations. International Journal of Biological Macromolecules, 338, 149745. DOI: 10.1016/j.ijbiomac.2025.149745",
"Liu, B., et al. (2023). Reclassification of Ideonella sakaiensis as Piscinibacter sakaiensis.",
"Lu, H. & Chen, L. (2026). Emended description of the genus Rhizobacter and reclassification of Ideonella sakaiensis as Pseudideonella sakaiensis gen. nov., comb. nov. International Journal of Systematic and Evolutionary Microbiology, 76.",
"UniProt Consortium. UniProt entry A0A0K8P6T7 (PETH_PISS1). www.uniprot.org",
"InterPro entries IPR029058, IPR041127, IPR050261. www.ebi.ac.uk/interpro",
"RCSB Protein Data Bank, entry 5XJH. www.rcsb.org",
];
refs.forEach((r,i)=>{
  children.push(new Paragraph({
    children: [new TextRun({ text: `${i+1}. `, size: 20 }), new TextRun({ text: r, size: 20 })],
    spacing: { after: 140 },
    indent: { left: 360, hanging: 360 }
  }));
});
children.push(note("Note: references 2-9 above are drawn directly from the project's literature-review dataset (or independently verified against the publisher record) and include DOIs. The Science (Yoshida et al., 2016) and taxonomy (Liu et al., 2023) citations were recalled rather than re-fetched this session and are worth a quick spot-check before final submission."));

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN }
      }
    },
    headers: {
      default: new Header({ children: [new Paragraph({
        children: [new TextRun({ text: "Computational Investigation of IsPETase", size: 16, color: GREY })],
        alignment: AlignmentType.RIGHT
      })]})
    },
    footers: {
      default: new Footer({ children: [new Paragraph({
        children: [new TextRun({ text: "Page ", size: 16, color: GREY }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: GREY })],
        alignment: AlignmentType.CENTER
      })]})
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/home/claude/IsPETase_Report.docx", buf);
  console.log("written");
});
