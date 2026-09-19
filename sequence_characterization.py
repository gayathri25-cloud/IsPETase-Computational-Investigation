seq = """MNFPRASRLMQAAVLGGLMAVSAAATAQTNPYARGPNPTAASLEASAGPFTVRSFTVSRP
SGYGAGTVYYPTNAGGTVGAIAIVPGYTARQSSIKWWGPRLASHGFVVITIDTNSTLDQP
SSRSSQQMAALRQVASLNGTSSSPIYGKVDTARMGVMGWSMGGGGSLISAANNPSLKAAA
PQAPWDSSTNFSSVTVPTLIFACENDSIAPVNSSALPIYDSMSRNAKQFLEINGGSHSCA
NSGNSNQALIGKKGVAWMKRFMDNDTRYSTFACENPNSTRVSDFRTANCS""".replace("\n","")

print("Length:", len(seq))

# check catalytic triad positions (1-indexed, as commonly cited in literature)
for pos in [160, 206, 237]:
    print(f"Residue at position {pos}: {seq[pos-1]}  (context: {seq[max(0,pos-6):pos+5]})")

# amino acid composition
from collections import Counter
counts = Counter(seq)
print("\nAmino acid composition:")
for aa in sorted(counts):
    pct = 100*counts[aa]/len(seq)
    print(f"  {aa}: {counts[aa]:3d}  ({pct:5.2f}%)")

# average residue masses (monoisotopic-average masses, Da) minus water, standard table
aa_mass = {
 'A':71.0788,'R':156.1875,'N':114.1038,'D':115.0886,'C':103.1388,
 'E':129.1155,'Q':128.1307,'G':57.0519,'H':137.1411,'I':113.1594,
 'L':113.1594,'K':128.1741,'M':131.1926,'F':147.1766,'P':97.1167,
 'S':87.0782,'T':101.1051,'W':186.2132,'Y':163.1760,'V':99.1326
}
water = 18.01528
mw = sum(aa_mass[a] for a in seq) + water
print(f"\nMolecular weight: {mw:,.2f} Da  ({mw/1000:.2f} kDa)")

# aliphatic index (Ikai 1980)
X = {a: 100*counts.get(a,0)/len(seq) for a in "AVIL"}
AI = X['A'] + 2.9*X['V'] + 3.9*(X['I']+X['L'])
print(f"Aliphatic index: {AI:.2f}")

# theoretical pI - bisection, standard pKa set (close to ExPASy/Bjellqvist scale)
pos_pk = {'Nterm':9.0,'K':10.0,'R':12.0,'H':5.98}
neg_pk = {'Cterm':2.0,'D':4.05,'E':4.45,'C':9.0,'Y':10.0}

def net_charge(pH):
    charge = 0.0
    charge += 1/(1+10**(pH-pos_pk['Nterm']))
    charge += counts.get('K',0)/(1+10**(pH-pos_pk['K']))
    charge += counts.get('R',0)/(1+10**(pH-pos_pk['R']))
    charge += counts.get('H',0)/(1+10**(pH-pos_pk['H']))
    charge -= 1/(1+10**(neg_pk['Cterm']-pH))
    charge -= counts.get('D',0)/(1+10**(neg_pk['D']-pH))
    charge -= counts.get('E',0)/(1+10**(neg_pk['E']-pH))
    charge -= counts.get('C',0)/(1+10**(neg_pk['C']-pH))
    charge -= counts.get('Y',0)/(1+10**(neg_pk['Y']-pH))
    return charge

lo, hi = 0.0, 14.0
for _ in range(100):
    mid = (lo+hi)/2
    if net_charge(mid) > 0:
        lo = mid
    else:
        hi = mid
print(f"Theoretical pI (approx): {(lo+hi)/2:.2f}")

# signal peptide check - where does the hydrophobic stretch end (rough eyeball marker only)
print("\nFirst 40 residues (signal-peptide region to inspect):", seq[:40])
