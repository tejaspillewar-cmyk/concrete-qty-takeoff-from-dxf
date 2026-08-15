import ezdxf

doc = ezdxf.readfile(r"C:\Users\OMEN\OneDrive\Desktop\Sample-1.dxf")
msp = doc.modelspace()

layers = [
    "STR-WALL-REG", 
    "STR-WALL-NS-100", 
    "STR-WALL-NS-150", 
    "STR-WALL-NS-200"
]

print("Wall layers analysis:")
for l in layers:
    ents = msp.query(f'*[layer=="{l}"]')
    types = {}
    for e in ents:
        t = e.dxftype()
        types[t] = types.get(t, 0) + 1
    
    print(f"\nLayer '{l}':")
    for t, count in types.items():
        print(f"  - {t}: {count}")
        
    # If there are hatches, let's see how many boundaries they have
    hatches = msp.query(f'HATCH[layer=="{l}"]')
    if len(hatches) > 0:
        h = hatches[0]
        print(f"  Sample HATCH handle: {h.dxf.handle}")
        print(f"  Number of path boundaries: {len(h.paths)}")
        
    # If polylines, check if closed
    polys = msp.query(f'LWPOLYLINE[layer=="{l}"]')
    if len(polys) > 0:
        closed = sum(1 for p in polys if p.closed)
        print(f"  Closed LWPOLYLINEs: {closed} / {len(polys)}")
