"""
Generate complete 129 Excel Formula Metadata for x64 Assembly.
"""
from xlsx_viewer.formulas import FUNCTION_METADATA

cat_map = {
    'Math & Trig': 1,
    'Statistical': 2,
    'Lookup & Reference': 3,
    'Logical': 4,
    'Text': 5,
    'Date & Time': 6,
    'Financial': 7,
    'Information': 8,
}

lines = []
lines.append('; Generated 129 Excel Formula Metadata for x64 Assembly\n')

sorted_funcs = sorted(FUNCTION_METADATA.items())

for idx, (name, meta) in enumerate(sorted_funcs):
    cat_id = cat_map.get(meta.get('cat'), 0)
    syn = meta.get('syntax', f'={name}()').replace("'", "''")
    if not syn.startswith('='):
        syn = '=' + syn
    desc = meta.get('desc', '').replace("'", "''")
    
    lines.append(f"fn_n_{idx} db '{name}', 0")
    lines.append(f"fn_s_{idx} db '{syn}', 0")
    lines.append(f"fn_d_{idx} db '{desc}', 0")

lines.append('\nalign 8\nfn_all_names:')
chunk = []
for idx in range(len(sorted_funcs)):
    chunk.append(f"fn_n_{idx}")
    if len(chunk) == 8:
        lines.append("    dq " + ", ".join(chunk))
        chunk = []
if chunk:
    lines.append("    dq " + ", ".join(chunk))

lines.append('\nalign 8\nfn_all_syntax:')
chunk = []
for idx in range(len(sorted_funcs)):
    chunk.append(f"fn_s_{idx}")
    if len(chunk) == 8:
        lines.append("    dq " + ", ".join(chunk))
        chunk = []
if chunk:
    lines.append("    dq " + ", ".join(chunk))

lines.append('\nalign 8\nfn_all_desc:')
chunk = []
for idx in range(len(sorted_funcs)):
    chunk.append(f"fn_d_{idx}")
    if len(chunk) == 8:
        lines.append("    dq " + ", ".join(chunk))
        chunk = []
if chunk:
    lines.append("    dq " + ", ".join(chunk))

lines.append('\nalign 4\nfn_all_cats:')
chunk = []
for idx, (name, meta) in enumerate(sorted_funcs):
    cid = cat_map.get(meta.get('cat'), 0)
    chunk.append(str(cid))
    if len(chunk) == 16:
        lines.append("    dd " + ", ".join(chunk))
        chunk = []
if chunk:
    lines.append("    dd " + ", ".join(chunk))

with open('asm/funcs_data.inc', 'w', encoding='ascii', errors='replace') as f:
    f.write('\n'.join(lines) + '\n')

print(f"Generated asm/funcs_data.inc with {len(sorted_funcs)} functions.")
