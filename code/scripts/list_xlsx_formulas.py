import zipfile
import re
from pathlib import Path

p = Path(__file__).resolve().parents[1] / 'assets' / 'colisa_logiciel_template.xlsx'
if not p.exists():
    print('MISSING', p)
    raise SystemExit(1)

z = zipfile.ZipFile(p)
out = []
for name in z.namelist():
    if name.startswith('xl/worksheets/'):
        data = z.read(name)
        try:
            s = data.decode('utf-8')
        except Exception:
            try:
                s = data.decode('cp1252')
            except Exception:
                continue
        formulas = re.findall(r'<f[^>]*>(.*?)</f>', s, flags=re.DOTALL)
        if formulas:
            out.append('--- ' + name)
            for f in formulas:
                out.append(f.strip())

dest = Path(__file__).resolve().parents[0] / 'formulas_text.txt'
dest.write_text('\n'.join(out), encoding='utf-8')
print('WROTE', dest)
