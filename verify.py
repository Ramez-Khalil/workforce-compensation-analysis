"""Validate the synthetic dataset and saved XLSX using only the standard library."""
import csv, json, math, zipfile
from collections import Counter
from datetime import date
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent
with (ROOT / 'synthetic-workforce.csv').open(newline='') as f:
    rows = list(csv.DictReader(f))
assert len(rows) == 747
assert len({r['Record ID'] for r in rows}) == len(rows)
assert all(r['Record ID'].startswith('SYN-') for r in rows)
salary = [int(r['Annual salary USD']) for r in rows]
assert all(s > 0 for s in salary)
snapshot = date(2025, 12, 31)
for r in rows:
    hire = date.fromisoformat(r['Hire date'])
    assert hire <= snapshot
    if r['Termination date']:
        assert hire <= date.fromisoformat(r['Termination date']) <= snapshot
    assert all(1 <= int(r[k]) <= 5 for k in ['Knowledge','Communication','Teamwork','Leadership','Quality'])
expected_total = sum(salary)
expected_active = sum(not r['Termination date'] for r in rows)
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
def cells(xml):
    return {c.attrib['r']: c for c in ET.fromstring(xml).findall('.//s:c', ns)}
def val(c):
    e = c.find('s:v', ns)
    return float(e.text) if e is not None and e.text is not None else None
with zipfile.ZipFile(ROOT / 'workforce-compensation.xlsx') as z:
    summary = cells(z.read('xl/worksheets/sheet1.xml'))
    data = cells(z.read('xl/worksheets/sheet2.xml'))
    assert val(summary['A9']) == len(rows)
    assert val(summary['B9']) == expected_active
    assert val(summary['D9']) == expected_total
    assert math.isclose(val(summary['C9']), expected_total / len(rows))
    for ref in ['B36','B37','B38','B39']:
        assert val(summary[ref]) == 0
    for n in [6,379,752]:
        for col in ['L','M','N']:
            assert data[f'{col}{n}'].find('s:f',ns) is not None
    errors = [c.attrib['r'] for c in list(summary.values()) + list(data.values()) if c.attrib.get('t') == 'e']
    assert not errors, errors
    charts = [name for name in z.namelist() if '/charts/chart' in name and name.endswith('.xml')]
    assert len(charts) == 1
    chart = z.read(charts[0]).decode()
    assert "'Compensation'!" in chart and '$C$14:$C$19' in chart and '$D$14:$D$19' in chart
    assert not any(name.startswith('xl/externalLinks/') for name in z.namelist())
    assert not any('vbaProject' in name for name in z.namelist())
print(json.dumps({'status':'passed','records':len(rows),'active':expected_active,'annual_salary_rate_total':expected_total,'mean_salary':expected_total/len(rows),'occupation_counts':dict(Counter(r['Occupation'] for r in rows))}, indent=2))
