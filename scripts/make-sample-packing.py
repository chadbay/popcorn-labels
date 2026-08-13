#!/usr/bin/env python3
"""Generates sample_packing_list.pdf — a fake Booster Club "Packing List Report"
with invented student names, count-matched to sample_orders.xlsx. Dev-only;
needs reportlab (pip install reportlab). Run: python3 scripts/make-sample-packing.py
"""
import os
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas

ROOMS = {  # teacher -> student count, matching sample_orders.xlsx
    'Aune': 12, 'Berry': 15, 'Davis': 13, 'Judy Hurst': 9, 'Martinez': 14,
    'Bernacki': 16, 'Brightwell': 11, 'Dunagan': 13, 'Martin': 17, 'Shipley': 12,
    'Espich': 18, 'Falsone': 14, 'Kissinger': 10, 'Richardson': 15, 'Slawson': 13,
    'Chamness': 20, 'Grant': 11, 'Sarah Hurst': 12, 'Medina': 14, 'Oden': 9,
    'Rivas': 16, 'Brookshire': 13, 'Hemphill': 15, 'Hernandez': 12, 'Ellis': 14,
    'Roe': 10, 'Rosier': 17, 'Fields': 19, 'Melanson': 12, 'Park': 16,
    'Platt': 11, 'Polivka': 14, 'Nguyen': 8,
}
UNSPECIFIED = 6

FIRST = ['Avery', 'Blake', 'Casey', 'Drew', 'Emery', 'Finley', 'Gray', 'Harper',
         'Indie', 'Jules', 'Kai', 'Lane', 'Morgan', 'Noel', 'Oakley', 'Parker',
         'Quinn', 'Reese', 'Sage', 'Tatum']
LAST = ['Abbott', 'Barnes', 'Cortez', 'Dalton', 'Ellison', 'Fleming', 'Granger',
        'Holloway', 'Ibarra', 'Jennings', 'Kessler', 'Lockhart', 'Mercer',
        'Norwood', 'Osborne', 'Pruitt', 'Quimby', 'Redmond', 'Sutton', 'Thatcher']

W, H = landscape(letter)  # 792 x 612
X = {'order': 40, 'date': 130, 'parent': 210, 'student': 380, 'option': 560, 'q': 600}

def fake_students(teacher, n):
    seed = sum(ord(c) for c in teacher)
    out = []
    for i in range(n):
        first = FIRST[(seed + i * 7) % len(FIRST)]
        last = LAST[(seed * 3 + i * 5) % len(LAST)]
        out.append(f'{first} {last}-{i}' if out and out[-1].startswith(first) else f'{first} {last}')
    return out

def draw_page(c, title, rows, page, total_pages):
    c.setFont('Helvetica-Bold', 11)
    c.drawString(X['order'], H - 40, title)
    c.setFont('Helvetica-Bold', 9)
    y = H - 60
    for key, label in [('order', 'Order #'), ('date', 'Date'), ('parent', 'Parent'),
                       ('student', 'Student'), ('option', 'Option'), ('q', 'Q')]:
        c.drawString(X[key], y, label)
    c.setFont('Helvetica', 9)
    y -= 16
    for i, (parent, student) in enumerate(rows):
        c.drawString(X['order'], y, f'02067-{15000000 + abs(hash(title + student)) % 999999}')
        c.drawString(X['date'], y, '08/1%d/2026' % (i % 9 + 1))
        c.drawString(X['parent'], y, parent)
        if student:
            c.drawString(X['student'], y, student)
        c.drawString(X['q'], y, '1')
        y -= 14
    c.setFont('Helvetica', 8)
    c.drawString(X['order'], 24,
                 f'Forest Trail Booster Club - Packing List Report run 08/12/2026 by Sample Data Page {page} of {total_pages}')
    c.showPage()

out = os.path.join(os.path.dirname(__file__), '..', 'sample_packing_list.pdf')
c = canvas.Canvas(out, pagesize=(W, H))
total = len(ROOMS) + 1
draw_page(c, 'UNSPECIFIED', [(f'Parent Unknown{i}', '') for i in range(UNSPECIFIED)], 1, total)
for p, (teacher, n) in enumerate(sorted(ROOMS.items()), start=2):
    students = fake_students(teacher, n)
    rows = [(f'P. {s.split()[-1]}', s) for s in students]
    draw_page(c, teacher, rows, p, total)
c.save()
print('Wrote', out)
