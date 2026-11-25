import argparse
import html
import json
import statistics
import tempfile
import xml.etree.ElementTree as ET
from decimal import Decimal, InvalidOperation
from pathlib import Path
from time import perf_counter

import pymupdf
import yaml
from docx import Document
from docx.shared import Inches as WordInches, RGBColor
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor as SlideColor
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.pdfmetrics import stringWidth

ROOT = Path(__file__).resolve().parent

HEADERS = [
    'Parameter', 'Value', 'Unit',
    'Minimum', 'Maximum', 'Status'
]

META = ('project', 'revision', 'interface', 'owner')
FORMATS = ('docx', 'xlsx', 'pptx', 'pdf', 'html', 'md')

PDF_X = [40, 204, 284, 338, 418, 498, 572]
PDF_TOP, PDF_ROW, PDF_LIMIT = 230, 24, 18


# INPUT PARSING

def load_input(path):
    path = Path(path)

    if path.suffix == '.json':
        return json.loads(
            path.read_text(encoding='utf-8-sig')
        )

    if path.suffix in ('.yaml', '.yml'):
        return yaml.safe_load(
            path.read_text(encoding='utf-8-sig')
        )

    if path.suffix == '.xml':
        root = ET.parse(path).getroot()

        if root.tag != 'specification':
            raise ValueError('Invalid XML root')

        return {
            **{k: root.findtext(k) for k in META},
            'parameters': [
                dict(p.attrib)
                for p in root.findall('parameters/parameter')
            ]
        }

    raise ValueError('Use JSON, YAML or XML')


# INPUT VALIDATION

def clean(value, limit=70):
    if not isinstance(value, str) or not value.strip():
        raise ValueError('Missing text field')

    if not value.isascii() or any(
        not 32 <= ord(c) <= 126 for c in value
    ):
        raise ValueError('Only printable ASCII is supported')

    value = ' '.join(value.split())

    if len(value) > limit:
        raise ValueError('Text field too long')

    return value


def numeric(value):
    try:
        n = Decimal(str(value))
    except (InvalidOperation, ValueError):
        raise ValueError('Invalid number') from None

    if (
        not n.is_finite()
        or abs(n) > Decimal('1e12')
        or n.as_tuple().exponent < -8
    ):
        raise ValueError(
            'Unsupported numeric range or precision'
        )

    return n


def normalize(raw):
    if not isinstance(raw, dict):
        raise ValueError('Expected specification object')

    meta = {k: clean(raw.get(k)) for k in META}
    params = raw.get('parameters')

    if (
        not isinstance(params, list)
        or not 1 <= len(params) <= 100
    ):
        raise ValueError('Expected 1 to 100 parameters')

    rows, seen = [], set()

    for p in params:
        if not isinstance(p, dict):
            raise ValueError('Invalid parameter')

        name = clean(p.get('name'), 32)
        unit = clean(p.get('unit'), 12)

        if name.casefold() in seen:
            raise ValueError('Duplicate parameter: ' + name)

        seen.add(name.casefold())

        v, lo, hi = (
            numeric(p.get(k))
            for k in ('value', 'min', 'max')
        )

        if lo > hi:
            raise ValueError('Invalid bounds: ' + name)

        rows.append([
            name,
            format(v, 'f'),
            unit,
            format(lo, 'f'),
            format(hi, 'f'),
            'PASS' if lo <= v <= hi else 'FAIL'
        ])

    return meta, rows


# REUSABLE REPORT TEMPLATE

def template(path):
    t = yaml.safe_load(
        Path(path).read_text(encoding='utf-8-sig')
    )

    if not isinstance(t, dict):
        raise ValueError('Invalid template')

    t['title'] = clean(t.get('title'))
    t['footer'] = clean(t.get('footer'))

    a = t.get('accent')

    if (
        not isinstance(a, str)
        or len(a) != 6
        or any(c not in '0123456789abcdefABCDEF' for c in a)
    ):
        raise ValueError(
            'Accent must be six hexadecimal digits'
        )

    return t


def summary(meta, rows):
    failed = sum(r[-1] == 'FAIL' for r in rows)

    return [
        *(f'{k.title()}: {meta[k]}' for k in META),
        f'Parameters: {len(rows)} | Pass: {len(rows)-failed} | Fail: {failed}'
    ]


# TECHNICAL DIAGRAM

def diagram(path, accent):
    img = Image.new('RGB', (1200, 200), 'white')
    d = ImageDraw.Draw(img)
    font = ImageFont.load_default(size=25)

    labels = [
        'Structured input',
        'Validation rules',
        'Shared template',
        'Verified reports'
    ]

    for i, label in enumerate(labels):
        x = 15 + 300*i

        d.rounded_rectangle(
            (x, 30, x+260, 170),
            15,
            fill='#'+accent
        )

        d.text(
            (x+130, 100),
            label,
            font=font,
            fill='white',
            anchor='mm'
        )

        if i < 3:
            d.line(
                (x+260, 100, x+292, 100),
                fill='#475569',
                width=5
            )

            d.polygon(
                [
                    (x+292, 100),
                    (x+280, 92),
                    (x+280, 108)
                ],
                fill='#475569'
            )

    img.save(path)


# WORD DOCUMENT GENERATION

def write_word(out, t, meta, rows):
    doc = Document()
    h = doc.add_heading(t['title'], 0)

    for r in h.runs:
        r.font.color.rgb = RGBColor.from_string(
            t['accent']
        )

    for line in summary(meta, rows):
        doc.add_paragraph(line)

    doc.add_picture(
        str(out/'workflow.png'),
        width=WordInches(6)
    )

    table = doc.add_table(rows=0, cols=6)
    table.style = 'Light Shading Accent 1'

    for row in [HEADERS, *rows]:
        for c, v in zip(table.add_row().cells, row):
            c.text = v

    doc.sections[0].footer.paragraphs[0].text = (
        t['footer']
    )

    doc.save(out/'report.docx')


# EXCEL WORKBOOK GENERATION

def write_excel(out, t, meta, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Specifications'

    for row in [HEADERS, *rows]:
        ws.append(row)

        for c in ws[ws.max_row]:
            # Preserve decimal text and prevent formulas.
            c.data_type = 's'

    for c in ws[1]:
        c.fill = PatternFill(
            'solid', fgColor=t['accent']
        )

        c.font = Font(
            bold=True, color='FFFFFF'
        )

    for r in ws.iter_rows(min_row=2):
        r[-1].fill = PatternFill(
            'solid',
            fgColor=(
                'DCFCE7'
                if r[-1].value == 'PASS'
                else 'FEE2E2'
            )
        )

    for col, width in zip(
        'ABCDEF', [36, 24, 15, 24, 24, 12]
    ):
        ws.column_dimensions[col].width = width

    ws.freeze_panes = 'B2'
    ws.auto_filter.ref = ws.dimensions

    s = wb.create_sheet('Summary', 0)

    for line in [
        t['title'], *summary(meta, rows), t['footer']
    ]:
        s.append([line])
        s.cell(s.max_row, 1).data_type = 's'

    s.column_dimensions['A'].width = 95

    s['A1'].font = Font(
        bold=True,
        size=18,
        color=t['accent']
    )

    wb.save(out/'report.xlsx')


# POWERPOINT PRESENTATION GENERATION

def write_slides(out, t, meta, rows):
    p = Presentation()

    p.slide_width = Inches(13.333)
    p.slide_height = Inches(7.5)

    def textbox(
        s, text, x, y, w, h, size,
        color='1E293B'
    ):
        b = s.shapes.add_textbox(
            Inches(x), Inches(y),
            Inches(w), Inches(h)
        )

        b.text_frame.text = text
        b.text_frame.word_wrap = True

        for para in b.text_frame.paragraphs:
            para.font.size = Pt(size)
            para.font.color.rgb = SlideColor.from_string(
                color
            )

    def slide(title):
        s = p.slides.add_slide(
            p.slide_layouts[6]
        )

        textbox(
            s, title, .5, .25,
            12.3, .85, 26, t['accent']
        )

        textbox(
            s, t['footer'], .5, 7.0,
            12.3, .35, 11
        )

        return s

    s = slide(t['title'])

    textbox(
        s, '\n'.join(summary(meta, rows)),
        .6, 1.25, 12, 3, 20
    )

    s.shapes.add_picture(
        str(out/'workflow.png'),
        Inches(.6),
        Inches(4.6),
        width=Inches(12)
    )

    for start in range(0, len(rows), 6):
        chunk = rows[start:start+6]

        s = slide(
            f'Specification review | {start+1}-{start+len(chunk)}'
        )

        tab = s.shapes.add_table(
            len(chunk)+1,
            6,
            Inches(.5),
            Inches(1.25),
            Inches(12.3),
            Inches(.78*(len(chunk)+1))
        ).table

        for j, width in enumerate(
            [3.1, 2.0, 1.4, 2.0, 2.0, 1.8]
        ):
            tab.columns[j].width = Inches(width)

        for i, row in enumerate([HEADERS, *chunk]):
            for j, value in enumerate(row):
                c = tab.cell(i, j)
                c.text = value
                c.fill.solid()

                c.fill.fore_color.rgb = (
                    SlideColor.from_string(
                        t['accent']
                        if i == 0
                        else (
                            'FEE2E2'
                            if j == 5 and value == 'FAIL'
                            else 'F1F5F9'
                        )
                    )
                )

                para = c.text_frame.paragraphs[0]
                para.font.size = Pt(14)
                para.font.bold = i == 0

                para.font.color.rgb = (
                    SlideColor.from_string(
                        'FFFFFF' if i == 0 else '1E293B'
                    )
                )

    p.save(out/'report.pptx')


# PDF REPORT GENERATION

def write_pdf(out, t, meta, rows):
    c = canvas.Canvas(
        str(out/'report.pdf'),
        pagesize=letter
    )

    _, height = letter

    def text(
        s, x, y, width, size=10,
        font='Helvetica', color='#1E293B'
    ):
        while size > 4 and stringWidth(s, font, size) > width:
            size -= .5

        c.setFont(font, size)
        c.setFillColor(color)
        c.drawString(x, height-y, s)

    for offset in range(0, len(rows), PDF_LIMIT):
        text(
            t['title'], 40, 42, 532, 20,
            'Helvetica-Bold', '#'+t['accent']
        )

        for i, line in enumerate(summary(meta, rows)):
            text(line, 40, 65+15*i, 532)

        c.drawImage(
            str(out/'workflow.png'),
            40, height-220, 532, 88
        )

        for i, row in enumerate(
            [HEADERS, *rows[offset:offset+PDF_LIMIT]]
        ):
            top = PDF_TOP+i*PDF_ROW

            for j, val in enumerate(row):
                x = PDF_X[j]
                w = PDF_X[j+1]-PDF_X[j]

                c.setFillColor(
                    '#'+t['accent']
                    if i == 0
                    else (
                        '#FEE2E2'
                        if j == 5 and val == 'FAIL'
                        else '#F1F5F9' if i%2
                        else '#FFFFFF'
                    )
                )

                c.rect(
                    x, height-top-PDF_ROW,
                    w, PDF_ROW, fill=1, stroke=0
                )

                text(
                    val, x+4, top+16, w-8, 10,
                    'Helvetica-Bold' if i == 0 else 'Helvetica',
                    '#FFFFFF' if i == 0 else '#1E293B'
                )

        text(t['footer'], 40, 750, 470, 9)

        text(
            f'Page {offset//PDF_LIMIT+1}',
            520, 750, 52, 9
        )

        c.showPage()

    c.save()


# HTML AND MARKDOWN GENERATION

def md_cell(value):
    return (
        html.escape(value)
        .replace('\\', '\\\\')
        .replace('|', '\\|')
    )


def md_row(row):
    return '| ' + ' | '.join(map(md_cell, row)) + ' |'


def write_web(out, t, meta, rows):
    e = html.escape

    headers = ''.join(
        f'<th>{e(v)}</th>' for v in HEADERS
    )

    body = ''.join(
        '<tr>'
        + ''.join(f'<td>{e(v)}</td>' for v in r)
        + '</tr>'
        for r in rows
    )

    page = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<title>{e(t['title'])}</title>
<style>
body {{
    font:16px Arial;
    max-width:1000px;
    margin:40px auto;
    padding:20px;
    color:#1e293b;
}}
h1,th {{color:#{t['accent']};}}
table {{border-collapse:collapse;width:100%;}}
td,th {{
    border:1px solid #cbd5e1;
    padding:10px;
    text-align:left;
}}
img {{width:100%;}}
</style>
</head>
<body>
<h1>{e(t['title'])}</h1>
{''.join(f'<p>{e(v)}</p>' for v in summary(meta, rows))}
<img src="workflow.png" alt="Documentation workflow"/>
<table>
<thead><tr>{headers}</tr></thead>
<tbody>{body}</tbody>
</table>
<footer>{e(t['footer'])}</footer>
</body>
</html>'''

    (out/'report.html').write_text(
        page, encoding='utf-8'
    )

    lines = [
        '# '+md_cell(t['title']),
        '',
        *['- '+md_cell(v) for v in summary(meta, rows)],
        '',
        '![Workflow](workflow.png)',
        '',
        md_row(HEADERS),
        md_row(['---']*6),
        *map(md_row, rows),
        '',
        md_cell(t['footer'])
    ]

    (out/'report.md').write_text(
        '\n'.join(lines)+'\n',
        encoding='utf-8'
    )


# AUTOMATED OUTPUT VERIFICATION

def verify(out, rows):
    out = Path(out)
    expected = [HEADERS, *rows]

    w = [
        [c.text for c in r.cells]
        for r in Document(
            out/'report.docx'
        ).tables[0].rows
    ]

    wb = load_workbook(
        out/'report.xlsx',
        read_only=True
    )

    try:
        x = [
            list(r)
            for r in wb['Specifications'].values
        ]
    finally:
        wb.close()

    p, p_ok = [], True

    for s in Presentation(out/'report.pptx').slides:
        for shape in s.shapes:
            if shape.has_table:
                chunk = [
                    [c.text for c in r.cells]
                    for r in shape.table.rows
                ]

                p_ok = p_ok and chunk[0] == HEADERS
                p.extend(chunk[1:])

    h = ET.parse(out/'report.html').getroot()

    hheader = [
        c.text for c in h.findall('.//thead/tr/th')
    ]

    hrows = [
        [c.text for c in r]
        for r in h.findall('.//tbody/tr')
    ]

    md = [
        l for l in (
            out/'report.md'
        ).read_text(encoding='utf-8').splitlines()
        if l.startswith('| ')
    ]

    pdf_rows, pdf_ok = [], True

    with pymupdf.open(out/'report.pdf') as pdf:
        pdf_ok = (
            len(pdf)
            == (len(rows)+PDF_LIMIT-1)//PDF_LIMIT
        )

        for page in pdf:
            for i in range(PDF_LIMIT+1):
                top = PDF_TOP+i*PDF_ROW

                row = [
                    page.get_textbox(
                        pymupdf.Rect(
                            PDF_X[j],
                            top,
                            PDF_X[j+1],
                            top+PDF_ROW
                        )
                    ).strip()
                    for j in range(6)
                ]

                if i == 0:
                    pdf_ok = pdf_ok and row == HEADERS
                elif any(row):
                    pdf_rows.append(row)

    return {
        'docx': w == expected,
        'xlsx': x == expected,
        'pptx': p_ok and p == rows,
        'html': hheader == HEADERS and hrows == rows,
        'md': md == [
            md_row(HEADERS),
            md_row(['---']*6),
            *map(md_row, rows)
        ],
        'pdf': pdf_ok and pdf_rows == rows
    }


# DOCUMENT AUTOMATION PIPELINE

def generate(input_path, template_path, out):
    start = perf_counter()

    meta, rows = normalize(load_input(input_path))
    t = template(template_path)

    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)

    diagram(out/'workflow.png', t['accent'])

    for writer in (
        write_word,
        write_excel,
        write_slides,
        write_pdf,
        write_web
    ):
        writer(out, t, meta, rows)

    checks = verify(out, rows)
    failed = sum(r[-1] == 'FAIL' for r in rows)

    report = {
        'metadata': meta,
        'headers': HEADERS,
        'rows': rows,
        'metrics': {
            'parameters': len(rows),
            'in_range': len(rows)-failed,
            'out_of_range': failed,
            'formats_generated': len(FORMATS),
            'formats_verified': sum(checks.values()),
            'format_consistency': (
                sum(checks.values())/len(FORMATS)
            ),
            'processing_seconds': round(
                perf_counter()-start, 6
            )
        },
        'format_checks': checks
    }

    (out/'audit.json').write_text(
        json.dumps(report, indent=2),
        encoding='utf-8'
    )

    if not all(checks.values()):
        raise RuntimeError(
            f'Output verification failed: {checks}'
        )

    return report


# PERFORMANCE BENCHMARK

def benchmark(
    input_path, template_path,
    runs=5, manual_seconds=None
):
    if not 1 <= runs <= 20:
        raise ValueError('Use 1 to 20 runs')

    if (
        manual_seconds is not None
        and not 0 < manual_seconds < float('inf')
    ):
        raise ValueError(
            'Manual baseline must be positive and finite'
        )

    times = []

    with tempfile.TemporaryDirectory() as tmp:
        for i in range(runs):
            start = perf_counter()

            generate(
                input_path,
                template_path,
                Path(tmp)/str(i)
            )

            times.append(perf_counter()-start)

    median = statistics.median(times)

    return {
        'runs': runs,
        'seconds_per_run': [
            round(v, 6) for v in times
        ],
        'median_automation_seconds': round(median, 6),
        'manual_baseline_seconds': manual_seconds,
        'estimated_time_reduction_percent': (
            round(100*(1-median/manual_seconds), 2)
            if manual_seconds else None
        )
    }


# COMMAND LINE INTERFACE

def main():
    p = argparse.ArgumentParser(
        description='DocuFlow engineering documentation automation'
    )

    sub = p.add_subparsers(
        dest='command', required=True
    )

    for cmd in ('generate', 'benchmark'):
        s = sub.add_parser(cmd)

        s.add_argument(
            '--input', type=Path,
            default=ROOT/'sample.json'
        )

        s.add_argument(
            '--template', type=Path,
            default=ROOT/'template.yaml'
        )

        s.add_argument(
            '--out', type=Path,
            default=ROOT/'output'
        )

        if cmd == 'benchmark':
            s.add_argument(
                '--runs', type=int, default=5
            )

            s.add_argument(
                '--manual-seconds', type=float
            )

    a = p.parse_args()

    if a.command == 'generate':
        report = generate(
            a.input, a.template, a.out
        )

        print(json.dumps(
            report['metrics'], indent=2
        ))

        print(
            'Format checks:',
            json.dumps(report['format_checks'])
        )

    else:
        report = benchmark(
            a.input, a.template,
            a.runs, a.manual_seconds
        )

        a.out.mkdir(
            parents=True, exist_ok=True
        )

        (a.out/'benchmark.json').write_text(
            json.dumps(report, indent=2),
            encoding='utf-8'
        )

        print(json.dumps(report, indent=2))

    print('Reports:', a.out.resolve())


if __name__ == '__main__':
    main()
