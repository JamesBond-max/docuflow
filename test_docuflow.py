import copy
import json
import tempfile
import unittest

from pathlib import Path
from openpyxl import load_workbook

from docuflow import (
    ROOT,
    load_input,
    normalize,
    generate,
    verify,
    benchmark
)


class DocuFlowTests(unittest.TestCase):

    def setUp(self):
        self.raw = load_input(ROOT/'sample.json')
        self.t = ROOT/'template.yaml'

    def test_input_equivalence(self):
        for ext in ('yaml', 'xml'):
            self.assertEqual(
                normalize(
                    load_input(ROOT/f'sample.{ext}')
                ),
                normalize(self.raw)
            )

    def test_range_and_boundary(self):
        _, rows = normalize(self.raw)

        self.assertEqual(
            sum(r[-1] == 'FAIL' for r in rows), 1
        )

        self.assertEqual(rows[2][-1], 'PASS')
        self.assertEqual(rows[3][-1], 'FAIL')

    def test_missing_metadata(self):
        del self.raw['revision']

        with self.assertRaises(ValueError):
            normalize(self.raw)

    def test_duplicate_parameter(self):
        self.raw['parameters'].append(
            copy.deepcopy(self.raw['parameters'][0])
        )

        with self.assertRaises(ValueError):
            normalize(self.raw)

    def test_invalid_numbers_and_bounds(self):
        for field, value in [
            ('value', 'NaN'),
            ('value', 'Infinity'),
            ('min', '999')
        ]:
            raw = copy.deepcopy(self.raw)
            raw['parameters'][0][field] = value

            with self.assertRaises(ValueError):
                normalize(raw)

    def test_all_formats(self):
        with tempfile.TemporaryDirectory() as tmp:
            report = generate(
                ROOT/'sample.json', self.t, tmp
            )

            self.assertEqual(
                report['metrics']['formats_verified'],
                6
            )

            self.assertEqual(
                report['metrics']['format_consistency'],
                1.0
            )

            self.assertTrue(
                (Path(tmp)/'workflow.png').is_file()
            )

            self.assertTrue(
                (Path(tmp)/'audit.json').is_file()
            )

    def test_tampered_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp)

            r = generate(
                ROOT/'sample.json', self.t, out
            )

            wb = load_workbook(out/'report.xlsx')
            wb['Specifications']['B2'] = '999'
            wb.save(out/'report.xlsx')
            wb.close()

            self.assertFalse(
                verify(out, r['rows'])['xlsx']
            )

    def test_multiple_pages_and_slides(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp)

            raw = copy.deepcopy(self.raw)

            raw['parameters'] = [
                {
                    **self.raw['parameters'][0],
                    'name': f'Parameter {i:02}'
                }
                for i in range(60)
            ]

            p = out/'large.json'

            p.write_text(
                json.dumps(raw),
                encoding='utf-8'
            )

            r = generate(
                p, self.t, out/'generated'
            )

            self.assertEqual(
                r['metrics']['parameters'], 60
            )

            self.assertEqual(
                r['metrics']['formats_verified'], 6
            )

    def test_template_reuse(self):
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp)
            t = out/'custom.yaml'

            t.write_text(
                'title: Alternate Review\n'
                'accent: "254060"\n'
                'footer: Draft only\n',
                encoding='utf-8'
            )

            generate(
                ROOT/'sample.json',
                t,
                out/'generated'
            )

            self.assertIn(
                'Alternate Review',
                (
                    out/'generated/report.html'
                ).read_text(encoding='utf-8')
            )

    def test_benchmark(self):
        r = benchmark(
            ROOT/'sample.json', self.t, 2
        )

        self.assertEqual(r['runs'], 2)

        self.assertGreater(
            r['median_automation_seconds'], 0
        )

        self.assertIsNone(
            r['estimated_time_reduction_percent']
        )


if __name__ == '__main__':
    unittest.main()
