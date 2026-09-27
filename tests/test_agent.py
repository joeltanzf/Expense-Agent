"""Tests use fictional transactions and temporary storage only."""
from contextlib import closing
from datetime import date
from decimal import Decimal
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch
import uuid
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'app'))
import agent


def transaction(number, day='2026-01-10'):
    return dict(id=f'synthetic-{number}', date=day, description=f'Example shop {number}',
                amount='5.00', type='Payment', reference=f'EXAMPLE-{number}',
                source='example.pdf', page=1)


class ExpenseTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.old_data, self.old_work = agent.DATA, agent.WORK
        agent.DATA, agent.WORK = Path(self.temp.name)/'data', Path(self.temp.name)/'work'
        self.db = agent.connect()

    def tearDown(self):
        self.db.close()
        agent.DATA, agent.WORK = self.old_data, self.old_work
        self.temp.cleanup()

    def test_overlapping_statements_skip_saved_transactions(self):
        with patch.object(agent, 'secret', return_value=''), patch.object(agent, 'export_workbook', return_value={'synced': True}):
            with patch.object(agent, 'extract_pdf', return_value=('first', [transaction(1), transaction(2)], 0, 2)):
                first = agent.import_pdf(self.db, 'first.pdf')
                repeat = agent.import_pdf(self.db, 'first.pdf')
            with patch.object(agent, 'extract_pdf', return_value=('second', [transaction(2), transaction(3)], 0, 2)):
                overlap = agent.import_pdf(self.db, 'second.pdf')
        self.assertEqual([first['added'], repeat['added'], overlap['added']], [2, 0, 1])
        self.assertEqual(overlap['duplicates'], 1)
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM transactions').fetchone()[0], 3)

    def test_repeated_reference_cannot_change_saved_amount(self):
        row = transaction(1)
        self.db.execute('INSERT INTO transactions VALUES (:id,:date,:description,:amount,:type,:reference,:source,:page)', row)
        self.db.commit()
        with patch.object(agent, 'secret', return_value=''), patch.object(agent, 'extract_pdf', return_value=('conflict', [dict(row, amount='50.00')], 0, 1)):
            with self.assertRaises(agent.AgentError):
                agent.import_pdf(self.db, 'conflict.pdf')
        self.assertEqual(self.db.execute('SELECT amount FROM transactions').fetchone()[0], '5.00')

    def test_manual_entry_retry_and_pending_export(self):
        request = dict(date='2026-02-03', description='Example lunch', price='12.50', entry_id=str(uuid.uuid4()))
        with patch.object(agent, 'export_workbook', return_value={'synced': False, 'message': 'Workbook open'}):
            result = agent.add_expense(self.db, request)
            agent.add_expense(self.db, request)
        self.assertTrue(result['saved'])
        self.assertEqual(agent.get_setting(self.db, 'dirty'), '1')
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM transactions').fetchone()[0], 1)
        self.assertEqual(self.db.execute('SELECT sheet FROM months').fetchone()[0], 'February')

    def test_invalid_manual_entry_saves_nothing(self):
        base = dict(date='2026-02-03', description='Example lunch', price='12.50', entry_id=str(uuid.uuid4()))
        for change in [dict(date='2026-02-30'), dict(price='-1'), dict(price='NaN'), dict(description='')]:
            with self.subTest(change=change), self.assertRaises(agent.AgentError):
                agent.add_expense(self.db, dict(base, **change))
        self.assertEqual(self.db.execute('SELECT COUNT(*) FROM transactions').fetchone()[0], 0)

    def test_month_names_remain_unique_across_years(self):
        agent.month_add(self.db, '2026-01-01')
        agent.ensure_months(self.db, date(2027, 1, 1))
        self.assertEqual(self.db.execute("SELECT sheet FROM months WHERE month='2027-01'").fetchone()[0], 'January 2027')

    def test_description_rules_preserve_original_text(self):
        rules = [{'pattern': 'example shop', 'replacement': 'Groceries', 'mode': 'contains'}]
        self.assertEqual(agent.apply_rules('Example Shop 1', rules), 'Groceries')
        self.assertEqual(agent.apply_rules('Other merchant', rules), 'Other merchant')

    def test_real_export_sorted_dates_and_cached_total(self):
        for row in [transaction(2, '2026-01-20'), transaction(1, '2026-01-05')]:
            self.db.execute('INSERT INTO transactions VALUES (:id,:date,:description,:amount,:type,:reference,:source,:page)', row)
            agent.month_add(self.db, row['date'])
        self.db.commit()
        result = agent.export_workbook(self.db)
        self.assertTrue(result['synced'])
        ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        with zipfile.ZipFile(agent.workbook_path('2026')) as archive:
            sheet = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        cells = {cell.get('r'): cell for cell in sheet.findall('.//s:c', ns)}
        self.assertLess(float(cells['A2'].find('s:v', ns).text), float(cells['A3'].find('s:v', ns).text))
        self.assertEqual(cells['C5'].find('s:f', ns).text, 'SUM(C2:C3)')
        self.assertEqual(Decimal(cells['C5'].find('s:v', ns).text), Decimal('10'))


if __name__ == '__main__':
    unittest.main(verbosity=2)
