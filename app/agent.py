"""Local TNG PDF -> SQLite -> monthly Excel agent."""
from __future__ import annotations

import base64
import calendar
import ctypes
from ctypes import wintypes
from contextlib import closing
from datetime import date, datetime
from decimal import Decimal
import hashlib
import io
import json
import os
from pathlib import Path
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import uuid

import pdfplumber
from xlsx_compat import validate_xlsx
from export_workbook import write_workbook

APP = Path(__file__).resolve().parent
PROJECT = APP.parent
DATA = Path(os.environ.get('TNG_AGENT_DATA', Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'TNGExpenseAgent' / 'data'))
WORK = Path(os.environ.get('TNG_AGENT_WORK', DATA.parent / 'work'))
DATE_RE = re.compile(r'^\d{1,2}/\d{1,2}/\d{4}$')
MONEY_RE = re.compile(r'^RM\s*([+-]?[\d,]+\.\d{2,4})$')
EXCLUDED = {'Reload', 'Receive from Wallet', 'eWallet Cash Out', 'GO+ Cash In',
            'GO+ Cash Out', 'GO+ Daily Earnings', 'Gold Withdraw', 'eWallet Cash In',
            'DUITNOW_RECEIVE', 'DUITNOW_RECEI VEFROM'}
EXPENSES = {'DuitNow QR', 'DuitNow QR TNGD', 'Transfer to Wallet', 'Payment',
            'PayDirect', 'RFID', 'Toll', 'Parking', 'DuitNow Transfer', 'DUITNOW_TRANS FERTO'}


class AgentError(Exception):
    pass


def digest(value):
    return hashlib.sha256(value if isinstance(value, bytes) else value.encode()).hexdigest()


def protect(value, decrypt=False):
    """Windows DPAPI binds saved PDF passwords to this Windows user."""
    if os.name != 'nt':
        raise AgentError('Saved passwords require Windows on this version of the app.')
    class Blob(ctypes.Structure):
        _fields_ = [('size', wintypes.DWORD), ('data', ctypes.POINTER(ctypes.c_byte))]
    raw = base64.b64decode(value) if decrypt else value.encode('utf-8')
    buf = ctypes.create_string_buffer(raw)
    src = Blob(len(raw), ctypes.cast(buf, ctypes.POINTER(ctypes.c_byte)))
    dst = Blob()
    crypt = ctypes.WinDLL('crypt32', use_last_error=True)
    fn = crypt.CryptUnprotectData if decrypt else crypt.CryptProtectData
    fn.argtypes = [ctypes.POINTER(Blob), ctypes.c_void_p, ctypes.c_void_p,
                   ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(Blob)]
    fn.restype = wintypes.BOOL
    if not fn(ctypes.byref(src), None, None, None, None, 1, ctypes.byref(dst)):
        raise AgentError('Windows could not save or unlock this password. You can enter it for one upload instead, without saving it.')
    try:
        result = ctypes.string_at(dst.data, dst.size)
        return result.decode('utf-8') if decrypt else base64.b64encode(result).decode('ascii')
    finally:
        ctypes.windll.kernel32.LocalFree.argtypes = [ctypes.c_void_p]
        ctypes.windll.kernel32.LocalFree(ctypes.cast(dst.data, ctypes.c_void_p))


def connect():
    DATA.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(DATA / 'transactions.sqlite3', timeout=60)
    db.row_factory = sqlite3.Row
    db.executescript('''
        CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS months(month TEXT PRIMARY KEY, sheet TEXT NOT NULL UNIQUE);
        CREATE TABLE IF NOT EXISTS imports(hash TEXT PRIMARY KEY, filename TEXT, imported_at TEXT,
          engine TEXT, row_count INTEGER, excluded_count INTEGER);
        CREATE TABLE IF NOT EXISTS transactions(id TEXT PRIMARY KEY, date TEXT NOT NULL,
          description TEXT NOT NULL, amount TEXT NOT NULL, type TEXT NOT NULL,
          reference TEXT NOT NULL, source TEXT NOT NULL, page INTEGER NOT NULL);
        CREATE TABLE IF NOT EXISTS rules(id INTEGER PRIMARY KEY, pattern TEXT NOT NULL,
          replacement TEXT NOT NULL, mode TEXT NOT NULL, UNIQUE(pattern,mode));
        CREATE TABLE IF NOT EXISTS pdf_matches(id TEXT PRIMARY KEY, manual_id TEXT NOT NULL UNIQUE,
          date TEXT NOT NULL, description TEXT NOT NULL, amount TEXT NOT NULL,
          type TEXT NOT NULL, reference TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS deleted_transactions(id TEXT PRIMARY KEY, deleted_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS year_exports(year TEXT PRIMARY KEY, file_hash TEXT NOT NULL DEFAULT '',
          payload_hash TEXT NOT NULL DEFAULT '', pending INTEGER NOT NULL DEFAULT 1, reason TEXT NOT NULL DEFAULT '');
    ''')
    return db


def get_setting(db, key, default=''):
    row = db.execute('SELECT value FROM settings WHERE key=?', (key,)).fetchone()
    return row[0] if row else default


def set_setting(db, key, value):
    db.execute('INSERT OR REPLACE INTO settings VALUES (?,?)', (key, str(value)))


def secret(db, key):
    value = get_setting(db, key)
    return protect(value, True) if value else ''


def month_add(db, value):
    ym = value[:7]
    if db.execute('SELECT 1 FROM months WHERE month=?', (ym,)).fetchone():
        return
    year, month = map(int, ym.split('-'))
    name = calendar.month_name[month]
    if db.execute('SELECT 1 FROM months WHERE sheet=?', (name,)).fetchone():
        name = f'{name} {year}'
    db.execute('INSERT INTO months VALUES (?,?)', (ym, name))


def ensure_months(db, today=None):
    today = today or date.today()
    last = db.execute('SELECT MAX(month) FROM months WHERE month<=?', (today.strftime('%Y-%m'),)).fetchone()[0]
    if not last:
        month_add(db, today.isoformat())
        return
    # Catch up from the latest non-future month; future expenses cannot block today.
    year, month = map(int, last.split('-'))
    while (year, month) <= (today.year, today.month):
        month_add(db, f'{year:04d}-{month:02d}-01')
        year, month = (year+1, 1) if month == 12 else (year, month+1)


def text_in(words, left, right):
    selected = [w for w in words if left-0.3 <= w['x0'] < right-0.3]
    return ' '.join(w['text'] for w in sorted(selected, key=lambda w: (round(w['top']), w['x0'])))


def extract_pdf(path, password):
    path = Path(path)
    raw = path.read_bytes()
    if len(raw) > 30*1024*1024:
        raise AgentError('Please use a PDF smaller than 30 MB.')
    if not raw.startswith(b'%PDF-'):
        raise AgentError('This file is not a PDF.')
    rows, account = [], ''
    try:
        with pdfplumber.open(io.BytesIO(raw), password=password) as pdf:
            if len(pdf.pages) > 150:
                raise AgentError('Please use a statement with 150 pages or fewer.')
            first = pdf.pages[0].extract_text() or ''
            if 'TNG WALLET TRANSACTION HISTORY' not in first:
                raise AgentError('This version supports text-based TNG eWallet statements. This PDF has a different layout or needs OCR.')
            match = re.search(r'Wallet ID\s+(\d+)', first)
            if not match:
                raise AgentError('The statement wallet ID could not be read.')
            account = digest(match.group(1))
            section = 'wallet'
            for page_num, page in enumerate(pdf.pages, 1):
                # Adjacent columns can touch (e.g. "Receive from Wallet" and
                # its reference). Split characters by column BEFORE making words.
                boundaries = [0,80,140,230,295,460,650,745,page.width]
                words = []
                for left,right in zip(boundaries,boundaries[1:]):
                    column = page.filter(lambda obj,lo=left,hi=right:
                        obj.get('object_type')=='char' and lo-0.3 <= obj['x0'] < hi-0.3)
                    words.extend(column.extract_words(x_tolerance=1,y_tolerance=2))
                if not words:
                    raise AgentError(f'Page {page_num} has no readable text. No transactions were saved.')
                if abs(page.width-842) > 3:
                    raise AgentError(f'Page {page_num} has an unfamiliar layout. No transactions were saved.')
                anchors = [w for w in words if w['x0'] < 65 and DATE_RE.fullmatch(w['text'])]
                headings = [w['top'] for w in words if w['text'] == 'GO+' and w['x0'] < 30]
                footers = [w['top'] for w in words if w['text'].startswith('*This')]
                for i, anchor in enumerate(anchors):
                    top = anchor['top']-1
                    if any(y < top for y in headings):
                        section = 'go+'
                    limits = [page.height-15] + [y-1 for y in footers+headings if y > top+2]
                    if i+1 < len(anchors):
                        limits.append(anchors[i+1]['top']-1)
                    bottom = min(limits)
                    band = [w for w in words if top <= w['top'] < bottom]
                    status = text_in(band, 80, 140)
                    kind = text_in(band, 140, 230)
                    reference = text_in(band, 230, 295).replace(' ', '')
                    description = text_in(band, 295, 460)
                    amount_text = text_in(band, 650, 745).replace(' ', '')
                    money = MONEY_RE.fullmatch(amount_text)
                    if not money or not description or not reference or not kind:
                        raise AgentError(f'Could not read a complete transaction on page {page_num}. No transactions were saved.')
                    amount = Decimal(money.group(1).replace(',', ''))
                    if not amount.is_finite():
                        raise AgentError('Invalid amount in PDF.')
                    day = datetime.strptime(anchor['text'], '%d/%m/%Y').date().isoformat()
                    rows.append({'id':digest(account+'|'+section+'|'+reference), 'date':day,
                        'description':description, 'amount':str(amount), 'type':kind,
                        'reference':reference, 'page':page_num, 'section':section,
                        'status':status, 'source':path.name})
    except AgentError:
        raise
    except Exception as exc:
        if 'password' in type(exc).__name__.lower() or 'password' in str(exc).lower():
            raise AgentError('The PDF password is incorrect. Update it in Settings.') from None
        raise AgentError('The PDF could not be read. Check the password and use an original TNG statement.') from None
    if not rows:
        raise AgentError('No transactions were found. No data was saved.')
    expenses = []
    for row in rows:
        if row['status'] != 'Success' or row['section'] == 'go+' or row['type'] in EXCLUDED:
            continue
        if row['type'] not in EXPENSES:
            raise AgentError(f"Unrecognized transaction type: {row['type']}. The statement needs review before importing.")
        expenses.append(row)
    return digest(raw), expenses, len(rows)-len(expenses), len(rows)


def apply_rules(description, rules):
    for rule in rules:
        match = rule['pattern'].casefold()
        original = description.casefold()
        if (rule['mode']=='exact' and original==match) or (rule['mode']=='contains' and match in original):
            return rule['replacement']
    return description


def valid_text(value):
    return isinstance(value, str) and bool(value) and len(value) <= 300 and all(
        32 <= ord(c) <= 0x10ffff and not 0xd800 <= ord(c) <= 0xdfff and
        ord(c) not in (0xfffe, 0xffff) for c in value)


def backup_database(db):
    """Keep bounded, consistent snapshots of transactions, rules and settings."""
    folder = DATA / 'backups'
    folder.mkdir(exist_ok=True)
    # Hash logical state, so retries do not create repeated identical backups.
    tables = ('transactions', 'rules', 'months', 'imports', 'pdf_matches', 'settings', 'deleted_transactions', 'year_exports')
    logical = {t: [tuple(r) for r in db.execute('SELECT * FROM ' + t + ' ORDER BY 1')]
               for t in tables}
    fingerprint = digest(json.dumps(logical, sort_keys=True, ensure_ascii=True))
    if any(folder.glob('Database-*-' + fingerprint[:24] + '.sqlite3')):
        return
    name = 'Database-' + datetime.now().strftime('%Y%m%d-%H%M%S-%f') + '-' + fingerprint[:24] + '.sqlite3'
    temporary = folder / ('.' + uuid.uuid4().hex + '.tmp')
    try:
        with closing(sqlite3.connect(temporary)) as target:
            db.backup(target)
        os.replace(temporary, folder / name)
    finally:
        temporary.unlink(missing_ok=True)
    for old in sorted(folder.glob('Database-*.sqlite3'), reverse=True)[30:]:
        old.unlink()


def workbook_path(year):
    if not re.fullmatch(r'\d{4}', str(year)):
        raise AgentError('Choose a valid year.')
    return DATA / f'Expenses-{year}.xlsx'


def _export_year(db, payload, year, force=False):
    if any(not valid_text(r['description']) for r in payload['transactions']):
        raise AgentError('A saved description contains an unsupported character. Correct its description rule before exporting.')
    path = workbook_path(year)
    previous = db.execute('SELECT * FROM year_exports WHERE year=?', (year,)).fetchone()
    fingerprint = previous['file_hash'] if previous else ''
    payload_hash = digest(json.dumps(payload, sort_keys=True, ensure_ascii=True))
    if path.exists() and fingerprint and digest(path.read_bytes()) != fingerprint and not force:
        return {'synced':False, 'reason':'workbook_modified', 'message':f'{path.name} was edited outside the app. Select {year} and use Back up and rebuild selected year to refresh it from saved transactions.'}
    if path.exists() and not fingerprint and not force:
        return {'synced':False, 'reason':'workbook_modified', 'message':f'{path.name} already exists but was not created by this app. Select {year} and rebuild it to back it up before replacement.'}
    if previous and previous['payload_hash'] == payload_hash and path.exists() and not force:
        return {'synced':True, 'message':f'{path.name} is up to date.'}
    WORK.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='export-', dir=WORK) as temp:
        tmp = Path(temp)
        output_path = tmp/'Expenses.xlsx'
        try:
            write_workbook(payload, output_path)
            validate_xlsx(output_path)
        except (ValueError, OSError):
            raise AgentError('Transactions are saved, but Excel generation failed. The previous workbook was kept.') from None
        # Windows private temporary directories have owner-only ACLs. Renaming
        # their files into DATA preserves those ACLs, preventing the real user
        # from opening the workbook. Create a fresh sibling with normal DATA
        # inheritance, copy bytes only, then replace the destination atomically.
        staging_path = DATA / ('.Expenses-' + uuid.uuid4().hex + '.xlsx')
        try:
            with output_path.open('rb') as source, staging_path.open('xb') as destination:
                shutil.copyfileobj(source, destination)
            if path.exists():
                backup = DATA/'backups'
                backup.mkdir(exist_ok=True)
                previous_hash = digest(path.read_bytes())
                if not any(backup.glob(path.stem + '-*-' + previous_hash[:24] + '.xlsx')):
                    stamp = datetime.now().strftime('%Y%m%d-%H%M%S-%f')
                    shutil.copy2(path, backup/f'{path.stem}-{stamp}-{previous_hash[:24]}.xlsx')
            os.replace(staging_path, path)
        except PermissionError:
            return {'synced':False,'reason':'workbook_open','message':f'Changes are saved. Close {path.name} and keep the agent open; it will update automatically.'}
        finally:
            staging_path.unlink(missing_ok=True)
    db.execute('''INSERT INTO year_exports(year,file_hash,payload_hash,pending,reason) VALUES (?,?,?,0,'')
        ON CONFLICT(year) DO UPDATE SET file_hash=excluded.file_hash,payload_hash=excluded.payload_hash,pending=0,reason='' ''',
        (year,digest(path.read_bytes()),payload_hash))
    db.commit()
    for old in sorted((DATA/'backups').glob('Expenses-*.xlsx'), reverse=True)[30:]:
        old.unlink()
    return {'synced':True,'message':f'{path.name} is up to date.'}


def export_workbook(db, force=False, year=None):
    """Update each year independently; a locked year cannot block other years."""
    months = [dict(r) for r in db.execute('SELECT * FROM months ORDER BY month')]
    if not months:
        with db:
            ensure_months(db)
        months = [dict(r) for r in db.execute('SELECT * FROM months ORDER BY month')]
    years = sorted({m['month'][:4] for m in months})
    if year is not None and str(year) not in years:
        raise AgentError('Choose a year that has a saved monthly sheet.')
    backup_database(db)
    rules = [dict(r) for r in db.execute("SELECT * FROM rules ORDER BY CASE mode WHEN 'exact' THEN 0 ELSE 1 END,id")]
    rows = [dict(r) for r in db.execute('SELECT * FROM transactions ORDER BY date,rowid')]
    results = []
    for value in years:
        payload = {
            'months':[dict(m, sheet=calendar.month_name[int(m['month'][5:7])]) for m in months if m['month'].startswith(value)],
            'transactions':[dict(r, description=apply_rules(r['description'],rules)) for r in rows if r['date'].startswith(value)]}
        try:
            result = _export_year(db,payload,value,force=force and (year is None or str(year)==value))
        except PermissionError:
            result = {'synced':False,'reason':'workbook_open','message':f'Close Expenses-{value}.xlsx to finish updating it.'}
        except (AgentError, OSError):
            result = {'synced':False,'reason':'export_failed','message':f'Changes are saved, but Expenses-{value}.xlsx could not be updated. Its previous version was kept. Please retry Refresh Excel.'}
        with db:
            db.execute('''INSERT INTO year_exports(year,pending,reason) VALUES (?,?,?)
                ON CONFLICT(year) DO UPDATE SET pending=excluded.pending,reason=excluded.reason''',
                (value,int(not result['synced']),result.get('reason','')))
        results.append(dict(result,year=value))
    pending = [r for r in results if not r['synced']]
    with db:
        set_setting(db,'dirty',int(bool(pending)))
    return {'synced':not pending, 'years':results,
            'reason':'workbook_modified' if any(r.get('reason')=='workbook_modified' for r in pending) else '',
            'message':' '.join(r['message'] for r in pending) if pending else 'Yearly Excel workbooks are up to date.'}


def import_pdf(db, path, password=None, resolutions=None, expected_hash=None):
    pdf_hash, rows, excluded, total = extract_pdf(path, secret(db,'pdf_password') if password is None else password)
    if expected_hash and expected_hash != pdf_hash:
        raise AgentError('The PDF changed during review. Upload it again to review the current file.')
    if db.execute('SELECT 1 FROM imports WHERE hash=?',(pdf_hash,)).fetchone():
        result = export_workbook(db) if get_setting(db,'dirty')=='1' else {'synced':True,'message':'This PDF was already imported.'}
        return dict(result, added=0, duplicates=len(rows), excluded=excluded, total=total)
    known = {r['id']:dict(r) for r in db.execute('SELECT * FROM transactions')}
    known.update({r['id']:dict(r) for r in db.execute('SELECT * FROM pdf_matches')})
    deleted = {r[0] for r in db.execute('SELECT id FROM deleted_transactions')}
    fresh=[]
    for row in rows:
        if row['id'] in deleted:
            continue
        prior=known.get(row['id'])
        if prior:
            if any(prior[k]!=row[k] for k in ('date','description','type','reference')) or Decimal(prior['amount']) != Decimal(row['amount']):
                raise AgentError('A transaction reference conflicts with saved data. No transactions were added.')
        else:
            if not valid_text(row['description']):
                raise AgentError('The statement contains an unsupported description. No transactions were saved.')
            fresh.append(row)
            known[row['id']] = row
    # Manual entries have no statement reference. Ask before linking a possible match.
    manual = [dict(r) for r in db.execute("SELECT * FROM transactions WHERE type='Manual expense' AND id NOT IN (SELECT manual_id FROM pdf_matches)")]
    candidates = []
    for row in fresh:
        matches = [m for m in manual if m['date'] == row['date'] and Decimal(m['amount']) == Decimal(row['amount'])]
        if matches:
            candidates.append(dict(row, matches=matches))
    decisions = resolutions or {}
    if any(row['id'] not in decisions for row in candidates):
        return {'needs_review':True, 'pdf_hash':pdf_hash, 'candidates':candidates,
                'message':'Review possible matches to expenses you entered yourself. Nothing from this PDF has been saved yet.'}
    links = {}
    for row in candidates:
        choice = decisions[row['id']]
        if choice == 'keep':
            continue
        if choice not in {m['id'] for m in row['matches']} or choice in links.values():
            raise AgentError('A saved expense can match only one PDF transaction. Review the matches again.')
        links[row['id']] = choice
    with db:
        for row in fresh:
            if row['id'] in links:
                db.execute('INSERT INTO pdf_matches VALUES (:id,:manual_id,:date,:description,:amount,:type,:reference)', dict(row, manual_id=links[row['id']]))
            else:
                db.execute('INSERT INTO transactions VALUES (:id,:date,:description,:amount,:type,:reference,:source,:page)',row)
            month_add(db,row['date'])
        ensure_months(db)
        db.execute('INSERT INTO imports VALUES (?,?,?,?,?,?)',
            (pdf_hash,Path(path).name,datetime.now().isoformat(), 'Local import',len(rows),excluded))
        set_setting(db,'dirty','1')
    result = export_workbook(db)
    return dict(result,added=len(fresh)-len(links),duplicates=len(rows)-len(fresh)+len(links),excluded=excluded,total=total)


def add_expense(db, request):
    value = str(request.get('date', '')).strip()
    try:
        if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', value):
            raise ValueError()
        date.fromisoformat(value)
    except ValueError:
        raise AgentError('Choose a valid expense date.') from None
    description = str(request.get('description', '')).strip()
    if not valid_text(description):
        raise AgentError('Enter a description of 1 to 300 characters.')
    price = str(request.get('price', '')).strip()
    if not re.fullmatch(r'\d{1,9}(?:\.\d{1,2})?', price) or Decimal(price) <= 0:
        raise AgentError('Enter a price greater than zero, with at most two decimal places (for example, 12.50).')
    entry_id = str(request.get('entry_id', ''))
    try:
        entry_id = 'manual:' + str(uuid.UUID(entry_id))
    except ValueError:
        raise AgentError('Please try saving the expense again.') from None
    amount = format(Decimal(price), '.2f')
    if db.execute('SELECT 1 FROM deleted_transactions WHERE id=?', (entry_id,)).fetchone():
        raise AgentError('This expense was deleted. Enter a new expense if you want to add it again.')
    prior = db.execute('SELECT * FROM transactions WHERE id=?', (entry_id,)).fetchone()
    if prior and (prior['date'], prior['description'], prior['amount']) != (value, description, amount):
        raise AgentError('This entry was already saved with different details. Reopen the agent to add a new expense.')
    with db:
        if not prior:
            db.execute('INSERT INTO transactions VALUES (?,?,?,?,?,?,?,?)',
                (entry_id, value, description, amount, 'Manual expense', entry_id, 'Manual entry', 0))
            month_add(db, value)
            set_setting(db, 'dirty', '1')
    try:
        result = export_workbook(db)
    except (AgentError, OSError, subprocess.TimeoutExpired):
        result = {'synced': False, 'message': 'Expense saved. Keep the agent open; it will retry updating Excel automatically.'}
    if result['synced']:
        sheet = calendar.month_name[int(value[5:7])]
        result['message'] = f'Expense saved to {sheet} in Expenses-{value[:4]}.xlsx. Excel total updated.'
    return dict(result, saved=True, month=value[:7])


def delete_expense(db, entry_id):
    if not isinstance(entry_id, str) or not entry_id:
        raise AgentError('Select an expense to delete.')
    row = db.execute('SELECT * FROM transactions WHERE id=?', (entry_id,)).fetchone()
    if row is None:
        return {'deleted':False,'message':'That expense is no longer saved. The list has been refreshed.'}
    # A restorable snapshot must exist before removing the selected record.
    backup_database(db)
    identifiers = [entry_id] + [r[0] for r in db.execute('SELECT id FROM pdf_matches WHERE manual_id=?', (entry_id,))]
    with db:
        db.executemany('INSERT OR IGNORE INTO deleted_transactions VALUES (?,?)',
                       [(value,datetime.now().isoformat()) for value in identifiers])
        db.execute('DELETE FROM transactions WHERE id=?', (entry_id,))
        db.execute('DELETE FROM pdf_matches WHERE manual_id=?', (entry_id,))
        set_setting(db,'dirty','1')
    try:
        result = export_workbook(db)
    except (AgentError,OSError):
        result = {'synced':False,'message':'The expense is deleted from saved data. Refresh Excel to finish updating the workbook.'}
    if result['synced']:
        result['message'] = f'Expense deleted from Expenses-{row["date"][:4]}.xlsx. Its total has been updated.'
    return dict(result,deleted=True,month=row['date'][:7])


def status(db):
    rules=[dict(r) for r in db.execute('SELECT * FROM rules ORDER BY id')]
    records=[dict(r) for r in db.execute('SELECT * FROM transactions ORDER BY date DESC,rowid DESC')]
    ordered=sorted(rules,key=lambda r:(r['mode']!='exact',r['id']))
    for row in records:
        row['display_description']=apply_rules(row['description'],ordered)
    monthly=[]
    for month in db.execute('SELECT * FROM months ORDER BY month DESC'):
        matching=[r for r in records if r['date'].startswith(month['month'])]
        monthly.append(dict(month,sheet=calendar.month_name[int(month['month'][5:7])],count=len(matching),total=str(sum((Decimal(r['amount']) for r in matching),Decimal(0)))))
    years=[]
    exports={r['year']:dict(r) for r in db.execute('SELECT * FROM year_exports')}
    for year in sorted({m['month'][:4] for m in monthly}):
        matching=[r for r in records if r['date'].startswith(year)]
        details=exports.get(year,{})
        years.append({'year':year,'count':len(matching),'total':str(sum((Decimal(r['amount']) for r in matching),Decimal(0))),
                      'workbook':str(workbook_path(year)),'pending':bool(details.get('pending',1)), 'reason':details.get('reason','')})
    preferred=str(date.today().year)
    if years and preferred not in {y['year'] for y in years}:
        preferred=years[-1]['year']
    return {'records':records,'rules':rules,'months':monthly,'count':len(records),
        'total':str(sum((Decimal(r['amount']) for r in records),Decimal(0))),
        'has_password':bool(get_setting(db,'pdf_password')),
        'years':years,'workbook':str(workbook_path(preferred)),
        'dirty':get_setting(db,'dirty')=='1','imports':[dict(r) for r in db.execute('SELECT * FROM imports ORDER BY imported_at DESC')]}


def dispatch(request):
    with closing(connect()) as db:
        action=request.get('action','status')
        result={}
        if action=='save_settings':
            with db:
                password=request.get('pdf_password', '')
                if password:
                    set_setting(db,'pdf_password',protect(password))
            backup_database(db)
        elif action=='forget_password':
            with db:
                db.execute("DELETE FROM settings WHERE key='pdf_password'")
            result={'message':'The saved password was removed. Older local database backups may still contain its Windows-encrypted copy.'}
        elif action=='import':
            result=import_pdf(db,request['path'],request.get('pdf_password'),request.get('resolutions'),request.get('expected_hash'))
        elif action=='add_expense':
            result=add_expense(db,request)
        elif action=='delete_expense':
            result=delete_expense(db,request.get('id'))
        elif action=='add_rule':
            pattern=request.get('pattern','').strip()
            replacement=request.get('replacement','').strip()
            mode=request.get('mode','contains')
            if not valid_text(pattern) or not valid_text(replacement) or mode not in ('exact','contains'):
                raise AgentError('Enter both a merchant phrase and a replacement description (up to 300 characters).')
            with db:
                db.execute('INSERT INTO rules(pattern,replacement,mode) VALUES (?,?,?) ON CONFLICT(pattern,mode) DO UPDATE SET replacement=excluded.replacement',(pattern,replacement,mode))
                set_setting(db,'dirty','1')
            result=export_workbook(db)
        elif action=='delete_rule':
            with db:
                db.execute('DELETE FROM rules WHERE id=?',(int(request['id']),))
                set_setting(db,'dirty','1')
            result=export_workbook(db)
        elif action in ('ensure_month','refresh','rebuild'):
            previous=db.execute('SELECT COUNT(*) FROM months').fetchone()[0]
            with db:
                ensure_months(db)
                if db.execute('SELECT COUNT(*) FROM months').fetchone()[0]!=previous:
                    set_setting(db,'dirty','1')
            result=export_workbook(db,force=action=='rebuild',year=request.get('year') if action=='rebuild' else None)
        elif action!='status':
            raise AgentError('Unknown action.')
        return {'ok':True,'result':result,'status':status(db)}


if __name__=='__main__':
    try:
        # The desktop UI sends UTF-8 through a pipe, independently of the user's
        # Windows code page and Python environment settings.
        sys.stdin.reconfigure(encoding='utf-8')
        sys.stdout.reconfigure(encoding='utf-8')
        if '--ensure-month' in sys.argv:
            request={'action':'ensure_month'}
        else:
            request=json.load(sys.stdin)
        # A separate process can run the monthly task while the desktop app is open.
        # Lock the entire operation, including export, to prevent stale replacements.
        import msvcrt
        DATA.mkdir(parents=True,exist_ok=True)
        with open(DATA/'agent.lock','a+b') as lock:
            lock.seek(0); lock.write(b'0'); lock.flush(); lock.seek(0)
            try:
                msvcrt.locking(lock.fileno(),msvcrt.LK_NBLCK,1)
            except OSError:
                raise AgentError('Another import or monthly update is running. Please try again shortly.') from None
            try:
                response=dispatch(request)
            finally:
                lock.seek(0); msvcrt.locking(lock.fileno(),msvcrt.LK_UNLCK,1)
        print(json.dumps(response,ensure_ascii=True))
    except AgentError as exc:
        print(json.dumps({'ok':False,'error':str(exc)}))
        sys.exit(1)
    except Exception:
        print(json.dumps({'ok':False,'error':'The operation could not finish. Your saved data remains available. Check the PDF, close Excel, and try again.'}))
        sys.exit(1)
