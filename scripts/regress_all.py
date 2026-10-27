"""Regression check: run an old (--base) and the current build_edi.py on every Downloads/check edi-* job
and print every EDI.xlsx cell whose text or fill changed.  Usage: python regress_all.py --base OLD.py [--jobs DIR]"""
import argparse, glob, os, subprocess, sys, tempfile
import openpyxl, pymupdf

HERE = os.path.dirname(os.path.abspath(__file__))
CUR = os.path.join(HERE, 'build_edi.py')
BASE = None
DL = os.path.join(os.path.expanduser('~'), 'Downloads')
OUT = os.path.join(tempfile.gettempdir(), 'edi_regress')


def find_inputs(job):
    d = os.path.join(job, 'input') if os.path.isdir(os.path.join(job, 'input')) else job
    files = [f for f in glob.glob(os.path.join(d, '*')) if os.path.isfile(f)]
    xls = [f for f in files if f.lower().endswith(('.xls', '.xlsx')) and 'CNTRS' not in os.path.basename(f).upper()
           and not os.path.basename(f).upper().startswith('EDI')]
    pdfs = sorted([f for f in files if f.lower().endswith('.pdf')], key=lambda p: (len(p), p))
    man = next((f for f in xls if 'MAN' in os.path.basename(f).upper() or 'MNF' in os.path.basename(f).upper()), None) \
        or (xls[0] if len(xls) == 1 else None)
    man_pdf = [p for p in pdfs if 'MANIFEST' in os.path.basename(p).upper()]
    if not man and man_pdf:
        man = man_pdf[0]
    enters = [p for p in pdfs if p != man]
    return man, enters


def merged_enter(job, enters):
    if len(enters) == 1:
        return enters[0]
    out = os.path.join(OUT, os.path.basename(job) + '_ENTER_merged.pdf')
    doc = pymupdf.open()
    for p in enters:
        doc.insert_pdf(pymupdf.open(p))
    doc.save(out)
    return out


def run(script, man, enter, outdir):
    os.makedirs(outdir, exist_ok=True)
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    r = subprocess.run([sys.executable, script, '--manifest', man, '--enter', enter, '--outdir', outdir],
                       capture_output=True, text=True, encoding='utf-8', errors='replace', env=env)
    return r.returncode, r.stdout + r.stderr


def cells(path):
    ws = openpyxl.load_workbook(path).active
    out = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None or (c.fill and c.fill.fgColor and c.fill.fgColor.rgb not in (None, '00000000')):
                out[c.coordinate] = (c.value, c.fill.fgColor.rgb if c.fill and c.fill.fgColor else None)
    hdr = {c.column_letter: c.value for c in ws[4]}
    bl = {r[0].row: r[1].value for r in ws.iter_rows(min_row=5) if len(r) > 1}
    return out, hdr, bl


def summary_line(text):
    return next((l for l in text.splitlines() if l.startswith('B/L ทั้งหมด')), '(no summary)')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--base', required=True, help='the old build_edi.py to compare against')
    ap.add_argument('--jobs', default=DL, help='folder holding the "check edi-*" job folders')
    args = ap.parse_args()
    BASE, DL = os.path.abspath(args.base), args.jobs
    os.makedirs(OUT, exist_ok=True)
    for job in sorted(glob.glob(os.path.join(DL, 'check edi-*'))):
        name = os.path.basename(job)
        man, enters = find_inputs(job)
        if not man or not enters:
            print(f'## {name}: SKIP (manifest={bool(man)}, enter files={len(enters)})')
            continue
        enter = merged_enter(job, enters)
        rb, ob = run(BASE, man, enter, os.path.join(OUT, name, 'base'))
        rc, oc = run(CUR, man, enter, os.path.join(OUT, name, 'cur'))
        print(f'## {name}: base rc={rb} | cur rc={rc}')
        print('   base:', summary_line(ob))
        print('   cur :', summary_line(oc))
        if rb or rc:
            print('   ERROR OUTPUT (cur):', oc[-800:])
            continue
        a, hdr, bls = cells(os.path.join(OUT, name, 'base', 'EDI.xlsx'))
        b, _, _ = cells(os.path.join(OUT, name, 'cur', 'EDI.xlsx'))
        diffs = sorted(set(a) | set(b), key=lambda k: (int(''.join(ch for ch in k if ch.isdigit())), k))
        diffs = [k for k in diffs if a.get(k) != b.get(k)]
        print(f'   cell differences: {len(diffs)}')
        for k in diffs[:40]:
            col = ''.join(ch for ch in k if ch.isalpha()); rown = int(''.join(ch for ch in k if ch.isdigit()))
            va, vb = a.get(k, (None, None)), b.get(k, (None, None))
            fill = '' if va[1] == vb[1] else f'  [fill {str(va[1])[-6:]} -> {str(vb[1])[-6:]}]'
            print(f'     {k} {str(bls.get(rown, ""))[-6:]:>6} {str(hdr.get(col, ""))[:22]:<22}: {str(va[0])[:90]!r} -> {str(vb[0])[:90]!r}{fill}')
