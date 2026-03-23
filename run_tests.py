"""Comprehensive test runner for FileForge web API."""
import requests, json
from pathlib import Path

BASE = 'http://localhost:8080'
TESTDIR = Path('test_files')
OUTDIR  = Path('test_output')
OUTDIR.mkdir(exist_ok=True)

ok = fail = skip = 0
results = []

def t(label, tool_id, options=None, file_pairs=None, filename=None):
    global ok, fail, skip
    opts = options or {}

    try:
        if file_pairs:
            flist = file_pairs
        elif filename:
            fp = TESTDIR / filename
            if not fp.exists():
                results.append(('SKIP', label, 'input missing'))
                skip += 1
                return
            flist = [('files', (filename, open(fp, 'rb')))]
        else:
            results.append(('SKIP', label, 'no file'))
            skip += 1
            return

        r = requests.post(f'{BASE}/api/convert',
            data={'tool': tool_id, 'options': json.dumps(opts)},
            files=flist, timeout=90)

        ct = r.headers.get('content-type', '')
        disp = r.headers.get('content-disposition', '')
        is_file_download = bool(disp) or 'attachment' in disp

        if r.ok and (is_file_download or 'json' not in ct):
            results.append(('OK', label, f'{len(r.content):,} bytes'))
            ok += 1
        elif 'application/json' in ct or r.status_code in (400, 500):
            try:
                d = r.json()
                if isinstance(d, dict) and (d.get('success') is False or 'error' in d):
                    results.append(('FAIL', label, d.get('error', str(d))[:90]))
                    fail += 1
                else:
                    results.append(('OK', label, str(d)[:80]))
                    ok += 1
            except Exception:
                results.append(('FAIL', label, r.text[:90]))
                fail += 1
        elif r.ok:
            results.append(('OK', label, f'{len(r.content):,} bytes'))
            ok += 1
        else:
            results.append(('FAIL', label, r.text[:90]))
            fail += 1
    except Exception as e:
        results.append(('ERR', label, str(e)[:90]))
        fail += 1


def fp(name):
    return open(TESTDIR / name, 'rb')


# ── PDF CORE ──────────────────────────────────────────────────────────────
t('merge_pdfs',       'merge_pdfs',       file_pairs=[('files',('a.pdf',fp('test.pdf'))),('files',('b.pdf',fp('test.pdf')))])
t('split_pdf',        'split_pdf',        {'split_every_page': True},           filename='test.pdf')
t('compress_pdf',     'compress_pdf',     {'quality_preset': 'ebook'},           filename='test.pdf')
t('rotate_pdf',       'rotate_pdf',       {'rotation': '90'},                    filename='test.pdf')
t('protect_pdf',      'protect_pdf',      {'user_password': 'test123'},          filename='test.pdf')
protected = OUTDIR / 'protected.pdf'
if not protected.exists():
    r2 = requests.post(f'{BASE}/api/convert',
        data={'tool': 'protect_pdf', 'options': json.dumps({'user_password': 'test123'})},
        files=[('files', ('test.pdf', open(TESTDIR/'test.pdf','rb')))], timeout=30)
    if r2.ok and 'json' not in r2.headers.get('content-type',''):
        protected.write_bytes(r2.content)
t('unlock_pdf',       'unlock_pdf',       {'password': 'test123'},
  file_pairs=[('files',('protected.pdf', open(protected,'rb')))] if protected.exists() else None)
t('add_watermark',    'add_watermark',    {'text': 'DRAFT', 'opacity': 0.3},     filename='test.pdf')
t('add_page_numbers', 'add_page_numbers', {'position': 'bottom-center', 'start_number': 1}, filename='test.pdf')
t('organize_pdf',     'organize_pdf',     {'page_order': '3,1,2'},               filename='test.pdf')
t('repair_pdf',       'repair_pdf',       {},                                     filename='test.pdf')
t('get_pdf_info',     'get_pdf_info',     {},                                     filename='test.pdf')

# ── PDF EXPORT ────────────────────────────────────────────────────────────
t('pdf_to_images',    'pdf_to_images',    {'output_format': 'png', 'dpi': '100'}, filename='test.pdf')
t('pdf_to_text',      'pdf_to_text',      {},                                      filename='test.pdf')
t('pdf_to_word',      'pdf_to_word',      {},                                      filename='test.pdf')
t('pdf_to_excel',     'pdf_to_excel',     {},                                      filename='test.pdf')
t('pdf_to_pptx',      'pdf_to_pptx',      {},                                      filename='test.pdf')
t('pdf_to_html',      'pdf_to_html',      {},                                      filename='test.pdf')
t('pdf_to_markdown',  'pdf_to_markdown',  {},                                      filename='test.pdf')

# ── PDF IMPORT ────────────────────────────────────────────────────────────
t('images_to_pdf',    'images_to_pdf',    {},
  file_pairs=[('files',('a.png',fp('test.png'))),('files',('b.jpg',fp('test.jpg')))])
t('html_to_pdf',      'html_to_pdf',      {},  filename='test.html')

# ── IMAGE ─────────────────────────────────────────────────────────────────
t('png->webp',        'convert_image',    {'output_format': 'webp', 'quality': '80'}, filename='test.png')
t('png->jpg',         'convert_image',    {'output_format': 'jpg',  'quality': '85'}, filename='test.png')
t('jpg->png',         'convert_image',    {'output_format': 'png'},                   filename='test.jpg')
t('png resize',       'convert_image',    {'output_format': 'png', 'resize_width': '200', 'resize_height': '150'}, filename='test.png')
t('get_image_info',   'get_image_info',   {},                                         filename='test.png')
t('image_to_base64',  'image_to_base64',  {},                                         filename='test.png')
gif_path = TESTDIR / 'test.gif'
if not gif_path.exists():
    from PIL import Image
    frames = []
    for col in ['#ff0000','#00ff00','#0000ff']:
        im = Image.new('RGB', (100,100), col)
        frames.append(im)
    frames[0].save(str(gif_path), format='GIF', save_all=True, append_images=frames[1:], duration=100, loop=0)
t('gif_to_frames',    'gif_to_frames',    {'output_format': 'png'},              filename='test.gif')
t('frames_to_gif',    'frames_to_gif',    {'duration_ms': 150},
  file_pairs=[('files',('a.png',fp('test.png'))),('files',('b.jpg',fp('test.jpg')))])

# ── DOCUMENTS ─────────────────────────────────────────────────────────────
t('markdown_to_html', 'markdown_to_html', {}, filename='test.md')
t('markdown_to_pdf',  'markdown_to_pdf',  {}, filename='test.md')
t('markdown_to_docx', 'markdown_to_docx', {}, filename='test.md')
t('markdown_to_txt',  'markdown_to_txt',  {}, filename='test.md')
t('html_to_markdown', 'html_to_markdown', {}, filename='test.html')
t('html_to_txt',      'html_to_txt',      {}, filename='test.html')
t('html_to_docx',     'html_to_docx',     {}, filename='test.html')
t('html_to_xlsx',     'html_to_xlsx',     {}, filename='test.html')
t('txt_to_pdf',       'txt_to_pdf',       {}, filename='test.txt')
t('txt_to_docx',      'txt_to_docx',      {}, filename='test.txt')
t('txt_to_html',      'txt_to_html',      {}, filename='test.txt')

# ── DATA ──────────────────────────────────────────────────────────────────
t('json->yaml',        'convert_data',     {'output_format': 'yaml'}, filename='test.json')
t('json->csv',         'convert_data',     {'output_format': 'csv'},  filename='test.json')
t('yaml->json',        'convert_data',     {'output_format': 'json'}, filename='test.yaml')
t('csv->json',         'convert_data',     {'output_format': 'json'}, filename='test.csv')
t('csv->xlsx',         'convert_data',     {'output_format': 'xlsx'}, filename='test.csv')
t('xml_to_json',       'xml_to_json',      {},                        filename='test.xml')
t('json_to_xml',       'json_to_xml',      {},                        filename='test.json')
t('xml_to_csv',        'xml_to_csv',       {},                        filename='test.xml')
t('ini_to_json',       'ini_to_json',      {},                        filename='test.ini')
t('ini_to_yaml',       'ini_to_yaml',      {},                        filename='test.ini')
t('env_to_json',       'env_to_json',      {},                        filename='test.env')
t('csv_to_markdown',   'csv_to_markdown',  {},                        filename='test.csv')
t('html_table_to_csv', 'html_table_to_csv',{},                        filename='test.html')
t('jsonl_to_json',     'jsonl_to_json',    {},                        filename='test.jsonl')
t('json_to_jsonl',     'json_to_jsonl',    {},                        filename='test.json')
t('csv_to_parquet',    'csv_to_parquet',   {},                        filename='test.csv')

# save parquet for next test
r_par = requests.post(f'{BASE}/api/convert',
    data={'tool': 'csv_to_parquet', 'options': '{}'},
    files=[('files', ('test.csv', open(TESTDIR/'test.csv','rb')))], timeout=30)
if r_par.ok and 'json' not in r_par.headers.get('content-type',''):
    (OUTDIR / 'test.parquet').write_bytes(r_par.content)
t('parquet_to_csv',    'parquet_to_csv',   {},  file_pairs=[('files',('test.parquet', open(OUTDIR/'test.parquet','rb')))] if (OUTDIR/'test.parquet').exists() else None)
t('parquet_to_json',   'parquet_to_json',  {},  file_pairs=[('files',('test.parquet', open(OUTDIR/'test.parquet','rb')))] if (OUTDIR/'test.parquet').exists() else None)

# ── ARCHIVES ──────────────────────────────────────────────────────────────
t('zip_files',  'zip_files',  {'archive_name': 'myzip'},
  file_pairs=[('files',('a.txt',fp('test.txt'))),('files',('b.json',fp('test.json')))])
t('tar_files',  'tar_files',  {'compression': 'gz'},
  file_pairs=[('files',('a.txt',fp('test.txt'))),('files',('b.json',fp('test.json')))])
t('create_7z',  'create_7z',  {},
  file_pairs=[('files',('a.txt',fp('test.txt'))),('files',('b.json',fp('test.json')))])

# save zip for unzip test
r_zip = requests.post(f'{BASE}/api/convert',
    data={'tool': 'zip_files', 'options': json.dumps({'archive_name': 'test_archive'})},
    files=[('files',('a.txt',open(TESTDIR/'test.txt','rb'))),('files',('b.json',open(TESTDIR/'test.json','rb')))],
    timeout=30)
if r_zip.ok and 'json' not in r_zip.headers.get('content-type',''):
    (OUTDIR / 'test_archive.zip').write_bytes(r_zip.content)
t('unzip_files', 'unzip_files', {},
  file_pairs=[('files',('test_archive.zip', open(OUTDIR/'test_archive.zip','rb')))] if (OUTDIR/'test_archive.zip').exists() else None)

# ── PRINT RESULTS ─────────────────────────────────────────────────────────
print(f'\n{"="*60}')
print(f'RESULTS:  {ok} OK   {fail} FAIL   {skip} SKIP')
print(f'{"="*60}\n')

for status, label, detail in results:
    icon = '[OK  ]' if status == 'OK' else ('[SKIP]' if status == 'SKIP' else '[FAIL]')
    print(f'{icon} {label}: {detail}')
