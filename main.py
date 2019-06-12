from requests_html import HTMLSession
import xlrd
import xlwt
import re
import json
import sys
import os

if len(sys.argv) < 2:
    print('usage: python3 main.py SOURCE_FILE_NAME')

filename = sys.argv[1]
filename_dst = os.path.splitext(filename)[0] + '_converted.xls'

wb = xlrd.open_workbook(filename = filename)
wbwt = xlwt.Workbook()

session = HTMLSession()

with open('dict.json') as f:
    saved_gene_ids = json.load(f)

sheet_names = wb.sheet_names()
for sheet_name in sheet_names:
    sheet = wb.sheet_by_name(sheet_name)
    sheet_wt = wbwt.add_sheet(sheet_name, cell_overwrite_ok=True)
    rows, cols = sheet.nrows, sheet.ncols
    sheet_wt.write(0, 0, 'Standard Name')
    for row in range(rows):
        for col in range(cols):
            if col == 0 and row > 0:
                gene_id = sheet.cell(row,0).value
                alt_name = gene_id
                if gene_id in saved_gene_ids:
                    alt_name = saved_gene_ids[gene_id]
                    print('\r{}/{} reuse {}                                         '.format(row+1, rows, gene_id), end =" ")
                else:
                    try:
                        print('\r{}/{} searching {}                                    '.format(row+1, rows, gene_id), end =" ")
                        r = session.get('https://www.yeastgenome.org/search?q={}&is_quick=true'.format(gene_id))
                        script = r.html.find('script', first=True)
                        alt_name = re.search(r'displayName: "(.*?)",', script.text)[1]
                        saved_gene_ids[gene_id] = alt_name
                    except Exception as err:
                        print('error', err)
                        with open('dict.json','w') as f:
                            json.dump(saved_gene_ids, f)
                sheet_wt.write(row, 0, alt_name)
            sheet_wt.write(row, col + 1, sheet.cell(row, col).value)
    print('\n done')
with open('dict.json','w') as f:
    json.dump(saved_gene_ids, f)
wbwt.save(filename_dst)