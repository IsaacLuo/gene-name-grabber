from requests_html import HTMLSession
import xlrd
import xlwt
import re
import json
import sys
import os
import demjson

if len(sys.argv) < 2:
    print('usage: python3 main.py SOURCE_FILE_NAME')
    exit(0)

filename = sys.argv[1]
filename_dst = os.path.splitext(filename)[0] + '_converted.xls'

wb = xlrd.open_workbook(filename = filename)
wbwt = xlwt.Workbook()

session = HTMLSession()

with open('dict.json') as f:
    saved_gene_ids = json.load(f)

with open("dump.json", 'a+') as f_dump:
    sheet_names = wb.sheet_names()
    for sheet_name in sheet_names:
        print("\n\n{}".format(sheet_name))

        sheet = wb.sheet_by_name(sheet_name)
        sheet_wt = wbwt.add_sheet(sheet_name, cell_overwrite_ok=True)
        rows, cols = sheet.nrows, sheet.ncols
        sheet_wt.write(0, 0, 'Standard Name')
        sheet_wt.write(0, cols+1, 'description')
        for row in range(rows):
            for col in range(cols):
                if col == 0 and row > 0:
                    gene_id = sheet.cell(row,0).value
                    alt_name = gene_id
                    if gene_id in saved_gene_ids:
                        alt_name = saved_gene_ids[gene_id]['displayName']
                        description = saved_gene_ids[gene_id]['description']
                        print('\r{}/{} reuse {}                                         '.format(row+1, rows, gene_id), end=" ")
                    else:
                        try:
                            print('\r{}/{} searching {}                                    '.format(row+1, rows, gene_id), end=" ")
                            r = session.get('https://www.yeastgenome.org/search?q={}&is_quick=true'.format(gene_id))
                            script = r.html.find('script', first=True)
                            # alt_name = re.search(r'displayName: "(.*?)",', script.text)[1]
                            js = re.search(r'var bootstrappedData = (.*);', script.text)[1]
                            my_dict = demjson.decode(js)
                            alt_name = my_dict['displayName']
                            description = my_dict['locusData']['description']
                            saved_gene_ids[gene_id] = {"displayName": my_dict["displayName"], "description": my_dict['locusData']['description']}
                            if row%1000==0:
                                with open('dict.json','w') as f:
                                    f.write(demjson.encode(saved_gene_ids,))

                            f_dump.write(demjson.encode(my_dict))
                            f_dump.write('\n')
                        except Exception as err:
                            print('error', err)
                            with open('dict.json','w') as f:
                                f.write(demjson.encode(saved_gene_ids))
                    sheet_wt.write(row, cols+1, description)
                    sheet_wt.write(row, 0, alt_name)
                sheet_wt.write(row, col + 1, sheet.cell(row, col).value)
        print('\n done')
with open('dict.json','w') as f:
    f.write(demjson.encode(saved_gene_ids))
wbwt.save(filename_dst)