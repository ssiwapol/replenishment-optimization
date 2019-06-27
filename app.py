#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import pandas as pd
from io import BytesIO
from flask import Flask, request, render_template, send_file, redirect
from models import optimize
os.chdir(os.path.dirname(os.path.realpath(__file__)))

app = Flask(__name__)
ALLOWED_EXTENSIONS = ['xlsx', 'xlsb', 'xlsm', 'xls']
TMP_DIR = "./tmp"

@app.route('/', methods=("POST", "GET"))
@app.route('/index', methods=("POST", "GET"))
def home():
    #click upload
    if 'upload' in request.form:
        if 'file' not in request.files:
            return redirect(request.url)
        else:
            f = request.files['file']
            if f.filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS: # check extension
                error_summary = {}
                opt = optimize.Optimize(f, request)
                upload_status = opt.upload_file()
                sheet_status = opt.validate_sheet()
                sheet_error = [x for x, val in sheet_status.items() if val == 0]
                error_summary['sheet'] = 'OK' if len(sheet_error) == 0 else "ERROR sheets (%s)" % (', '.join(sheet_error))
                if len(sheet_error) > 0:
                    return render_template('error_sheet.html', status=opt.status, error_summary=error_summary)
                else:
                    master_status = opt.validate_master()
                    master_error = [x for x, val in master_status.items() if val == 0]
                    error_summary['master'] = 'OK' if len(master_error) == 0 else "ERROR sheets (%s)" % (', '.join(master_error))
                    if len(master_error) > 0:
                        return render_template('error_master.html', status=opt.status, error_summary=error_summary)
                    else:
                        feas_status = opt.validate_feas()
                        feas_error = [x for x, val in feas_status.items() if val == 0]
                        error_summary['feas'] = 'OK' if len(feas_error) == 0 else "ERROR (%s)" % (', '.join(feas_error))
                        if request.form.get('pass-feas-val'):
                            opt_status = opt.optimize()
                            return render_template('result.html', status=opt.status, error_summary=error_summary)
                        elif len(feas_error) > 0:
                            return render_template('error_feas.html', status=opt.status, error_summary=error_summary)
                        else:
                            opt_status = opt.optimize()
                            return render_template('result.html', status=opt.status, error_summary=error_summary)
            else:
                return redirect(request.url)
    else:
        return render_template('home.html')

@app.route("/output/<tablename>", methods=['GET', 'POST'])
def output(tablename):
    dtype = {"plant": str, "truck": str, "cust": str, "mat": str}
    df = pd.read_excel(os.path.join(TMP_DIR, 'output.xlsx'), sheet_name=tablename, dtype=dtype)
    col_transform = ['trans_total_vol', 'trans_total_weight', 'trans_total_price', 'trans_total_cost', 'trans_total_ratio',
                     'sales_qty', 'dos_per_lot', 'dos_per_unit', 'dos_before', 'dos_after']
    for x in col_transform:
        if x in list(df.columns):
            df[x] = df[x].apply(lambda x: '%.2f' % x)
        else:
            pass
    return render_template('output.html',
                           tablename=tablename,
                           df=[df.to_html(classes='output', index=False)])

@app.route("/input_template", methods=['GET', 'POST'])
def input_template():
    return send_file('./models/input_template.xlsx', attachment_filename='input_template.xlsx', as_attachment=True, cache_timeout=0)

@app.route("/download/<filename>", methods=['GET', 'POST'])
def download(filename):
    return send_file(os.path.join(TMP_DIR, filename), attachment_filename=filename, as_attachment=True, cache_timeout=0)

@app.route("/output/downloadsheet/<tablename>", methods=['GET', 'POST'])
def downloadsheet(tablename):
    #get table
    dtype = {"plant": str, "truck": str, "cust": str, "mat": str}
    df = pd.read_excel(os.path.join(TMP_DIR, 'output.xlsx'), sheet_name=tablename, dtype=dtype)
    #save to bytes
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name=tablename, index=False)
    writer.close()
    output.seek(0)
    filename = tablename + '.xlsx'
    return send_file(output, attachment_filename=filename, as_attachment=True, cache_timeout=0)

if __name__ == '__main__':
    app.run(port=5000, debug=True)
