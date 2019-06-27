# -*- coding: utf-8 -*-

import os
import pandas as pd
from io import BytesIO
from flask import Flask, request, render_template, send_file, redirect
from models.optimize import optimize
from models.optimize import validate_file
from models.optimize import validate_model
os.chdir(os.path.dirname(os.path.realpath(__file__)))

app = Flask(__name__)
ALLOWED_EXTENSIONS = ['xlsx', 'xlsb', 'xlsm', 'xls']
MODEL_DIR = './tmp'
DOWNLOAD_DIR = './downloads'

@app.route('/', methods=("POST", "GET"))
@app.route('/index', methods=("POST", "GET"))
def home():
    #click upload
    if 'upload' in request.form:
        if 'file' not in request.files:
            return redirect(request.url)
        else:
            f = request.files['file']
            if f.filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS:
                model_dict = validate_file(f, MODEL_DIR, DOWNLOAD_DIR)
                val_sheet = [x for x, val in model_dict['val_sheet'].items() if val == 0]
                sheet_status = 'OK' if len(val_sheet) == 0 else "ERROR sheets (%s)" % (', '.join(val_sheet))
                if len(val_sheet) > 0:
                    return render_template('error_file.html',
                                           upload_file=model_dict['upload']['filename'],
                                           upload_time=model_dict['upload']['upload_time'],
                                           sheet_status=sheet_status)
                else:
                    model_dict = validate_model(MODEL_DIR, DOWNLOAD_DIR)
                    val_sheet = [x for x, val in model_dict['val_sheet'].items() if val == 0]
                    val_master = [x for x, val in model_dict['val_master'].items() if val == 0]
                    val_feas = [x for x, val in model_dict['val_feas'].items() if val == 0]
                    sheet_status = 'OK' if len(val_sheet) == 0 else "ERROR sheets (%s)" % (', '.join(val_sheet))
                    master_status = 'OK' if len(val_master) == 0 else "ERROR sheets (%s)" % (', '.join(val_master))
                    feas_status = 'OK' if len(val_feas) == 0 else "ERROR (%s)" % (', '.join(val_feas))
                    if len(val_master) > 0:
                        return render_template('error_master.html',
                                               upload_file=model_dict['upload']['filename'],
                                               upload_time=model_dict['upload']['upload_time'],
                                               sheet_status=sheet_status,
                                               master_status=master_status,
                                               feas_status=feas_status)
                    elif len(val_feas) > 0:
                        return render_template('error_feas.html',
                                               upload_file=model_dict['upload']['filename'],
                                               upload_time=model_dict['upload']['upload_time'],
                                               sheet_status=sheet_status,
                                               master_status=master_status,
                                               feas_status=feas_status)
                    else:
                        return render_template('upload.html',
                                               upload_file=model_dict['upload']['filename'],
                                               upload_time=model_dict['upload']['upload_time'],
                                               sheet_status=sheet_status,
                                               master_status=master_status,
                                               feas_status=feas_status)
            else:
                return redirect(request.url)
    #click solve
    elif 'solve' in request.form:
        model_dict = optimize(MODEL_DIR, DOWNLOAD_DIR)
        val_sheet = [x for x, val in model_dict['val_sheet'].items() if val == 0]
        val_master = [x for x, val in model_dict['val_master'].items() if val == 0]
        val_feas = [x for x, val in model_dict['val_feas'].items() if val == 0]
        sheet_status = 'OK' if len(val_sheet) == 0 else "ERROR sheets (%s)" % (', '.join(val_sheet))
        master_status = 'OK' if len(val_master) == 0 else "ERROR sheets (%s)" % (', '.join(val_master))
        feas_status = 'OK' if len(val_feas) == 0 else "ERROR (%s)" % (', '.join(val_feas))
        return render_template('result.html',
                               upload_file=model_dict['upload']['filename'],
                               upload_time=model_dict['upload']['upload_time'],
                               sheet_status=sheet_status,
                               master_status=master_status,
                               feas_status=feas_status,
                               opt_start=model_dict['opt']['start_time'],
                               opt_end=model_dict['opt']['end_time'],
                               opt_status=model_dict['opt']['status'],
                               opt_time="%.2f" % model_dict['opt']['total_time'])
    else:
        return render_template('home.html')

@app.route("/output/<tablename>", methods=['GET', 'POST'])
def output(tablename):
    dtype = {"plant": str, "truck": str, "cust": str, "mat": str}
    df = pd.read_excel(os.path.join(DOWNLOAD_DIR, 'output.xlsx'), sheet_name=tablename, dtype=dtype)
    return render_template('output.html',
                           tablename=tablename,
                           df=[df.to_html(classes='output', index=False)])

@app.route('/test', methods=("POST", "GET"))
def test():
    return render_template('test.html')

@app.route("/download/<filename>", methods=['GET', 'POST'])
def download(filename):
    return send_file(os.path.join(DOWNLOAD_DIR, filename), attachment_filename=filename, as_attachment=True, cache_timeout=0)

@app.route("/output/downloadsheet/<tablename>", methods=['GET', 'POST'])
def downloadsheet(tablename):
    #get table
    dtype = {"plant": str, "truck": str, "cust": str, "mat": str}
    df = pd.read_excel(os.path.join(DOWNLOAD_DIR, 'output.xlsx'), sheet_name=tablename, dtype=dtype)
    #save to bytes
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name=tablename, index=False)
    writer.close()
    output.seek(0)
    filename = tablename + '.xlsx'
    return send_file(output, attachment_filename=filename, as_attachment=True, cache_timeout=0)

if __name__ == '__main__':
    app.run(debug=True)
