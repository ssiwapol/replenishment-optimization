# -*- coding: utf-8 -*-
import os
import pandas as pd
from io import BytesIO
from flask import Flask, Response, request, render_template, send_file, redirect, session
from flask_login import LoginManager, UserMixin, login_required, login_user, logout_user, current_user
from models import optimize

import mod

# set environment
os.chdir(os.path.dirname(os.path.realpath(__file__)))
ALLOWED_EXTENSIONS = ['xlsx', 'xlsb', 'xlsm', 'xls']

# start app
app = Flask(__name__)
app.secret_key = os.urandom(12)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"
p = mod.PathFile()
users = p.getuser()


class User(UserMixin):

    def __init__(self, user_id):
        self.id = user_id
        self.name = user_id
        self.password = users[user_id]


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if username in users.keys() and password == users[username]['password']:
            user = User(username)
            login_user(user)
            return redirect(request.args.get("next"))
        else:
            return redirect(request.url)
    else:
        return render_template('login.html')


@app.route("/logout")
@login_required
def logout():
    logout_user()
    return Response('<p>Logged out</p>')


@login_manager.user_loader
def load_user(userid):
    return User(userid)


@app.route('/login', methods=['POST', 'GET'])
def user_login():
    if request.method == "POST":
        req = request.form
        username = req.get("username")
        password = req.get("password")
        if username in users.keys() and password == users[username]['password']:
            session['logged_in'] = True
            session['user'] = username
            print("session username set")
            return home()
        else:
            session['logged_in'] = False
            return redirect(request.url)
    return home()


@app.route('/', methods=("POST", "GET"))
@login_required
def home():
    opt = optimize.Optimize(current_user.name)
    if 'upload' in request.form:
        if 'file' not in request.files:
            return redirect(request.url)
        else:
            f = request.files['file']
            if f.filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS:  # check extension
                error_summary = {}
                upload_status = opt.upload_file(f, request)
                sheet_status = opt.validate_sheet()
                sheet_error = [x for x, val in sheet_status.items() if val == 0]
                error_summary['sheet'] = 'OK' if len(
                    sheet_error) == 0 else "ERROR sheets (%s)" % (', '.join(sheet_error))
                if len(sheet_error) > 0:
                    return render_template('error_sheet.html', status=opt.status, error_summary=error_summary)
                else:
                    master_status = opt.validate_master()
                    master_error = [x for x, val in master_status.items() if val == 0]
                    error_summary['master'] = 'OK' if len(
                        master_error) == 0 else "ERROR sheets (%s)" % (', '.join(master_error))
                    if len(master_error) > 0:
                        return render_template('error_master.html', status=opt.status, error_summary=error_summary)
                    else:
                        feas_status = opt.validate_feas()
                        feas_error = [x for x, val in feas_status.items() if val == 0]
                        error_summary['feas'] = 'OK' if len(
                            feas_error) == 0 else "ERROR (%s)" % (', '.join(feas_error))
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


@app.route("/download/<file>", methods=['GET', 'POST'])
@login_required
def download(file):
    p.setuser(current_user.name)
    if file == "input_template":
        filename = "input_template.xlsx"
        pathfile = os.path.join("./models", filename)
    else:
        filename = p.config['file'][file]
        pathfile = p.loadfile(filename)
    return send_file(pathfile, attachment_filename=filename, as_attachment=True, cache_timeout=0)


@app.route("/output/<tablename>", methods=['GET', 'POST'])
@login_required
def output(tablename):
    p.setuser(current_user.name)
    dtype = {"plant": str, "truck": str, "cust": str, "mat": str}
    df = pd.read_excel(p.loadfile(p.config['file']['output']), sheet_name=tablename, dtype=dtype)
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


@app.route("/output/downloadsheet/<tablename>", methods=['GET', 'POST'])
def downloadsheet(tablename):
    p.setuser(current_user.name)
    # get table
    dtype = {"plant": str, "truck": str, "cust": str, "mat": str}
    df = pd.read_excel(p.loadfile(p.config['file']['output']), sheet_name=tablename, dtype=dtype)
    # save to bytes
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name=tablename, index=False)
    writer.close()
    output.seek(0)
    filename = tablename + '.xlsx'
    return send_file(output, attachment_filename=filename, as_attachment=True, cache_timeout=0)


if __name__ == '__main__':
    app.run(debug=p.config['app']['debug'], host='0.0.0.0', port=8080)
