# -*- coding: utf-8 -*-
import datetime
from pytz import timezone
import pandas as pd
import numpy as np
from ortools.linear_solver import pywraplp
from io import BytesIO

import mod

p = mod.PathFile()


class Optimize():
    def __init__(self, user):
        self.user = user
        p.setuser(self.user)
        self.stream = False if p.config['app']['run'] == 'local' else True

    @staticmethod
    def makedict(df, id_col, val_col, crossproduct, replace_str=0, null_val=0):
        df_dict = pd.Series(df[val_col].values, index=df[id_col]).to_dict()
        df_dict = dict((x, df_dict[x]) if x in df_dict else (x, null_val) for x in crossproduct)
        try:
            df_dict = dict((x, float(y)) for x, y in df_dict.items())
        except ValueError:
            df_dict = dict((x, replace_str) for x, y in df_dict.items())
        return df_dict

    @staticmethod
    def checkmaster(name, df, col_list, master_list):
        e = 0
        for col, master in zip(col_list, master_list):
            if set(df[col]) > set(master):
                e = e + 1
        return 1 if e == 0 else 0

    @staticmethod
    def write_dict_to_worksheet(x, sheetname, workbook):
        worksheet = workbook.add_worksheet(sheetname)
        row = 0
        for key, val in x.items():
            worksheet.write(row, 0, key)
            worksheet.write(row, 1, val)
            row += 1

    def upload_file(self, file, request):
        self.file = file
        self.request = request
        self.status = {}
        self.status['upload'] = {}
        self.status['upload']['filename'] = self.file.filename
        self.status['upload']['user'] = self.user
        self.status['upload']['upload_time'] = datetime.datetime.now(
            timezone('Asia/Bangkok')).strftime("%Y-%m-%d %H:%M:%S")
        self.status['upload']['upload_ip'] = self.request.remote_addr
        self.input_status = self.status['upload']
        return self.status['upload']

    def validate_sheet(self):
        self.status['val_sheet'] = {}
        sheet_dict = {"cust": ["cust", "custname"],
                      "mat": ["mat", "matname", "lot", "vol", "weight"],
                      "truck": ["truck", "truckname"],
                      "plant": ["plant", "plantname"],
                      "order": ["cust", "mat", "order", "order_priority"],
                      "custmat": ["cust", "mat", "inv_qty", "inv_min", "sales_qty", "price", "cost", "dos_min", "dos_max"],
                      "custtruck": ["cust", "truck", "min_vol", "min_weight", "min_price", "min_cost", "max_vol", "max_weight", "trans_cost", "extraplant_cost", "cost_ratio", "plant_limit"],
                      "custplant": ["cust", "plant", "custplant_avl"],
                      "truckplant": ["truck", "plant", "truckplant_avl"],
                      "matplant": ["plant", "mat", "supply_qty"]}
        sheet_list = [sh for sh, col in sheet_dict.items()]
        dtypes = {"plant": str, "truck": str, "cust": str, "mat": str}
        for sheet in sheet_list:
            try:
                df = pd.read_excel(self.file, sheet_name=sheet, dtype=dtypes)
                self.status['val_sheet'][sheet] = 1 if set(
                    list(df.columns)) == set(sheet_dict[sheet]) else 0
            except Exception:
                self.status['val_sheet'][sheet] = 0

        return self.status['val_sheet']

    def validate_master(self):
        self.status['val_master'] = {}

        # import data
        dtype = {"plant": str, "truck": str, "cust": str, "mat": str}
        self.df_cust = pd.read_excel(self.file, sheet_name='cust', dtype=dtype).replace(
            'nan', np.nan).dropna(subset=['cust'])
        self.df_mat = pd.read_excel(self.file, sheet_name='mat', dtype=dtype).replace(
            'nan', np.nan).dropna(subset=['mat'])
        self.df_truck = pd.read_excel(self.file, sheet_name='truck', dtype=dtype).replace(
            'nan', np.nan).dropna(subset=['truck'])
        self.df_plant = pd.read_excel(self.file, sheet_name='plant', dtype=dtype).replace(
            'nan', np.nan).dropna(subset=['plant'])
        self.df_order = pd.read_excel(self.file, sheet_name='order', dtype=dtype).replace(
            'nan', np.nan).dropna(subset=['cust', 'mat'])
        self.df_custmat = pd.read_excel(self.file, sheet_name='custmat', dtype=dtype).replace(
            'nan', np.nan).dropna(subset=['cust', 'mat'])
        self.df_custtruck = pd.read_excel(self.file, sheet_name='custtruck', dtype=dtype).replace(
            'nan', np.nan).dropna(subset=['cust', 'truck'])
        self.df_custplant = pd.read_excel(self.file, sheet_name='custplant', dtype=dtype).replace(
            'nan', np.nan).dropna(subset=['cust', 'plant'])
        self.df_truckplant = pd.read_excel(self.file, sheet_name='truckplant', dtype=dtype).replace(
            'nan', np.nan).dropna(subset=['truck', 'plant'])
        self.df_matplant = pd.read_excel(self.file, sheet_name='matplant', dtype=dtype).replace(
            'nan', np.nan).dropna(subset=['mat', 'plant'])

        # master data
        self.cust = pd.Series(self.df_cust['custname'].values, index=self.df_cust['cust']).to_dict()
        self.mat = pd.Series(self.df_mat['matname'].values, index=self.df_mat['mat']).to_dict()
        self.truck = pd.Series(self.df_truck['truckname'].values,
                               index=self.df_truck['truck']).to_dict()
        self.plant = pd.Series(self.df_plant['plantname'].values,
                               index=self.df_plant['plant']).to_dict()

        # cross product
        self.custmat = [(i, j) for i in self.cust for j in self.mat]
        self.custtruck = [(i, k) for i in self.cust for k in self.truck]
        self.custplant = [(i, l) for i in self.cust for l in self.plant]
        self.truckplant = [(k, l) for k in self.truck for l in self.plant]
        self.matplant = [(j, l) for j in self.mat for l in self.plant]

        # sheet mat
        self.mat_lot = pd.Series(self.df_mat['lot'].values, index=self.df_mat['mat']).to_dict()
        self.mat_weight = pd.Series(
            self.df_mat['weight'].values, index=self.df_mat['mat']).to_dict()
        self.mat_vol = pd.Series(self.df_mat['vol'].values, index=self.df_mat['mat']).to_dict()

        # sheet order
        self.df_order['id'] = self.df_order.apply(lambda x: (x['cust'], x['mat']), axis=1)
        self.order = self.makedict(self.df_order, 'id', 'order', self.custmat)
        self.order_lot = dict((x[0], int(x[1] / self.mat_lot[x[0][1]]) * self.mat_lot[x[0][1]])
                              for x in self.order.items())  # check for full lot
        self.order_qty = dict((x[0], self.order[x[0]] - x[1])
                              for x in self.order_lot.items())  # check remainder
        self.order_priority = self.makedict(
            self.df_order, 'id', 'order_priority', self.custmat, 99999, 99999)

        # sheet custmat
        self.df_custmat['id'] = self.df_custmat.apply(lambda x: (x['cust'], x['mat']), axis=1)
        self.inv_qty = self.makedict(self.df_custmat, 'id', 'inv_qty', self.custmat, 0, 0)
        self.inv_min = self.makedict(self.df_custmat, 'id', 'inv_min', self.custmat, 0, 0)
        self.sales_qty = self.makedict(
            self.df_custmat, 'id', 'sales_qty', self.custmat, 0.0001, 0.0001)
        self.sales_qty = dict((x, 0.00001) if y <= 0 else (x, y) for x, y in self.sales_qty.items())
        self.custmat_price = self.makedict(self.df_custmat, 'id', 'price', self.custmat, 0, 0)
        self.custmat_cost = self.makedict(self.df_custmat, 'id', 'cost', self.custmat, 0, 0)
        self.dos_min = self.makedict(self.df_custmat, 'id', 'dos_min', self.custmat, 0, 0)
        self.dos_max = self.makedict(self.df_custmat, 'id', 'dos_max', self.custmat, 0, 0)

        # sheet custtruck
        self.df_custtruck['id'] = self.df_custtruck.apply(lambda x: (x['cust'], x['truck']), axis=1)
        self.min_vol = self.makedict(self.df_custtruck, 'id', 'min_vol', self.custtruck, 0, 0)
        self.min_weight = self.makedict(self.df_custtruck, 'id', 'min_weight', self.custtruck, 0, 0)
        self.min_price = self.makedict(self.df_custtruck, 'id', 'min_price', self.custtruck, 0, 0)
        self.min_cost = self.makedict(self.df_custtruck, 'id', 'min_cost', self.custtruck, 0, 0)
        self.max_vol = self.makedict(self.df_custtruck, 'id', 'max_vol', self.custtruck, 9999999, 0)
        self.max_weight = self.makedict(self.df_custtruck, 'id',
                                        'max_weight', self.custtruck, 9999999, 0)
        self.trans_cost = self.makedict(self.df_custtruck, 'id', 'trans_cost', self.custtruck, 0, 0)
        self.extraplant_cost = self.makedict(
            self.df_custtruck, 'id', 'extraplant_cost', self.custtruck, 0, 0)
        self.cost_ratio = self.makedict(self.df_custtruck, 'id', 'cost_ratio', self.custtruck, 1, 1)
        self.plant_limit = self.makedict(
            self.df_custtruck, 'id', 'plant_limit', self.custtruck, 1, 1)

        # sheet custplant
        self.df_custplant['id'] = self.df_custplant.apply(lambda x: (x['cust'], x['plant']), axis=1)
        self.custplant_avl = self.makedict(
            self.df_custplant, 'id', 'custplant_avl', self.custplant, 0, 0)

        # sheet truckplant
        self.df_truckplant['id'] = self.df_truckplant.apply(
            lambda x: (x['truck'], x['plant']), axis=1)
        self.truckplant_avl = self.makedict(
            self.df_truckplant, 'id', 'truckplant_avl', self.truckplant, 0, 0)

        # sheet matplant
        self.df_matplant['id'] = self.df_matplant.apply(lambda x: (x['mat'], x['plant']), axis=1)
        self.supply_qty = self.makedict(self.df_matplant, 'id', 'supply_qty', self.matplant, 0, 0)

        # check master
        self.status['val_master']['cust'] = 1 if len(self.cust) > 0 else 0
        self.status['val_master']['mat'] = 1 if len(self.mat) > 0 else 0
        self.status['val_master']['truck'] = 1 if len(self.truck) > 0 else 0
        self.status['val_master']['plant'] = 1 if len(self.plant) > 0 else 0
        self.status['val_master']['order'] = self.checkmaster(
            'order', self.df_order, ['cust', 'mat'], [self.cust, self.mat])
        self.status['val_master']['custmat'] = self.checkmaster(
            'custmat', self.df_custmat, ['cust', 'mat'], [self.cust, self.mat])
        self.status['val_master']['custtruck'] = self.checkmaster(
            'custtruck', self.df_custtruck, ['cust', 'truck'], [self.cust, self.truck])
        self.status['val_master']['custplant'] = self.checkmaster(
            'custplant', self.df_custplant, ['cust', 'plant'], [self.cust, self.plant])
        self.status['val_master']['truckplant'] = self.checkmaster(
            'truckplant', self.df_truckplant, ['truck', 'plant'], [self.mat, self.truck])
        self.status['val_master']['matplant'] = self.checkmaster(
            'matplant', self.df_matplant, ['mat', 'plant'], [self.mat, self.plant])

        return self.status['val_master']

    def validate_feas(self):
        self.status['val_feas'] = {}

        # write file to BytesIO
        output_file = BytesIO()
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

        # write input information
        self.write_dict_to_worksheet(self.input_status, "input", writer.book)

        # check minimum replenishment by customer minimum qty
        df_req = pd.DataFrame([{'cust': i, 'mat': j} for i in self.cust for j in self.mat])
        df_req['id'] = df_req.apply(lambda x: (x['cust'], x['mat']), axis=1)
        df_req['inv_qty'] = df_req['id'].map(self.inv_qty)
        df_req['inv_min'] = df_req['id'].map(self.inv_min)
        df_req['mat_lot'] = df_req['mat'].map(self.mat_lot)
        df_req['mat_vol'] = df_req['mat'].map(self.mat_vol)
        df_req['mat_weight'] = df_req['mat'].map(self.mat_weight)
        df_req['req_qty'] = df_req.apply(lambda x: np.ceil((x['inv_min'] - x['inv_qty']) / x['mat_lot']) * x['mat_lot']
                                         if x['inv_min'] > x['inv_qty'] else 0, axis=1)
        df_req['req_vol'] = df_req['req_qty'] * df_req['mat_vol']
        df_req['req_weight'] = df_req['req_qty'] * df_req['mat_weight']
        # check available truck by customer
        df_feas_cust = pd.DataFrame([{'cust': i, 'truck': k}
                                     for i in self.cust for k in self.truck])
        df_feas_cust['id'] = df_feas_cust.apply(lambda x: (x['cust'], x['truck']), axis=1)
        df_feas_cust['max_vol'] = df_feas_cust['id'].map(self.max_vol)
        df_feas_cust['max_weight'] = df_feas_cust['id'].map(self.max_weight)
        df_feas_cust = pd.merge(df_feas_cust, self.df_truckplant.groupby(
            ['truck'], as_index=False).agg({"truckplant_avl": "sum"}), on='truck', how='left')
        df_feas_cust = df_feas_cust[df_feas_cust['truckplant_avl'] > 0]
        df_feas_cust = df_feas_cust.groupby(['cust'], as_index=False).agg(
            {"max_vol": "sum", "max_weight": "sum"})
        df_feas_cust = pd.merge(df_feas_cust, df_req.groupby(['cust'], as_index=False).agg(
            {"req_vol": "sum", "req_weight": "sum"}), on='cust', how='right')
        df_feas_cust['feas_vol'] = df_feas_cust.apply(
            lambda x: True if x['max_vol'] >= x['req_vol'] else False, axis=1)
        df_feas_cust['feas_weight'] = df_feas_cust.apply(
            lambda x: True if x['max_weight'] >= x['req_weight'] else False, axis=1)
        df_feas_cust['feas'] = df_feas_cust.apply(lambda x: True if (
            x['feas_vol'] == True) & (x['feas_weight'] == True) else False, axis=1)
        if len(df_feas_cust[df_feas_cust['feas'] == False]) > 0:
            self.status['val_feas']['cust'] = 0
        else:
            self.status['val_feas']['cust'] = 1
        df_feas_cust.to_excel(writer, sheet_name='feas_cust', index=False)

        # check available supply by material
        df_feas_custmat = pd.DataFrame([{'mat': j, 'plant': l}
                                        for j in self.mat for l in self.plant])
        df_feas_custmat['id'] = df_feas_custmat.apply(lambda x: (x['mat'], x['plant']), axis=1)
        df_feas_custmat['supply_qty'] = df_feas_custmat['id'].map(self.supply_qty)
        df_feas_custmat = pd.merge(
            df_feas_custmat, self.df_custplant[self.df_custplant['custplant_avl'] == 1], on='plant', how='inner')
        df_feas_custmat = df_feas_custmat.groupby(
            ['cust', 'mat'], as_index=False).agg({"supply_qty": "sum"})
        df_feas_custmat = pd.merge(df_feas_custmat, df_req[['cust', 'mat', 'req_qty']], on=[
                                   'cust', 'mat'], how='right')
        df_feas_custmat = df_feas_custmat.fillna(0)
        df_feas_custmat['feas_qty'] = df_feas_custmat.apply(
            lambda x: True if x['supply_qty'] >= x['req_qty'] else False, axis=1)
        if len(df_feas_custmat[df_feas_custmat['feas_qty'] == False]) > 0:
            self.status['val_feas']['custmat'] = 0
        else:
            self.status['val_feas']['custmat'] = 1
        df_feas_custmat.to_excel(writer, sheet_name='feas_custmat', index=False)

        # check available supply by truck
        df_feas_custtruck = pd.DataFrame([{'cust': i, 'truck': k, 'plant': l}
                                          for i in self.cust for k in self.truck for l in self.plant])
        df_feas_custtruck['id'] = df_feas_custtruck.apply(lambda x: (x['cust'], x['truck']), axis=1)
        df_feas_custtruck['min_vol'] = df_feas_custtruck['id'].map(self.min_vol)
        df_feas_custtruck['min_weight'] = df_feas_custtruck['id'].map(self.min_weight)
        df_feas_custtruck['min_price'] = df_feas_custtruck['id'].map(self.min_price)
        df_feas_custtruck['min_cost'] = df_feas_custtruck['id'].map(self.min_cost)
        df_feas_custtruck['trans_cost'] = df_feas_custtruck['id'].map(self.trans_cost)
        df_feas_custtruck['cost_ratio'] = df_feas_custtruck['id'].map(self.cost_ratio)
        df_feas_custtruck = pd.merge(df_feas_custtruck, self.df_custplant[self.df_custplant['custplant_avl'] == 1], on=[
                                     'cust', 'plant'], how='left')
        df_feas_custtruck = pd.merge(df_feas_custtruck, self.df_truckplant[self.df_truckplant['truckplant_avl'] == 1], on=[
                                     'truck', 'plant'], how='left')
        df_feas_custtruck = df_feas_custtruck[(df_feas_custtruck['custplant_avl'] == 1) & (
            df_feas_custtruck['truckplant_avl'] == 1)]
        df_feas_custtruck = pd.merge(df_feas_custtruck, self.df_matplant, on='plant', how='left')
        df_feas_custtruck['supply_vol'] = df_feas_custtruck.apply(lambda x: 0 if pd.isnull(x['mat'])
                                                                  else x['supply_qty'] * self.mat_vol[x['mat']], axis=1)
        df_feas_custtruck['supply_weight'] = df_feas_custtruck.apply(lambda x: 0 if pd.isnull(x['mat'])
                                                                     else x['supply_qty'] * self.mat_weight[x['mat']], axis=1)
        df_feas_custtruck['supply_price'] = df_feas_custtruck.apply(lambda x: 0 if pd.isnull(x['mat'])
                                                                    else x['supply_qty'] * self.custmat_price[(x['cust'], x['mat'])], axis=1)
        df_feas_custtruck['supply_cost'] = df_feas_custtruck.apply(lambda x: 0 if pd.isnull(x['mat'])
                                                                   else x['supply_qty'] * self.custmat_cost[(x['cust'], x['mat'])], axis=1)
        df_feas_custtruck = df_feas_custtruck.groupby(['cust', 'truck', 'min_vol', 'min_weight', 'min_price', 'min_cost',
                                                       'trans_cost', 'cost_ratio'], as_index=False).agg({"mat": "count", "supply_vol": "sum", "supply_price": "sum", "supply_cost": "sum"})
        df_feas_custtruck['trans_cost_ratio'] = df_feas_custtruck['supply_cost'] * \
            df_feas_custtruck['cost_ratio']
        df_feas_custtruck['feas_vol'] = df_feas_custtruck.apply(
            lambda x: True if x['supply_vol'] >= x['min_vol'] else False, axis=1)
        df_feas_custtruck['feas_weight'] = df_feas_custtruck.apply(
            lambda x: True if x['supply_vol'] >= x['min_weight'] else False, axis=1)
        df_feas_custtruck['feas_price'] = df_feas_custtruck.apply(
            lambda x: True if x['supply_price'] >= x['min_price'] else False, axis=1)
        df_feas_custtruck['feas_cost'] = df_feas_custtruck.apply(
            lambda x: True if x['supply_cost'] >= x['min_cost'] else False, axis=1)
        df_feas_custtruck['feas_trans_cost'] = df_feas_custtruck.apply(
            lambda x: True if x['trans_cost_ratio'] >= x['trans_cost'] else False, axis=1)
        df_feas_custtruck['feas'] = df_feas_custtruck.apply(lambda x: True if (x['feas_vol'] == True) &
                                                            (x['feas_weight'] == True) &
                                                            (x['feas_price'] == True) &
                                                            (x['feas_cost'] == True) &
                                                            (x['feas_trans_cost'] == True) else False, axis=1)
        if len(df_feas_custtruck[df_feas_custtruck['feas'] == False]) > 0:
            self.status['val_feas']['custtruck'] = 0
        else:
            self.status['val_feas']['custtruck'] = 1
        df_feas_custtruck.to_excel(writer, sheet_name='feas_custtruck', index=False)

        # write file to path
        writer.save()
        output_file.seek(0)
        p.savefile(output_file, p.config['file']['error'])

        return self.status['val_feas']

    def optimize(self):

        self.status['optimize'] = {}

        '''Start optimization'''
        start_time = datetime.datetime.now(timezone('Asia/Bangkok'))
        self.status['optimize']['start_time'] = start_time.strftime("%Y-%m-%d %H:%M:%S")
        solver = pywraplp.Solver('Fulfillment Optimization',
                                 pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING)

        '''Add Decision Variable'''
        # cust mat truck plant
        trans_stk_lot = {}  # transportation to stock by lot
        trans_order_lot = {}  # transportation of order lot
        trans_order_qty = {}  # transportation of order qty
        for i in self.cust:
            for j in self.mat:
                for k in self.truck:
                    for l in self.plant:
                        trans_stk_lot[i, j, k, l] = solver.IntVar(
                            0, solver.infinity(), 'trans_stk_lot[%s,%s,%s,%s]' % (i, j, k, l))
                        trans_order_lot[i, j, k, l] = solver.IntVar(
                            0, solver.infinity(), 'trans_order_lot[%s,%s,%s,%s]' % (i, j, k, l))
                        trans_order_qty[i, j, k, l] = solver.IntVar(
                            0, 1, 'trans_order_qty[%s,%s,%s,%s]' % (i, j, k, l))

        '''Calculated Variable'''
        # cust mat truck plant
        trans_total_qty = {}  # transportation total quantity
        for i in self.cust:
            for j in self.mat:
                for k in self.truck:
                    for l in self.plant:
                        trans_total_qty[i, j, k, l] = solver.NumVar(
                            0, solver.infinity(), 'trans_total_qty[%s,%s,%s,%s]' % (i, j, k, l))

        # cust mat
        order_lot_left = {}  # order lot left
        order_qty_left = {}  # order qty left
        dos_after = {}  # day of supply after replenish
        dos_below_min = {}  # different between minimum day of supply and day of supply after replenish
        dos_above_min = {}  # different between day of supply after replenish and minimum day of supply
        dos_below_max = {}  # different between maximum day of supply and day of supply after replenish
        dos_above_max = {}  # different between day of supply after replenish and maximum day of supply
        for i in self.cust:
            for j in self.mat:
                order_lot_left[i, j] = solver.NumVar(
                    0, solver.infinity(), 'order_lot_left[%s,%s]' % (i, j))
                order_qty_left[i, j] = solver.NumVar(
                    0, solver.infinity(), 'order_qty_left[%s,%s]' % (i, j))
                dos_after[i, j] = solver.NumVar(-solver.infinity(),
                                                solver.infinity(), 'dos_after[%s,%s]' % (i, j))
                dos_below_min[i, j] = solver.NumVar(
                    0, solver.infinity(), 'dos_below_min[%s,%s]' % (i, j))
                dos_above_min[i, j] = solver.NumVar(
                    0, solver.infinity(), 'dos_above_min[%s,%s]' % (i, j))
                dos_below_max[i, j] = solver.NumVar(
                    0, solver.infinity(), 'dos_below_max[%s,%s]' % (i, j))
                dos_above_max[i, j] = solver.NumVar(
                    0, solver.infinity(), 'dos_above_max[%s,%s]' % (i, j))

        # cust truck plant
        plant_visit = {}  # plant visit
        for i in self.cust:
            for k in self.truck:
                for l in self.plant:
                    plant_visit[i, k, l] = solver.IntVar(0, 1, 'plant_visit[%s,%s,%s]' % (i, k, l))

        '''Objective Function'''
        penalty_order_left = 100000
        penalty_below_min = 1000
        penalty_above_min = 1
        penalty_below_max = 10
        penalty_above_max = 100
        # maximize transportation
        solver.Minimize(solver.Sum([dos_below_min[x[0], x[1]] for x in dos_below_min]) * penalty_below_min  # dos after below minimum dos
                        + solver.Sum([dos_above_min[x[0], x[1]] for x in dos_above_min]
                                     ) * penalty_above_min  # dos after above minimum dos
                        + solver.Sum([dos_below_max[x[0], x[1]] for x in dos_below_max]
                                     ) * penalty_below_max  # dos after below maximum dos
                        + solver.Sum([dos_above_max[x[0], x[1]] for x in dos_above_max]
                                     ) * penalty_above_max  # dos after above maximum dos
                        + solver.Sum([order_lot_left[x[0], x[1]] / self.order_priority[x[0], x[1]]
                                      for x in order_lot_left]) * penalty_order_left  # Minimize order lot left
                        + solver.Sum([order_qty_left[x[0], x[1]] / self.order_priority[x[0], x[1]]
                                      for x in order_qty_left]) * penalty_order_left  # Minimize order qty left
                        )

        '''Calculate Variable'''
        # cust mat truck plant
        for i in self.cust:
            for j in self.mat:
                for k in self.truck:
                    for l in self.plant:
                        # transportation total
                        solver.Add(trans_stk_lot[i, j, k, l] * self.mat_lot[j]
                                   + trans_order_lot[i, j, k, l] * self.mat_lot[j]
                                   + trans_order_qty[i, j, k, l] * self.order_qty[(i, j)]
                                   <= trans_total_qty[i, j, k, l])
                        solver.Add(trans_stk_lot[i, j, k, l] * self.mat_lot[j]
                                   + trans_order_lot[i, j, k, l] * self.mat_lot[j]
                                   + trans_order_qty[i, j, k, l] * self.order_qty[(i, j)]
                                   >= trans_total_qty[i, j, k, l])

        # cust mat
        for i in self.cust:
            for j in self.mat:
                # order lot left
                solver.Add(self.order_lot[(i, j)]
                           - solver.Sum([trans_order_lot[x[0], x[1], x[2], x[3]]
                                         for x in trans_order_lot if (x[0] == i and x[1] == j)]) * self.mat_lot[j]
                           <= order_lot_left[(i, j)])
                solver.Add(self.order_lot[(i, j)]
                           - solver.Sum([trans_order_lot[x[0], x[1], x[2], x[3]]
                                         for x in trans_order_lot if (x[0] == i and x[1] == j)]) * self.mat_lot[j]
                           >= order_lot_left[(i, j)])
                # order qty left
                solver.Add(self.order_qty[(i, j)]
                           - solver.Sum([trans_order_qty[x[0], x[1], x[2], x[3]]
                                         for x in trans_order_qty if (x[0] == i and x[1] == j)]) * self.order_qty[i, j]
                           <= order_qty_left[(i, j)])
                solver.Add(self.order_qty[(i, j)]
                           - solver.Sum([trans_order_qty[x[0], x[1], x[2], x[3]]
                                         for x in trans_order_qty if (x[0] == i and x[1] == j)]) * self.order_qty[i, j]
                           >= order_qty_left[(i, j)])
                # dos after replenish
                solver.Add((self.inv_qty[(i, j)]
                            + solver.Sum([trans_stk_lot[x[0], x[1], x[2], x[3]]
                                          for x in trans_stk_lot if x[0] == i and x[1] == j]) * self.mat_lot[j]
                            ) / self.sales_qty[(i, j)]
                           <= dos_after[(i, j)])
                solver.Add((self.inv_qty[(i, j)]
                            + solver.Sum([trans_stk_lot[x[0], x[1], x[2], x[3]]
                                          for x in trans_stk_lot if x[0] == i and x[1] == j]) * self.mat_lot[j]
                            ) / self.sales_qty[(i, j)]
                           >= dos_after[(i, j)])
                # dos different between min and max
                solver.Add(dos_below_min[i, j] >= self.dos_min[(i, j)] - dos_after[i, j])
                solver.Add(dos_above_min[i, j] >= dos_after[i, j] - self.dos_min[(i, j)])
                solver.Add(dos_below_max[i, j] >= self.dos_max[(i, j)] - dos_after[i, j])
                solver.Add(dos_above_max[i, j] >= dos_after[i, j] - self.dos_max[(i, j)])

        # cust truck plant
        for i in self.cust:
            for k in self.truck:
                for l in self.plant:
                    # plant visit
                    solver.Add(solver.Sum([trans_stk_lot[x[0], x[1], x[2], x[3]] for x in trans_stk_lot if (x[0] == i and x[2] == k and x[3] == l)])
                               + solver.Sum([trans_order_lot[x[0], x[1], x[2], x[3]]
                                             for x in trans_order_lot if (x[0] == i and x[2] == k and x[3] == l)])
                               + solver.Sum([trans_order_qty[x[0], x[1], x[2], x[3]]
                                             for x in trans_order_qty if (x[0] == i and x[2] == k and x[3] == l)])
                               >= plant_visit[i, k, l] * 1)
                    solver.Add(solver.Sum([trans_stk_lot[x[0], x[1], x[2], x[3]] for x in trans_stk_lot if (x[0] == i and x[2] == k and x[3] == l)])
                               + solver.Sum([trans_order_lot[x[0], x[1], x[2], x[3]]
                                             for x in trans_order_lot if (x[0] == i and x[2] == k and x[3] == l)])
                               + solver.Sum([trans_order_qty[x[0], x[1], x[2], x[3]]
                                             for x in trans_order_qty if (x[0] == i and x[2] == k and x[3] == l)])
                               <= plant_visit[i, k, l] * 10000)

        '''Add Constraint'''
        # mat plant
        for j in self.mat:
            for l in self.plant:
                # replenish less than available supply
                solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] for x in trans_total_qty if x[1] == j and x[3] == l])
                           <= self.supply_qty[(j, l)])

        # cust mat
        for i in self.cust:
            for j in self.mat:
                # replenish more than minimum inventory
                solver.Add(self.inv_qty[(i, j)]
                           + solver.Sum([trans_stk_lot[x[0], x[1], x[2], x[3]]
                                         for x in trans_stk_lot if x[0] == i and x[1] == j]) * self.mat_lot[j]
                           >= self.inv_min[(i, j)])

        # cust mat truck plant
        for i in self.cust:
            for j in self.mat:
                for k in self.truck:
                    for l in self.plant:
                        # custplant available
                        if self.custplant_avl[(i, l)] <= 0:
                            solver.Add(trans_total_qty[i, j, k, l] <= 0)
                        # truckplant available
                        if self.truckplant_avl[(k, l)] <= 0:
                            solver.Add(trans_total_qty[i, j, k, l] <= 0)

        # cust truck
        for i in self.cust:
            for k in self.truck:
                # truck min vol
                solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * self.mat_vol[x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                           >= self.min_vol[(i, k)])
                # truck min weight
                solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * self.mat_weight[x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                           >= self.min_weight[(i, k)])
                # truck min price
                solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * self.custmat_price[x[0], x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                           >= self.min_price[(i, k)])
                # truck min cost
                solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * self.custmat_cost[x[0], x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                           >= self.min_cost[(i, k)])
                # truck max volume
                solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * self.mat_vol[x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                           <= self.max_vol[(i, k)])
                # truck max weight
                solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * self.mat_weight[x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                           <= self.max_weight[(i, k)])
                # truck max transportation ratio
                solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * self.custmat_cost[x[0], x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                           * self.cost_ratio[(i, k)]
                           >= self.trans_cost[(i, k)] + self.extraplant_cost[(i, k)]
                           * (solver.Sum([plant_visit[x[0], x[1], x[2]] for x in plant_visit if (x[0] == i and x[1] == k)]) - 1))
                # plant limit
                solver.Add(solver.Sum([plant_visit[x[0], x[1], x[2]] for x in plant_visit if (
                    x[0] == i and x[1] == k)]) <= self.plant_limit[(i, k)])

        '''Solve'''
        solver.Solve()
        obj_val = solver.Objective().Value()
        if obj_val > 0:
            self.status['optimize']['status'] = "OPTIMAL"
        else:
            self.status['optimize']['status'] = "INFEASIBLE"

        '''Result'''
        trans_stk_lot_output = dict(
            ((x[0], x[1], x[2], x[3]), trans_stk_lot[x[0], x[1], x[2], x[3]].solution_value()) for x in trans_stk_lot)
        trans_order_lot_output = dict(
            ((x[0], x[1], x[2], x[3]), trans_order_lot[x[0], x[1], x[2], x[3]].solution_value()) for x in trans_order_lot)
        trans_order_qty_output = dict(
            ((x[0], x[1], x[2], x[3]), trans_order_qty[x[0], x[1], x[2], x[3]].solution_value()) for x in trans_order_qty)
        trans_total_qty_output = dict(
            ((x[0], x[1], x[2], x[3]), trans_total_qty[x[0], x[1], x[2], x[3]].solution_value()) for x in trans_total_qty)
        order_lot_left_output = dict(
            ((x[0], x[1]), order_lot_left[x[0], x[1]].solution_value()) for x in order_lot_left)
        order_qty_left_output = dict(
            ((x[0], x[1]), order_qty_left[x[0], x[1]].solution_value()) for x in order_qty_left)
        dos_after_output = dict(
            ((x[0], x[1]), dos_after[x[0], x[1]].solution_value()) for x in dos_after)
        dos_below_min_output = dict(
            ((x[0], x[1]), dos_below_min[x[0], x[1]].solution_value()) for x in dos_below_min)
        dos_above_min_output = dict(
            ((x[0], x[1]), dos_above_min[x[0], x[1]].solution_value()) for x in dos_above_min)
        dos_below_max_output = dict(
            ((x[0], x[1]), dos_below_max[x[0], x[1]].solution_value()) for x in dos_below_max)
        dos_above_max_output = dict(
            ((x[0], x[1]), dos_above_max[x[0], x[1]].solution_value()) for x in dos_above_max)
        plant_visit_output = dict(
            ((x[0], x[1], x[2]), plant_visit[x[0], x[1], x[2]].solution_value()) for x in plant_visit)

        # all transportation
        data = []
        for x in trans_stk_lot:
            data.append({
                'cust': x[0],
                'mat': x[1],
                'truck': x[2],
                'plant': x[3],
                'trans_stk_lot': trans_stk_lot[x[0], x[1], x[2], x[3]].solution_value(),
                'trans_order_lot': trans_order_lot[x[0], x[1], x[2], x[3]].solution_value(),
                'trans_order_qty': trans_order_qty[x[0], x[1], x[2], x[3]].solution_value()
            })
        df_trans = pd.DataFrame(data)
        df_trans['id'] = df_trans.apply(lambda x: (
            x['cust'], x['mat'], x['truck'], x['plant']), axis=1)
        df_trans['custmat'] = df_trans.apply(lambda x: (x['cust'], x['mat']), axis=1)
        df_trans['mat_lot'] = df_trans['mat'].map(self.mat_lot)
        df_trans['mat_vol'] = df_trans['mat'].map(self.mat_vol)
        df_trans['mat_weight'] = df_trans['mat'].map(self.mat_weight)
        df_trans['custmat_price'] = df_trans['custmat'].map(self.custmat_price)
        df_trans['custmat_cost'] = df_trans['custmat'].map(self.custmat_cost)
        df_trans['order'] = df_trans['custmat'].map(self.order)
        df_trans['order_qty'] = df_trans['custmat'].map(self.order_qty)
        df_trans['trans_total_stk'] = (df_trans['trans_stk_lot'] * df_trans['mat_lot'])
        df_trans['trans_total_order'] = ((df_trans['trans_order_lot'] * df_trans['mat_lot'])
                                         + (df_trans['trans_order_qty'] * df_trans['order_qty']))
        df_trans['trans_total_qty'] = ((df_trans['trans_stk_lot'] * df_trans['mat_lot'])
                                       + (df_trans['trans_order_lot'] * df_trans['mat_lot'])
                                       + (df_trans['trans_order_qty'] * df_trans['order_qty']))
        df_trans['trans_total_vol'] = df_trans['trans_total_qty'] * df_trans['mat_vol']
        df_trans['trans_total_weight'] = df_trans['trans_total_qty'] * df_trans['mat_weight']
        df_trans['trans_total_price'] = df_trans['trans_total_qty'] * df_trans['custmat_price']
        df_trans['trans_total_cost'] = df_trans['trans_total_qty'] * df_trans['custmat_cost']
        df_trans = df_trans[df_trans['trans_total_qty'] > 0]
        df_trans['cost_byplant'] = df_trans.groupby(['cust', 'truck', 'plant'])[
            'trans_total_cost'].transform(np.sum)
        df_trans['index'] = df_trans.sort_values(['cost_byplant', 'trans_total_order', 'trans_total_stk'],
                                                 ascending=[False, False, False]).groupby(['cust', 'truck']).cumcount() + 1
        df_trans['id'] = df_trans['cust'] + df_trans['truck'] + df_trans['index'].astype(str)
        df_trans = df_trans[['id', 'cust', 'truck', 'index', 'plant', 'mat', 'order', 'order_qty', 'mat_lot',
                             'trans_order_lot', 'trans_order_qty', 'trans_stk_lot', 'trans_total_order', 'trans_total_stk', 'trans_total_qty',
                             'trans_total_vol', 'trans_total_weight', 'trans_total_price', 'trans_total_cost']]
        df_trans = df_trans.sort_values(by=['cust', 'truck', 'index']).reset_index(drop=True)

        # check by custtruck
        df_custtruck_output = df_trans.groupby(['cust', 'truck'], as_index=False).agg(
            {"trans_total_vol": "sum", "trans_total_weight": "sum", "trans_total_price": "sum", "trans_total_cost": "sum", "plant": "nunique"})
        df_custtruck_output = df_custtruck_output.rename(columns={'plant': 'total_plant'})
        df_custtruck_output = pd.merge(self.df_custtruck.drop(
            'id', axis=1), df_custtruck_output, on=['cust', 'truck'], how='left')
        if len(df_trans) > 0:
            df_custtruck_plant = df_trans[['cust', 'truck', 'plant']].drop_duplicates().groupby(['cust', 'truck'])[
                'plant'].apply(list).reset_index()
            df_custtruck_output = pd.merge(df_custtruck_output, df_custtruck_plant, on=[
                                           'cust', 'truck'], how='left')
        else:
            df_custtruck_output['plant'] = None
        df_custtruck_output['trans_cost_total'] = df_custtruck_output['trans_cost'] + \
            (df_custtruck_output['total_plant']-1) * df_custtruck_output['extraplant_cost']
        df_custtruck_output['trans_total_ratio'] = df_custtruck_output['cost_ratio'] * \
            df_custtruck_output['trans_total_cost']
        df_custtruck_output['id'] = df_custtruck_output['cust'] + df_custtruck_output['truck']
        df_custtruck_output = df_custtruck_output.set_index('id').reset_index()
        df_custtruck_output = df_custtruck_output[['id', 'cust', 'truck', 'min_vol', 'min_weight', 'min_price', 'min_cost', 'max_vol', 'max_weight',
                                                   'trans_cost', 'extraplant_cost', 'cost_ratio', 'trans_total_vol', 'trans_total_weight', 'trans_total_price', 'trans_total_cost',
                                                   'total_plant', 'plant', 'trans_cost_total', 'trans_total_ratio']]
        df_custtruck_output = df_custtruck_output.sort_values(
            by=['cust', 'truck']).reset_index(drop=True)

        # check by custmat
        df_custmat_output = self.df_custmat.drop(['id', 'price', 'cost'], axis=1)
        df_custmat_output['mat_lot'] = df_custmat_output['mat'].map(self.mat_lot)
        df_custmat_output['dos_per_lot'] = df_custmat_output['mat_lot'] / \
            df_custmat_output['sales_qty']
        df_custmat_output['dos_per_unit'] = 1 / df_custmat_output['sales_qty']
        df_custmat_output = pd.merge(df_custmat_output,
                                     df_trans.groupby(['cust', 'mat'], as_index=False).agg(
                                         {"trans_total_stk": "sum"}),
                                     on=['cust', 'mat'], how='left')
        df_custmat_output['trans_total_stk'] = df_custmat_output['trans_total_stk'].apply(
            lambda x: x if pd.notnull(x) else 0)
        df_custmat_output['inv_after'] = df_custmat_output['inv_qty'] + \
            df_custmat_output['trans_total_stk']
        df_custmat_output['dos_before'] = df_custmat_output['inv_qty'] / \
            df_custmat_output['sales_qty']
        df_custmat_output['dos_after'] = df_custmat_output['inv_after'] / \
            df_custmat_output['sales_qty']
        if len(df_trans) > 0:
            df_custmat_plant = df_trans.groupby(['cust', 'mat'])['plant'].apply(list).reset_index()
            df_custmat_output = pd.merge(df_custmat_output, df_custmat_plant, on=[
                                         'cust', 'mat'], how='left')
        else:
            df_custmat_output['plant'] = None
        df_custmat_plantavl = pd.merge(self.df_custplant[self.df_custplant['custplant_avl'] == 1].drop(
            'id', axis=1), self.df_matplant.drop('id', axis=1), on=['plant'], how='left')
        df_custmat_plantavl = df_custmat_plantavl.groupby(['cust', 'mat'])['plant'].apply(
            list).reset_index().rename(columns={'plant': 'plant_avl'})
        df_custmat_output = pd.merge(df_custmat_output, df_custmat_plantavl, on=[
                                     'cust', 'mat'], how='left')
        df_custmat_output['id'] = df_custmat_output['cust'] + df_custmat_output['mat']
        df_custmat_output = df_custmat_output.set_index('id').reset_index()
        df_custmat_output = df_custmat_output.sort_values(['cust', 'trans_total_stk', 'dos_after', 'dos_before'],
                                                          ascending=[True, False, True, True]).reset_index(drop=True)

        # check by order
        df_order_output = pd.merge(self.df_order.drop('id', axis=1),
                                   df_trans.groupby(['cust', 'mat'], as_index=False).agg(
                                       {"trans_total_order": "sum"}),
                                   on=['cust', 'mat'], how='left')
        df_order_output['trans_total_order'] = df_order_output['trans_total_order'].apply(
            lambda x: x if pd.notnull(x) else 0)
        df_order_output['order_left'] = df_order_output['order'] - \
            df_order_output['trans_total_order']
        df_order_output['id'] = df_order_output['cust'] + df_order_output['mat']
        df_order_output = df_order_output.set_index('id').reset_index()
        df_order_output = df_order_output.sort_values(by=['order_priority']).reset_index(drop=True)

        # export output data to excel
        output_file = BytesIO()
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        self.write_dict_to_worksheet(self.input_status, "input", writer.book)
        df_trans.to_excel(writer, sheet_name='trans', index=False)
        df_custtruck_output.to_excel(writer, sheet_name='custtruck', index=False)
        df_custmat_output.to_excel(writer, sheet_name='custmat', index=False)
        df_order_output.to_excel(writer, sheet_name='order', index=False)
        writer.save()
        output_file.seek(0)
        p.savefile(output_file, p.config['file']['output'])

        # summarize time
        end_time = datetime.datetime.now(timezone('Asia/Bangkok'))

        self.status['optimize']['end_time'] = end_time.strftime("%Y-%m-%d %H:%M:%S")
        self.status['optimize']['total_time'] = (end_time - start_time).total_seconds()

        return self.status['optimize']
