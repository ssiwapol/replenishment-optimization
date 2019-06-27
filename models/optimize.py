# -*- coding: utf-8 -*-

import os
import datetime
import pandas as pd
import numpy as np
import json
from ortools.linear_solver import pywraplp

def convnum(x, replace_str=0):
    try:
        return float(x)
    except ValueError:
        return replace_str

def makedict(df, id_col, val_col, crossproduct, replace_str=0, null_val=0):
    df_dict = pd.Series(df[val_col].values, index=df[id_col]).to_dict()
    df_dict = dict((x, df_dict[x]) if x in df_dict else (x, null_val) for x in crossproduct)
    df_dict = dict((x, convnum(y, replace_str)) for x, y in df_dict.items())
    return df_dict

def checkmaster(name, df, col_list, master_list):
    e = 0
    for col, master in zip(col_list, master_list):
        if set(df[col]) > set(master): e = e + 1
    return 1 if e==0 else 0

def validate_file(file, tmp_dir):
    model_dict = {}
    model_dict['upload'] = {}
    model_dict['val_sheet'] = {}

    '''Upload'''
    model_dict['upload']['filename'] = file.filename
    model_dict['upload']['upload_time'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    '''Validate sheet'''
    # check error and write file
    sheet_dict = json.load(open(os.path.join('./models', "sheet_dict.json")))
    sheet_list = [sh for sh, col in sheet_dict.items()]
    dtypes = {"plant": str, "truck": str, "cust": str, "mat": str}
    writer = pd.ExcelWriter(os.path.join(tmp_dir, 'input.xlsx'), engine='xlsxwriter')
    for sheet in sheet_list:
        try:
            df = pd.read_excel(file, sheet_name=sheet, dtype=dtypes)
            model_dict['val_sheet'][sheet] = 1 if set(list(df.columns)) == set(sheet_dict[sheet]) else 0
            df.to_excel(writer, sheet_name=sheet, index=False)
        except Exception:
            model_dict['val_sheet'][sheet] = 0
    writer.save()

    # Save model_dict to json
    json.dump(model_dict, open(os.path.join(tmp_dir, "model_dict.json"), 'w'))

    return model_dict

def validate_model(tmp_dir):

    model_dict = json.load(open(os.path.join(tmp_dir, "model_dict.json")))
    model_dict['val_master'] = {}
    model_dict['val_feas'] = {}

    '''Validate master'''
    input_path = os.path.join(tmp_dir, 'input.xlsx')
    # import data
    dtype = {"plant": str, "truck": str, "cust": str, "mat": str}
    df_cust = pd.read_excel(input_path, sheet_name='cust', dtype=dtype)
    df_mat = pd.read_excel(input_path, sheet_name='mat', dtype=dtype)
    df_truck = pd.read_excel(input_path, sheet_name='truck', dtype=dtype)
    df_plant = pd.read_excel(input_path, sheet_name='plant', dtype=dtype)
    df_order = pd.read_excel(input_path, sheet_name='order', dtype=dtype)
    df_custmat = pd.read_excel(input_path, sheet_name='custmat', dtype=dtype)
    df_custtruck = pd.read_excel(input_path, sheet_name='custtruck', dtype=dtype)
    df_custplant = pd.read_excel(input_path, sheet_name='custplant', dtype=dtype)
    df_truckplant = pd.read_excel(input_path, sheet_name='truckplant', dtype=dtype)
    df_matplant = pd.read_excel(input_path, sheet_name='matplant', dtype=dtype)
    # master data
    cust = pd.Series(df_cust['custname'].values, index=df_cust['cust']).to_dict()
    mat = pd.Series(df_mat['matname'].values, index=df_mat['mat']).to_dict()
    truck = pd.Series(df_truck['truckname'].values, index=df_truck['truck']).to_dict()
    plant = pd.Series(df_plant['plantname'].values, index=df_plant['plant']).to_dict()
    mat_lot = pd.Series(df_mat['lot'].values, index=df_mat['mat']).to_dict()
    mat_weight = pd.Series(df_mat['weight'].values, index=df_mat['mat']).to_dict()
    mat_vol = pd.Series(df_mat['vol'].values, index=df_mat['mat']).to_dict()
    # sheet custmat
    df_custmat['id'] = df_custmat.apply(lambda x: (x['cust'], x['mat']), axis=1)
    custmat = [(i, j) for i in cust for j in mat]
    inv_qty = makedict(df_custmat, 'id', 'inv_qty', custmat, 0, 0)
    inv_min = makedict(df_custmat, 'id', 'inv_min', custmat, 0, 0)
    sales_qty = makedict(df_custmat, 'id', 'sales_qty', custmat, 0.0001, 0.0001)
    sales_qty = dict((x, 0.00001) if y <= 0 else (x, y) for x, y in sales_qty.items())
    custmat_price = makedict(df_custmat, 'id', 'price', custmat, 0, 0)
    custmat_cost = makedict(df_custmat, 'id', 'cost', custmat, 0, 0)
    # sheet custtruck
    df_custtruck['id'] = df_custtruck.apply(lambda x: (x['cust'], x['truck']), axis=1)
    custtruck = [(i, k) for i in cust for k in truck]
    min_vol = makedict(df_custtruck, 'id', 'min_vol', custtruck, 0, 0)
    min_weight = makedict(df_custtruck, 'id', 'min_weight', custtruck, 0, 0)
    min_price = makedict(df_custtruck, 'id', 'min_price', custtruck, 0, 0)
    min_cost = makedict(df_custtruck, 'id', 'min_cost', custtruck, 0, 0)
    max_vol = makedict(df_custtruck, 'id', 'max_vol', custtruck, 9999999, 0)
    max_weight = makedict(df_custtruck, 'id', 'max_weight', custtruck, 9999999, 0)
    trans_cost = makedict(df_custtruck, 'id', 'trans_cost', custtruck, 0, 0)
    extraplant_cost = makedict(df_custtruck, 'id', 'extraplant_cost', custtruck, 0, 0)
    cost_ratio = makedict(df_custtruck, 'id', 'cost_ratio', custtruck, 1, 1)
    plant_limit = makedict(df_custtruck, 'id', 'plant_limit', custtruck, 1, 1)
    #sheet matplant
    df_matplant['id'] = df_matplant.apply(lambda x: (x['mat'], x['plant']), axis=1)
    matplant = [(j, l) for j in mat for l in plant]
    supply_qty = makedict(df_matplant, 'id', 'supply_qty', matplant, 0, 0)
    # check master
    model_dict['val_master']['cust'] = 1 if len(cust) > 0 else 0
    model_dict['val_master']['mat'] = 1 if len(mat) > 0 else 0
    model_dict['val_master']['truck'] = 1 if len(truck) > 0 else 0
    model_dict['val_master']['plant'] = 1 if len(plant) > 0 else 0
    model_dict['val_master']['order'] = checkmaster('order', df_order, ['cust', 'mat'], [cust, mat])
    model_dict['val_master']['custmat'] = checkmaster('custmat', df_custmat, ['cust', 'mat'], [cust, mat])
    model_dict['val_master']['custtruck'] = checkmaster('custtruck', df_custtruck, ['cust', 'truck'], [cust, truck])
    model_dict['val_master']['custplant'] = checkmaster('custplant', df_custplant, ['cust', 'plant'], [cust, plant])
    model_dict['val_master']['truckplant'] = checkmaster('truckplant', df_truckplant, ['truck', 'plant'], [mat, truck])
    model_dict['val_master']['matplant'] = checkmaster('matplant', df_matplant, ['mat', 'plant'], [mat, plant])

    val_master = [x for x, val in model_dict['val_master'].items() if val == 0]
    if len(val_master) > 0:
        json.dump(model_dict, open(os.path.join(tmp_dir, "model_dict.json"), 'w'))
        return model_dict

    else:
        '''Validate feasible'''
        # check minimum replenishment by customer minimum qty
        writer = pd.ExcelWriter(os.path.join(tmp_dir, 'error.xlsx'), engine='xlsxwriter')
        df_req = pd.DataFrame([{'cust': i, 'mat': j} for i in cust for j in mat])
        df_req['id'] = df_req.apply(lambda x: (x['cust'], x['mat']), axis=1)
        df_req['inv_qty'] = df_req['id'].map(inv_qty)
        df_req['inv_min'] = df_req['id'].map(inv_min)
        df_req['mat_lot'] = df_req['mat'].map(mat_lot)
        df_req['mat_vol'] = df_req['mat'].map(mat_vol)
        df_req['mat_weight'] = df_req['mat'].map(mat_weight)
        df_req['req_qty'] = df_req.apply(lambda x: np.ceil((x['inv_min'] - x['inv_qty']) / x['mat_lot']) * x['mat_lot']
                                         if x['inv_min'] > x['inv_qty'] else 0, axis=1)
        df_req['req_vol'] = df_req['req_qty'] * df_req['mat_vol']
        df_req['req_weight'] = df_req['req_qty'] * df_req['mat_weight']
        # check available truck by customer
        df_feas_cust = pd.DataFrame([{'cust': i, 'truck': k} for i in cust for k in truck])
        df_feas_cust['id'] = df_feas_cust.apply(lambda x: (x['cust'], x['truck']), axis=1)
        df_feas_cust['max_vol'] = df_feas_cust['id'].map(max_vol)
        df_feas_cust['max_weight'] = df_feas_cust['id'].map(max_weight)
        df_feas_cust = pd.merge(df_feas_cust, df_truckplant.groupby(['truck'], as_index=False).agg({"truckplant_avl": "sum"}), on='truck', how='left')
        df_feas_cust = df_feas_cust[df_feas_cust['truckplant_avl'] > 0]
        df_feas_cust = df_feas_cust.groupby(['cust'], as_index=False).agg({"max_vol": "sum", "max_weight": "sum"})
        df_feas_cust = pd.merge(df_feas_cust, df_req.groupby(['cust'], as_index=False).agg({"req_vol": "sum", "req_weight": "sum"}), on='cust', how='right')
        df_feas_cust['feas_vol'] = df_feas_cust.apply(lambda x: True if x['max_vol'] >= x['req_vol'] else False, axis=1)
        df_feas_cust['feas_weight'] = df_feas_cust.apply(lambda x: True if x['max_weight'] >= x['req_weight'] else False, axis=1)
        df_feas_cust['feas'] = df_feas_cust.apply(lambda x: True if (x['feas_vol'] == True) & (x['feas_weight'] == True) else False, axis=1)
        if len(df_feas_cust[df_feas_cust['feas'] == False]) > 0:
            model_dict['val_feas']['cust'] = 0
        else:
            model_dict['val_feas']['cust'] = 1
        df_feas_cust.to_excel(writer, sheet_name='feas_cust', index=False)
        # check available supply by material
        df_feas_custmat = pd.DataFrame([{'mat': j, 'plant': l} for j in mat for l in plant])
        df_feas_custmat['id'] = df_feas_custmat.apply(lambda x: (x['mat'], x['plant']), axis=1)
        df_feas_custmat['supply_qty'] = df_feas_custmat['id'].map(supply_qty)
        df_feas_custmat = pd.merge(df_feas_custmat, df_custplant[df_custplant['custplant_avl'] == 1], on='plant', how='inner')
        df_feas_custmat = df_feas_custmat.groupby(['cust', 'mat'], as_index=False).agg({"supply_qty": "sum"})
        df_feas_custmat = pd.merge(df_feas_custmat, df_req[['cust', 'mat', 'req_qty']], on=['cust', 'mat'], how='right')
        df_feas_custmat = df_feas_custmat.fillna(0)
        df_feas_custmat['feas_qty'] = df_feas_custmat.apply(lambda x: True if x['supply_qty'] >= x['req_qty'] else False, axis=1)
        if len(df_feas_custmat[df_feas_custmat['feas_qty'] == False]) > 0:
            model_dict['val_feas']['custmat'] = 0
        else:
            model_dict['val_feas']['custmat'] = 1
        df_feas_custmat.to_excel(writer, sheet_name='feas_custmat', index=False)
        # check available supply by truck
        df_feas_custtruck = pd.DataFrame([{'cust': i, 'truck': k, 'plant': l} for i in cust for k in truck for l in plant])
        df_feas_custtruck['id'] = df_feas_custtruck.apply(lambda x: (x['cust'], x['truck']), axis=1)
        df_feas_custtruck['min_vol'] = df_feas_custtruck['id'].map(min_vol)
        df_feas_custtruck['min_weight'] = df_feas_custtruck['id'].map(min_weight)
        df_feas_custtruck['min_price'] = df_feas_custtruck['id'].map(min_price)
        df_feas_custtruck['min_cost'] = df_feas_custtruck['id'].map(min_cost)
        df_feas_custtruck['trans_cost'] = df_feas_custtruck['id'].map(trans_cost)
        df_feas_custtruck['cost_ratio'] = df_feas_custtruck['id'].map(cost_ratio)
        df_feas_custtruck = pd.merge(df_feas_custtruck, df_custplant[df_custplant['custplant_avl'] == 1], on=['cust', 'plant'], how='left')
        df_feas_custtruck = pd.merge(df_feas_custtruck, df_truckplant[df_truckplant['truckplant_avl'] == 1], on=['truck', 'plant'], how='left')
        df_feas_custtruck = df_feas_custtruck[(df_feas_custtruck['custplant_avl']==1) & (df_feas_custtruck['truckplant_avl']==1)]
        df_feas_custtruck = pd.merge(df_feas_custtruck, df_matplant, on='plant', how='left')
        df_feas_custtruck['supply_vol'] = df_feas_custtruck.apply(lambda x: 0 if pd.isnull(x['mat'])
                                                                  else x['supply_qty'] * mat_vol[x['mat']], axis=1)
        df_feas_custtruck['supply_weight'] = df_feas_custtruck.apply(lambda x: 0 if pd.isnull(x['mat'])
                                                                     else x['supply_qty'] * mat_weight[x['mat']], axis=1)
        df_feas_custtruck['supply_price'] = df_feas_custtruck.apply(lambda x: 0 if pd.isnull(x['mat'])
                                                                    else x['supply_qty'] * custmat_price[(x['cust'], x['mat'])], axis=1)
        df_feas_custtruck['supply_cost'] = df_feas_custtruck.apply(lambda x: 0 if pd.isnull(x['mat'])
                                                                   else x['supply_qty'] * custmat_cost[(x['cust'], x['mat'])], axis=1)
        df_feas_custtruck = df_feas_custtruck.groupby(['cust', 'truck', 'min_vol', 'min_weight', 'min_price', 'min_cost', 'trans_cost', 'cost_ratio'], as_index=False).agg({"mat": "count", "supply_vol": "sum", "supply_price": "sum", "supply_cost": "sum"})
        df_feas_custtruck['trans_cost_ratio'] = df_feas_custtruck['supply_cost'] * df_feas_custtruck['cost_ratio']
        df_feas_custtruck['feas_vol'] = df_feas_custtruck.apply(lambda x: True if x['supply_vol'] >= x['min_vol'] else False, axis=1)
        df_feas_custtruck['feas_weight'] = df_feas_custtruck.apply(lambda x: True if x['supply_vol'] >= x['min_weight'] else False, axis=1)
        df_feas_custtruck['feas_price'] = df_feas_custtruck.apply(lambda x: True if x['supply_price'] >= x['min_price'] else False, axis=1)
        df_feas_custtruck['feas_cost'] = df_feas_custtruck.apply(lambda x: True if x['supply_cost'] >= x['min_cost'] else False, axis=1)
        df_feas_custtruck['feas_trans_cost'] = df_feas_custtruck.apply(lambda x: True if x['trans_cost_ratio'] >= x['trans_cost'] else False, axis=1)
        df_feas_custtruck['feas'] = df_feas_custtruck.apply(lambda x: True if (x['feas_vol'] == True) &
                                                            (x['feas_weight'] == True) &
                                                            (x['feas_price'] == True) &
                                                            (x['feas_cost'] == True) &
                                                            (x['feas_trans_cost'] == True) else False, axis=1)
        if len(df_feas_custtruck[df_feas_custtruck['feas'] == False]) > 0:
            model_dict['val_feas']['custtruck'] = 0
        else:
            model_dict['val_feas']['custtruck'] = 1
        df_feas_custtruck.to_excel(writer, sheet_name='feas_custtruck', index=False)
        writer.save()

        # Save model_dict to json
        json.dump(model_dict, open(os.path.join(tmp_dir, "model_dict.json"), 'w'))

        return model_dict

def optimize(tmp_dir):

    model_dict = json.load(open(os.path.join(tmp_dir, "model_dict.json")))
    model_dict['opt'] = {}

    input_path = os.path.join(tmp_dir, 'input.xlsx')
    output_path = os.path.join(tmp_dir, 'output.xlsx')

    start_time = datetime.datetime.now()
    model_dict['opt']['start_time'] = start_time.strftime("%Y-%m-%d %H:%M:%S")

    '''Data Preparation'''
    start = datetime.datetime.now()
    # import data
    dtype = {"plant": str, "truck": str, "cust": str, "mat": str}
    df_cust = pd.read_excel(input_path, sheet_name='cust', dtype=dtype)
    df_mat = pd.read_excel(input_path, sheet_name='mat', dtype=dtype)
    df_truck = pd.read_excel(input_path, sheet_name='truck', dtype=dtype)
    df_plant = pd.read_excel(input_path, sheet_name='plant', dtype=dtype)
    df_order = pd.read_excel(input_path, sheet_name='order', dtype=dtype)
    df_custmat = pd.read_excel(input_path, sheet_name='custmat', dtype=dtype)
    df_custtruck = pd.read_excel(input_path, sheet_name='custtruck', dtype=dtype)
    df_custplant = pd.read_excel(input_path, sheet_name='custplant', dtype=dtype)
    df_truckplant = pd.read_excel(input_path, sheet_name='truckplant', dtype=dtype)
    df_matplant = pd.read_excel(input_path, sheet_name='matplant', dtype=dtype)

    # name master
    cust = pd.Series(df_cust['custname'].values, index=df_cust['cust']).to_dict()
    mat = pd.Series(df_mat['matname'].values, index=df_mat['mat']).to_dict()
    truck = pd.Series(df_truck['truckname'].values, index=df_truck['truck']).to_dict()
    plant = pd.Series(df_plant['plantname'].values, index=df_plant['plant']).to_dict()

    # variable
    mat_lot = pd.Series(df_mat['lot'].values, index=df_mat['mat']).to_dict()
    mat_weight = pd.Series(df_mat['weight'].values, index=df_mat['mat']).to_dict()
    mat_vol = pd.Series(df_mat['vol'].values, index=df_mat['mat']).to_dict()

    # cross product
    custmat = [(i, j) for i in cust for j in mat]
    custtruck = [(i, k) for i in cust for k in truck]
    custplant = [(i, l) for i in cust for l in plant]
    truckplant = [(k, l) for k in truck for l in plant]
    matplant = [(j, l) for j in mat for l in plant]

    # sheet order
    df_order['id'] = df_order.apply(lambda x: (x['cust'], x['mat']), axis=1)
    order = makedict(df_order, 'id', 'order', custmat)
    order_lot = dict((x[0], int(x[1] / mat_lot[x[0][1]]) * mat_lot[x[0][1]]) for x in order.items())  # check for full lot
    order_qty = dict((x[0], order[x[0]] - x[1]) for x in order_lot.items())  # check remainder
    order_priority = makedict(df_order, 'id', 'order_priority', custmat, 99999, 99999)

    # sheet custmat
    df_custmat['id'] = df_custmat.apply(lambda x: (x['cust'], x['mat']), axis=1)
    inv_qty = makedict(df_custmat, 'id', 'inv_qty', custmat, 0, 0)
    inv_min = makedict(df_custmat, 'id', 'inv_min', custmat, 0, 0)
    sales_qty = makedict(df_custmat, 'id', 'sales_qty', custmat, 0.0001, 0.0001)
    sales_qty = dict((x, 0.00001) if y <= 0 else (x, y) for x, y in sales_qty.items())
    custmat_price = makedict(df_custmat, 'id', 'price', custmat, 0, 0)
    custmat_cost = makedict(df_custmat, 'id', 'cost', custmat, 0, 0)
    dos_min = makedict(df_custmat, 'id', 'dos_min', custmat, 0, 0)
    dos_max = makedict(df_custmat, 'id', 'dos_max', custmat, 0, 0)

    # sheet custtruck
    df_custtruck['id'] = df_custtruck.apply(lambda x: (x['cust'], x['truck']), axis=1)
    min_vol = makedict(df_custtruck, 'id', 'min_vol', custtruck, 0, 0)
    min_weight = makedict(df_custtruck, 'id', 'min_weight', custtruck, 0, 0)
    min_price = makedict(df_custtruck, 'id', 'min_price', custtruck, 0, 0)
    min_cost = makedict(df_custtruck, 'id', 'min_cost', custtruck, 0, 0)
    max_vol = makedict(df_custtruck, 'id', 'max_vol', custtruck, 9999999, 0)
    max_weight = makedict(df_custtruck, 'id', 'max_weight', custtruck, 9999999, 0)
    trans_cost = makedict(df_custtruck, 'id', 'trans_cost', custtruck, 0, 0)
    extraplant_cost = makedict(df_custtruck, 'id', 'extraplant_cost', custtruck, 0, 0)
    cost_ratio = makedict(df_custtruck, 'id', 'cost_ratio', custtruck, 1, 1)
    plant_limit = makedict(df_custtruck, 'id', 'plant_limit', custtruck, 1, 1)

    # sheet custplant
    df_custplant['id'] = df_custplant.apply(lambda x: (x['cust'], x['plant']), axis=1)
    custplant_avl = makedict(df_custplant, 'id', 'custplant_avl', custplant, 0, 0)

    # sheet truckplant
    df_truckplant['id'] = df_truckplant.apply(lambda x: (x['truck'], x['plant']), axis=1)
    truckplant_avl = makedict(df_truckplant, 'id', 'truckplant_avl', truckplant, 0, 0)

    # sheet matplant
    df_matplant['id'] = df_matplant.apply(lambda x: (x['mat'], x['plant']), axis=1)
    supply_qty = makedict(df_matplant, 'id', 'supply_qty', matplant, 0, 0)

    model_dict['opt']['prep_time'] = (datetime.datetime.now() - start).total_seconds()

    '''Start optimization'''
    start = datetime.datetime.now()
    solver = pywraplp.Solver('Fulfillment Optimization', pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING)

    '''Add Decision Variable'''
    # cust mat truck plant
    trans_stk_lot = {} # transportation to stock by lot
    trans_order_lot = {} # transportation of order lot
    trans_order_qty = {} # transportation of order qty
    for i in cust:
        for j in mat:
            for k in truck:
                for l in plant:
                    trans_stk_lot[i, j, k, l] = solver.IntVar(0, solver.infinity(), 'trans_stk_lot[%s,%s,%s,%s]' % (i, j, k, l))
                    trans_order_lot[i, j, k, l] = solver.IntVar(0, solver.infinity(), 'trans_order_lot[%s,%s,%s,%s]' % (i, j, k, l))
                    trans_order_qty[i, j, k, l] = solver.IntVar(0, 1, 'trans_order_qty[%s,%s,%s,%s]' % (i, j, k, l))

    '''Calculated Variable'''
    # cust mat truck plant
    trans_total_qty = {} # transportation total quantity
    for i in cust:
        for j in mat:
            for k in truck:
                for l in plant:
                    trans_total_qty[i, j, k, l] = solver.NumVar(0, solver.infinity(), 'trans_total_qty[%s,%s,%s,%s]' % (i, j, k, l))

    # cust mat
    order_lot_left = {} # order lot left
    order_qty_left = {} # order qty left
    dos_after = {} # day of supply after replenish
    dos_below_min = {} # different between minimum day of supply and day of supply after replenish
    dos_above_min = {} # different between day of supply after replenish and minimum day of supply
    dos_below_max = {} # different between maximum day of supply and day of supply after replenish
    dos_above_max = {} # different between day of supply after replenish and maximum day of supply
    for i in cust:
        for j in mat:
            order_lot_left[i, j] = solver.NumVar(0, solver.infinity(), 'order_lot_left[%s,%s]' % (i, j))
            order_qty_left[i, j] = solver.NumVar(0, solver.infinity(), 'order_qty_left[%s,%s]' % (i, j))
            dos_after[i, j] = solver.NumVar(-solver.infinity(), solver.infinity(), 'dos_after[%s,%s]' % (i, j))
            dos_below_min[i, j] = solver.NumVar(0, solver.infinity(), 'dos_below_min[%s,%s]' % (i, j))
            dos_above_min[i, j] = solver.NumVar(0, solver.infinity(), 'dos_above_min[%s,%s]' % (i, j))
            dos_below_max[i, j] = solver.NumVar(0, solver.infinity(), 'dos_below_max[%s,%s]' % (i, j))
            dos_above_max[i, j] = solver.NumVar(0, solver.infinity(), 'dos_above_max[%s,%s]' % (i, j))

    # cust truck plant
    plant_visit = {} # plant visit
    for i in cust:
        for k in truck:
            for l in plant:
                plant_visit[i, k, l] = solver.IntVar(0, 1, 'plant_visit[%s,%s,%s]' % (i, k, l))

    '''Objective Function'''
    penalty_order_left = 100000
    penalty_below_min = 1000
    penalty_above_min = 1
    penalty_below_max = 10
    penalty_above_max = 100
    # maximize transportation
    solver.Minimize(solver.Sum([dos_below_min[x[0], x[1]] for x in dos_below_min]) * penalty_below_min # dos after below minimum dos
                    + solver.Sum([dos_above_min[x[0], x[1]] for x in dos_above_min]) * penalty_above_min # dos after above minimum dos
                    + solver.Sum([dos_below_max[x[0], x[1]] for x in dos_below_max]) * penalty_below_max # dos after below maximum dos
                    + solver.Sum([dos_above_max[x[0], x[1]] for x in dos_above_max]) * penalty_above_max # dos after above maximum dos
                    + solver.Sum([order_lot_left[x[0], x[1]] / order_priority[x[0], x[1]] for x in order_lot_left]) * penalty_order_left  # Minimize order lot left
                    + solver.Sum([order_qty_left[x[0], x[1]] / order_priority[x[0], x[1]] for x in order_qty_left]) * penalty_order_left # Minimize order qty left
                    )

    '''Calculate Variable'''
    # cust mat truck plant
    for i in cust:
        for j in mat:
            for k in truck:
                for l in plant:
                    # transportation total
                    solver.Add(trans_stk_lot[i, j, k, l] * mat_lot[j]
                               + trans_order_lot[i, j, k, l] * mat_lot[j]
                               + trans_order_qty[i, j, k, l] * order_qty[(i, j)]
                               <= trans_total_qty[i, j, k, l])
                    solver.Add(trans_stk_lot[i, j, k, l] * mat_lot[j]
                               + trans_order_lot[i, j, k, l] * mat_lot[j]
                               + trans_order_qty[i, j, k, l] * order_qty[(i, j)]
                               >= trans_total_qty[i, j, k, l])

    # cust mat
    for i in cust:
        for j in mat:
            # order lot left
            solver.Add(order_lot[(i, j)]
                       - solver.Sum([trans_order_lot[x[0], x[1], x[2], x[3]] for x in trans_order_lot if (x[0] == i and x[1] == j)]) * mat_lot[j]
                       <= order_lot_left[(i, j)])
            solver.Add(order_lot[(i, j)]
                       - solver.Sum([trans_order_lot[x[0], x[1], x[2], x[3]] for x in trans_order_lot if (x[0] == i and x[1] == j)]) * mat_lot[j]
                       >= order_lot_left[(i, j)])
            # order qty left
            solver.Add(order_qty[(i, j)]
                       - solver.Sum([trans_order_qty[x[0], x[1], x[2], x[3]] for x in trans_order_qty if (x[0] == i and x[1] == j)]) * order_qty[i, j]
                       <= order_qty_left[(i, j)])
            solver.Add(order_qty[(i, j)]
                       - solver.Sum([trans_order_qty[x[0], x[1], x[2], x[3]] for x in trans_order_qty if (x[0] == i and x[1] == j)]) * order_qty[i, j]
                       >= order_qty_left[(i, j)])
            # dos after replenish
            solver.Add((inv_qty[(i, j)]
                        + solver.Sum([trans_stk_lot[x[0], x[1], x[2], x[3]] for x in trans_stk_lot if x[0] == i and x[1] == j]) * mat_lot[j]
                        ) / sales_qty[(i, j)]
                       <= dos_after[(i, j)])
            solver.Add((inv_qty[(i, j)]
                        + solver.Sum([trans_stk_lot[x[0], x[1], x[2], x[3]] for x in trans_stk_lot if x[0] == i and x[1] == j]) * mat_lot[j]
                        ) / sales_qty[(i, j)]
                       >= dos_after[(i, j)])
            # dos different between min and max
            solver.Add(dos_below_min[i, j] >= dos_min[(i, j)] - dos_after[i, j])
            solver.Add(dos_above_min[i, j] >= dos_after[i, j] - dos_min[(i, j)])
            solver.Add(dos_below_max[i, j] >= dos_max[(i, j)] - dos_after[i, j])
            solver.Add(dos_above_max[i, j] >= dos_after[i, j] - dos_max[(i, j)])

    # cust truck plant
    for i in cust:
        for k in truck:
            for l in plant:
                # plant visit
                solver.Add(solver.Sum([trans_stk_lot[x[0], x[1], x[2], x[3]] for x in trans_stk_lot if (x[0] == i and x[2] == k and x[3] == l)])
                           + solver.Sum([trans_order_lot[x[0], x[1], x[2], x[3]] for x in trans_order_lot if (x[0] == i and x[2] == k and x[3] == l)])
                           + solver.Sum([trans_order_qty[x[0], x[1], x[2], x[3]] for x in trans_order_qty if (x[0] == i and x[2] == k and x[3] == l)])
                           >= plant_visit[i, k, l] * 1)
                solver.Add(solver.Sum([trans_stk_lot[x[0], x[1], x[2], x[3]] for x in trans_stk_lot if (x[0] == i and x[2] == k and x[3] == l)])
                           + solver.Sum([trans_order_lot[x[0], x[1], x[2], x[3]] for x in trans_order_lot if (x[0] == i and x[2] == k and x[3] == l)])
                           + solver.Sum([trans_order_qty[x[0], x[1], x[2], x[3]] for x in trans_order_qty if (x[0] == i and x[2] == k and x[3] == l)])
                           <= plant_visit[i, k, l] * 10000)

    '''Add Constraint'''
    # mat plant
    for j in mat:
        for l in plant:
            # replenish less than available supply
            solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] for x in trans_total_qty if x[1] == j and x[3] == l])
                       <= supply_qty[(j, l)])

    # cust mat
    for i in cust:
        for j in mat:
            # replenish more than minimum inventory
            solver.Add(inv_qty[(i, j)]
                       + solver.Sum([trans_stk_lot[x[0], x[1], x[2], x[3]] for x in trans_stk_lot if x[0] == i and x[1] == j]) * mat_lot[j]
                       >= inv_min[(i, j)])

    # cust mat truck plant
    for i in cust:
        for j in mat:
            for k in truck:
                for l in plant:
                    # custplant available
                    if custplant_avl[(i, l)] <= 0:
                        solver.Add(trans_total_qty[i, j, k, l] <= 0)
                    # truckplant available
                    if truckplant_avl[(k, l)] <= 0:
                        solver.Add(trans_total_qty[i, j, k, l] <= 0)

    # cust truck
    for i in cust:
        for k in truck:
            # truck min vol
            solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * mat_vol[x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                       >= min_vol[(i, k)])
            # truck min weight
            solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * mat_weight[x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                       >= min_weight[(i, k)])
            # truck min price
            solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * custmat_price[x[0], x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                       >= min_price[(i, k)])
            # truck min cost
            solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * custmat_cost[x[0], x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                       >= min_cost[(i, k)])
            # truck max volume
            solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * mat_vol[x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                       <= max_vol[(i, k)])
            # truck max weight
            solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * mat_weight[x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                       <= max_weight[(i, k)])
            # truck max transportation ratio
            solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * custmat_cost[x[0], x[1]] for x in trans_total_qty if x[0] == i and x[2] == k])
                       * cost_ratio[(i, k)]
                       >= trans_cost[(i, k)] + extraplant_cost[(i, k)]
                       * (solver.Sum([plant_visit[x[0], x[1], x[2]] for x in plant_visit if (x[0] == i and x[1] == k)]) - 1))
            # plant limit
            solver.Add(solver.Sum([plant_visit[x[0], x[1], x[2]] for x in plant_visit if (x[0] == i and x[1] == k)]) <= plant_limit[(i, k)])

    '''Solve'''
    solver.Solve()
    obj_val = solver.Objective().Value()
    if obj_val > 0:
        model_dict['opt']['status'] = "OPTIMAL"
    else:
        model_dict['opt']['status'] = "INFEASIBLE"

    model_dict['opt']['opt_time'] = (datetime.datetime.now() - start).total_seconds()

    '''Result'''
    trans_stk_lot_output = dict(((x[0], x[1], x[2], x[3]), trans_stk_lot[x[0], x[1], x[2], x[3]].solution_value()) for x in trans_stk_lot)
    trans_order_lot_output = dict(((x[0], x[1], x[2], x[3]), trans_order_lot[x[0], x[1], x[2], x[3]].solution_value()) for x in trans_order_lot)
    trans_order_qty_output = dict(((x[0], x[1], x[2], x[3]), trans_order_qty[x[0], x[1], x[2], x[3]].solution_value()) for x in trans_order_qty)
    trans_total_qty_output = dict(((x[0], x[1], x[2], x[3]), trans_total_qty[x[0], x[1], x[2], x[3]].solution_value()) for x in trans_total_qty)
    order_lot_left_output = dict(((x[0], x[1]), order_lot_left[x[0], x[1]].solution_value()) for x in order_lot_left)
    order_qty_left_output = dict(((x[0], x[1]), order_qty_left[x[0], x[1]].solution_value()) for x in order_qty_left)
    dos_after_output = dict(((x[0], x[1]), dos_after[x[0], x[1]].solution_value()) for x in dos_after)
    dos_below_min_output = dict(((x[0], x[1]), dos_below_min[x[0], x[1]].solution_value()) for x in dos_below_min)
    dos_above_min_output = dict(((x[0], x[1]), dos_above_min[x[0], x[1]].solution_value()) for x in dos_above_min)
    dos_below_max_output = dict(((x[0], x[1]), dos_below_max[x[0], x[1]].solution_value()) for x in dos_below_max)
    dos_above_max_output = dict(((x[0], x[1]), dos_above_max[x[0], x[1]].solution_value()) for x in dos_above_max)
    plant_visit_output = dict(((x[0], x[1], x[2]), plant_visit[x[0], x[1], x[2]].solution_value()) for x in plant_visit)

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
    if len(df_trans) > 0:
        df_trans['id'] = df_trans.apply(lambda x: (x['cust'], x['mat'], x['truck'], x['plant']), axis=1)
        df_trans['custmat'] = df_trans.apply(lambda x: (x['cust'], x['mat']), axis=1)
        df_trans['mat_lot'] = df_trans['mat'].map(mat_lot)
        df_trans['mat_vol'] = df_trans['mat'].map(mat_vol)
        df_trans['mat_weight'] = df_trans['mat'].map(mat_weight)
        df_trans['custmat_price'] = df_trans['custmat'].map(custmat_price)
        df_trans['custmat_cost'] = df_trans['custmat'].map(custmat_cost)
        df_trans['order'] = df_trans['custmat'].map(order)
        df_trans['order_qty'] = df_trans['custmat'].map(order_qty)
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
        df_trans['cost_byplant'] = df_trans.groupby(['cust', 'truck', 'plant'])['trans_total_cost'].transform(np.sum)
        df_trans['index'] = df_trans.sort_values(['cost_byplant', 'trans_total_order', 'trans_total_stk'],
                                                 ascending=[False, False, False]).groupby(['cust', 'truck']).cumcount() + 1
        df_trans['id'] = df_trans['cust'] + df_trans['truck'] + df_trans['index'].astype(str)
        df_trans = df_trans[['id', 'cust', 'truck', 'index', 'plant', 'mat', 'order', 'order_qty', 'mat_lot',
                             'trans_order_lot', 'trans_order_qty', 'trans_stk_lot', 'trans_total_order', 'trans_total_stk', 'trans_total_qty',
                             'trans_total_vol', 'trans_total_weight', 'trans_total_price', 'trans_total_cost']]
        df_trans = df_trans.sort_values(by=['cust', 'truck', 'index']).reset_index(drop=True)

    # check by custtruck
    df_custtruck_output = df_trans.groupby(['cust', 'truck'], as_index=False).agg({"trans_total_vol": "sum", "trans_total_weight": "sum", "trans_total_price": "sum", "trans_total_cost": "sum", "plant": "nunique"})
    df_custtruck_output = df_custtruck_output.rename(columns={'plant': 'total_plant'})
    df_custtruck_output = pd.merge(df_custtruck_output, df_custtruck.drop('id', axis=1), on=['cust', 'truck'], how='left')
    df_custtruck_output = pd.merge(df_custtruck_output,
                                   df_trans[['cust', 'truck', 'plant']].drop_duplicates().groupby(['cust', 'truck'])['plant'].apply(list).reset_index(),
                                   on=['cust', 'truck'], how='left')
    df_custtruck_output['trans_cost_total'] = df_custtruck_output['trans_cost'] + (df_custtruck_output['total_plant']-1) * df_custtruck_output['extraplant_cost']
    df_custtruck_output['trans_total_ratio'] = df_custtruck_output['cost_ratio'] * df_custtruck_output['trans_total_cost']
    df_custtruck_output['id'] = df_custtruck_output['cust'] + df_custtruck_output['truck']
    df_custtruck_output = df_custtruck_output.set_index('id').reset_index()
    df_custtruck_output = df_custtruck_output[['id', 'cust', 'truck', 'min_vol', 'min_price', 'min_cost', 'max_vol', 'max_weight',
                                               'trans_cost', 'extraplant_cost', 'cost_ratio', 'trans_total_vol', 'trans_total_weight', 'trans_total_price', 'trans_total_cost',
                                               'total_plant', 'plant', 'trans_cost_total', 'trans_total_ratio']]
    df_custtruck_output = df_custtruck_output.sort_values(by=['cust', 'truck']).reset_index(drop=True)

    # check by custmat
    df_custmat_output = df_custmat.drop(['id', 'price', 'cost'], axis=1)
    df_custmat_output['mat_lot'] = df_custmat_output['mat'].map(mat_lot)
    df_custmat_output['dos_per_lot'] = df_custmat_output['mat_lot'] / df_custmat_output['sales_qty']
    df_custmat_output['dos_per_unit'] = 1 / df_custmat_output['sales_qty']
    df_custmat_output = pd.merge(df_custmat_output,
                                 df_trans.groupby(['cust', 'mat'], as_index=False).agg({"trans_total_stk": "sum"}),
                                 on=['cust', 'mat'], how='left')
    df_custmat_output['trans_total_stk'] = df_custmat_output['trans_total_stk'].apply(lambda x: x if pd.notnull(x) else 0)
    df_custmat_output['inv_after'] = df_custmat_output['inv_qty'] + df_custmat_output['trans_total_stk']
    df_custmat_output['dos_before'] = df_custmat_output['inv_qty'] / df_custmat_output['sales_qty']
    df_custmat_output['dos_after'] = df_custmat_output['inv_after'] / df_custmat_output['sales_qty']
    df_custmat_output = pd.merge(df_custmat_output, df_trans.groupby(['cust', 'mat'])['plant'].apply(list).reset_index(), on=['cust', 'mat'], how='left')
    df_custmat_avl = pd.merge(df_custplant[df_custplant['custplant_avl'] == 1].drop('id', axis=1), df_matplant.drop('id', axis=1), on=['plant'], how='left')
    df_custmat_avl = df_custmat_avl.groupby(['cust', 'mat'])['plant'].apply(list).reset_index().rename(columns={'plant': 'plant_avl'})
    df_custmat_output = pd.merge(df_custmat_output, df_custmat_avl, on=['cust', 'mat'], how='left')
    df_custmat_output['id'] = df_custmat_output['cust'] + df_custmat_output['mat']
    df_custmat_output = df_custmat_output.set_index('id').reset_index()
    df_custmat_output = df_custmat_output.sort_values(['cust', 'trans_total_stk', 'dos_after', 'dos_before'],
                                                      ascending=[True, False, True, True]).reset_index(drop=True)

    # check by order
    df_order_output = pd.merge(df_order.drop('id', axis=1),
                               df_trans.groupby(['cust', 'mat'], as_index=False).agg({"trans_total_order": "sum"}),
                               on=['cust', 'mat'], how='left')
    df_order_output['trans_total_order'] = df_order_output['trans_total_order'].apply(lambda x: x if pd.notnull(x) else 0)
    df_order_output['order_left'] = df_order_output['order'] - df_order_output['trans_total_order']
    df_order_output['id'] = df_order_output['cust'] + df_order_output['mat']
    df_order_output = df_order_output.set_index('id').reset_index()
    df_order_output = df_order_output.sort_values(by=['order_priority']).reset_index(drop=True)

    # export output data to excel
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    df_trans.to_excel(writer, sheet_name='trans', index=False)
    df_custtruck_output.to_excel(writer, sheet_name='custtruck', index=False)
    df_custmat_output.to_excel(writer, sheet_name='custmat', index=False)
    df_order_output.to_excel(writer, sheet_name='order', index=False)
    writer.save()

    # summarize time
    end_time = datetime.datetime.now()
    model_dict['opt']['end_time'] = end_time.strftime("%Y-%m-%d %H:%M:%S")
    model_dict['opt']['total_time'] = (end_time - start_time).total_seconds()

    # Save model_dict to json
    json.dump(model_dict, open(os.path.join(tmp_dir, "model_dict.json"), 'w'))

    return model_dict

if __name__== "__main__":
    validate_model('./test')
    optimize('./test')
