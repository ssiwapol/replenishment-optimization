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
    df_dict = dict((x, convnum(y)) for x, y in df_dict.items())
    return df_dict

def checkmaster(name, df, col_list, master_list):
    e = 0
    for col, master in zip(col_list, master_list):
        if set(df[col]) > set(master): e = e + 1
    return 1 if e==0 else 0

def validate_file(file, model_dir, download_dir):
    model_dict = {}
    model_dict['upload'] = {}
    model_dict['val_sheet'] = {}

    '''Upload'''
    model_dict['upload']['filename'] = file.filename
    model_dict['upload']['upload_time'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    '''Validate sheet'''
    # check error and write file
    sheet_dict = json.load(open(os.path.join(model_dir, "sheet_dict.json")))
    sheet_list = [sh for sh, col in sheet_dict.items()]
    dtypes = {"plant": str, "truck": str, "cust": str, "mat": str}
    writer = pd.ExcelWriter(os.path.join(model_dir, 'input.xlsx'), engine='xlsxwriter')
    for sheet in sheet_list:
        try:
            df = pd.read_excel(file, sheet_name=sheet, dtype=dtypes)
            model_dict['val_sheet'][sheet] = 1 if set(list(df.columns)) == set(sheet_dict[sheet]) else 0
            df.to_excel(writer, sheet_name=sheet, index=False)
        except Exception:
            model_dict['val_sheet'][sheet] = 0
    writer.save()

    # Save model_dict to json
    json.dump(model_dict, open(os.path.join(model_dir, "model_dict.json"), 'w'))

    return model_dict

def validate_model(model_dir, download_dir):

    model_dict = json.load(open(os.path.join(model_dir, "model_dict.json")))
    model_dict['val_master'] = {}
    model_dict['val_feas'] = {}

    '''Validate master'''
    input_path = os.path.join(model_dir, 'input.xlsx')
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
    df_mattruck = pd.read_excel(input_path, sheet_name='mattruck', dtype=dtype)
    df_matplant = pd.read_excel(input_path, sheet_name='matplant', dtype=dtype)
    # master data
    cust = pd.Series(df_cust['custname'].values, index=df_cust['cust']).to_dict()
    mat = pd.Series(df_mat['matname'].values, index=df_mat['mat']).to_dict()
    truck = pd.Series(df_truck['truckname'].values, index=df_truck['truck']).to_dict()
    plant = pd.Series(df_plant['plantname'].values, index=df_plant['plant']).to_dict()
    mat_lot = pd.Series(df_mat['lot'].values, index=df_mat['mat']).to_dict()
    mat_weight = pd.Series(df_mat['weight'].values, index=df_mat['mat']).to_dict()
    mat_vol = pd.Series(df_mat['vol'].values, index=df_mat['mat']).to_dict()
    df_custmat['id'] = df_custmat.apply(lambda x: (x['cust'], x['mat']), axis=1)
    custmat = [(i, j) for i in cust for j in mat]
    custmat_price = makedict(df_custmat, 'id', 'price', custmat, 0, 0)
    custmat_cost = makedict(df_custmat, 'id', 'cost', custmat, 0, 0)
    # check master
    model_dict['val_master']['cust'] = 1 if len(cust) > 0 else 0
    model_dict['val_master']['mat'] = 1 if len(mat) > 0 else 0
    model_dict['val_master']['truck'] = 1 if len(truck) > 0 else 0
    model_dict['val_master']['plant'] = 1 if len(plant) > 0 else 0
    model_dict['val_master']['order'] = checkmaster('order', df_order, ['cust', 'mat'], [cust, mat])
    model_dict['val_master']['custmat'] = checkmaster('custmat', df_custmat, ['cust', 'mat'], [cust, mat])
    model_dict['val_master']['custtruck'] = checkmaster('custtruck', df_custtruck, ['cust', 'truck'], [cust, truck])
    model_dict['val_master']['custplant'] = checkmaster('custplant', df_custplant, ['cust', 'plant'], [cust, plant])
    model_dict['val_master']['mattruck'] = checkmaster('mattruck', df_mattruck, ['mat', 'truck'], [mat, truck])
    model_dict['val_master']['matplant'] = checkmaster('matplant', df_matplant, ['mat', 'plant'], [mat, plant])

    val_master = [x for x, val in model_dict['val_master'].items() if val == 0]
    if len(val_master) > 0:
        json.dump(model_dict, open(os.path.join(model_dir, "model_dict.json"), 'w'))
        return model_dict

    else:
        '''Validate feasible'''
        # check minimum replenishment by customer minimum qty
        writer = pd.ExcelWriter(os.path.join(download_dir, 'error.xlsx'), engine='xlsxwriter')
        df_req = df_custmat.copy()
        df_req['mat_lot'] = df_req['mat'].map(mat_lot)
        df_req['mat_vol'] = df_req['mat'].map(mat_vol)
        df_req['mat_weight'] = df_req['mat'].map(mat_weight)
        df_req['req_qty'] = df_req.apply(
            lambda x: np.ceil((x['inv_min'] - x['inv_qty']) / x['mat_lot']) * x['mat_lot'] if x['inv_min'] > x[
                'inv_qty'] else 0, axis=1)
        df_req['req_vol'] = df_req['req_qty'] * df_req['mat_vol']
        df_req['req_weight'] = df_req['req_qty'] * df_req['mat_weight']
        # check available truck by customer
        df_bycust = df_custtruck.groupby(['cust'], as_index=False).agg({"max_vol": "sum", "max_weight": "sum"})
        df_bycust['max_vol'] = df_bycust['max_vol'].apply(lambda x: convnum(x, 9999999))
        df_bycust['max_weight'] = df_bycust['max_weight'].apply(lambda x: convnum(x, 9999999))
        df_bycust = pd.merge(df_bycust,
                             df_req.groupby(['cust'], as_index=False).agg({"req_vol": "sum", "req_weight": "sum"}),
                             on='cust', how='right')
        df_bycust = df_bycust.fillna(0)
        df_bycust['feas_vol'] = df_bycust.apply(lambda x: True if x['max_vol'] >= x['req_vol'] else False, axis=1)
        df_bycust['feas_weight'] = df_bycust.apply(lambda x: True if x['max_weight'] >= x['req_weight'] else False, axis=1)
        df_bycust['feas'] = df_bycust.apply(
            lambda x: True if (x['feas_vol'] == True) & (x['feas_weight'] == True) else False, axis=1)
        df_bycust_infeas = df_bycust[df_bycust['feas'] == False].reset_index(drop=True)
        if len(df_bycust_infeas) > 0:
            model_dict['val_feas']['bycust'] = 0
        else:
            model_dict['val_feas']['bycust'] = 1
        df_bycust_infeas.to_excel(writer, sheet_name='cust_infeas', index=False)
        # check available supply by material
        df_bymat = pd.merge(df_matplant, df_custplant[df_custplant['plant_avl'] == 1], on='plant', how='inner')
        df_bymat = df_bymat.groupby(['cust', 'mat'], as_index=False).agg({"supply_qty": "sum"})
        df_bymat = pd.merge(df_bymat, df_req[['cust', 'mat', 'req_qty']], on=['cust', 'mat'], how='right')
        df_bymat = df_bymat.fillna(0)
        df_bymat['feas_qty'] = df_bymat.apply(lambda x: True if x['supply_qty'] >= x['req_qty'] else False, axis=1)
        df_bymat_infeas = df_bymat[df_bymat['feas_qty'] == False].reset_index(drop=True)
        if len(df_bymat_infeas) > 0:
            model_dict['val_feas']['bymat'] = 0
        else:
            model_dict['val_feas']['bymat'] = 1
        df_bymat_infeas.to_excel(writer, sheet_name='mat_infeas', index=False)
        # check available supply by truck
        df_bytruck = pd.merge(df_custtruck, df_custplant[df_custplant['plant_avl'] == 1], on='cust', how='left')
        df_bytruck = pd.merge(df_bytruck, df_mattruck[df_mattruck['truck_avl'] == 1], on='truck', how='left')
        df_bytruck = pd.merge(df_bytruck, df_matplant, on=['mat', 'plant'], how='left')
        df_bytruck['supply_vol'] = df_bytruck.apply(
            lambda x: 0 if pd.isnull(x['mat']) else x['supply_qty'] * mat_vol[x['mat']], axis=1)
        df_bytruck['supply_weight'] = df_bytruck.apply(
            lambda x: 0 if pd.isnull(x['mat']) else x['supply_qty'] * mat_weight[x['mat']], axis=1)
        df_bytruck['supply_price'] = df_bytruck.apply(
            lambda x: 0 if pd.isnull(x['mat']) else x['supply_qty'] * custmat_price[(x['cust'], x['mat'])], axis=1)
        df_bytruck['supply_cost'] = df_bytruck.apply(
            lambda x: 0 if pd.isnull(x['mat']) else x['supply_qty'] * custmat_cost[(x['cust'], x['mat'])], axis=1)
        df_bytruck = df_bytruck.groupby(['cust', 'truck', 'min_vol', 'min_weight', 'min_price', 'min_cost', 'trans_cost', 'cost_ratio'],
                                        as_index=False).agg(
            {"mat": "count", "supply_vol": "sum", "supply_price": "sum", "supply_cost": "sum"})
        df_bytruck['trans_cost_ratio'] = df_bytruck['supply_cost'] * df_bytruck['cost_ratio']
        df_bytruck['feas_vol'] = df_bytruck.apply(lambda x: True if x['supply_vol'] >= x['min_vol'] else False, axis=1)
        df_bytruck['feas_weight'] = df_bytruck.apply(lambda x: True if x['supply_vol'] >= x['min_weight'] else False, axis=1)
        df_bytruck['feas_price'] = df_bytruck.apply(lambda x: True if x['supply_price'] >= x['min_price'] else False,
                                                    axis=1)
        df_bytruck['feas_cost'] = df_bytruck.apply(lambda x: True if x['supply_cost'] >= x['min_cost'] else False, axis=1)
        df_bytruck['feas_trans_cost'] = df_bytruck.apply(
            lambda x: True if x['trans_cost_ratio'] >= x['trans_cost'] else False, axis=1)
        df_bytruck['feas'] = df_bytruck.apply(
            lambda x: True if (x['feas_vol'] == True) & (x['feas_weight'] == True) & (x['feas_price'] == True) &
                    (x['feas_cost'] == True) & (x['feas_trans_cost'] == True) else False, axis=1)
        df_bytruck_infeas = df_bytruck[df_bytruck['feas'] == False]
        if len(df_bytruck_infeas) > 0:
            model_dict['val_feas']['bytruck'] = 0
        else:
            model_dict['val_feas']['bytruck'] = 1
        df_bytruck_infeas.to_excel(writer, sheet_name='truck_infeas', index=False)
        writer.save()

        # Save model_dict to json
        json.dump(model_dict, open(os.path.join(model_dir, "model_dict.json"), 'w'))

        return model_dict

def optimize(model_dir, download_dir):

    model_dict = json.load(open(os.path.join(model_dir, "model_dict.json")))
    model_dict['opt'] = {}

    input_path = os.path.join(model_dir, 'input.xlsx')
    output_path = os.path.join(download_dir, 'output.xlsx')

    start_time = datetime.datetime.now()
    model_dict['opt']['start_time'] = start_time.strftime("%Y-%m-%d %H:%M:%S")
    order_left_penalty = 1000

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
    df_mattruck = pd.read_excel(input_path, sheet_name='mattruck', dtype=dtype)
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
    mattruck = [(j, k) for j in mat for k in truck]
    matplant = [(j, l) for j in mat for l in plant]

    # sheet order
    df_order['id'] = df_order.apply(lambda x: (x['cust'], x['mat']), axis=1)
    order = makedict(df_order, 'id', 'order', custmat)
    order_lot = dict(
        (x[0], int(x[1] / mat_lot[x[0][1]]) * mat_lot[x[0][1]]) for x in order.items())  # check for full lot
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

    # sheet custtruck
    df_custtruck['id'] = df_custtruck.apply(lambda x: (x['cust'], x['truck']), axis=1)
    min_vol = makedict(df_custtruck, 'id', 'min_vol', custtruck, 0, 0)
    min_weight = makedict(df_custtruck, 'id', 'min_weight', custtruck, 0, 0)
    min_price = makedict(df_custtruck, 'id', 'min_price', custtruck, 0, 0)
    min_cost = makedict(df_custtruck, 'id', 'min_cost', custtruck, 0, 0)
    max_vol = makedict(df_custtruck, 'id', 'max_vol', custtruck, 9999999, 0)
    max_weight = makedict(df_custtruck, 'id', 'max_weight', custtruck, 9999999, 0)
    trans_cost = makedict(df_custtruck, 'id', 'trans_cost', custtruck, 0, 0)
    cost_ratio = makedict(df_custtruck, 'id', 'cost_ratio', custtruck, 0, 0)
    plant_limit = makedict(df_custtruck, 'id', 'plant_limit', custtruck, 1, 1)

    # sheet custplant
    df_custplant['id'] = df_custplant.apply(lambda x: (x['cust'], x['plant']), axis=1)
    plant_avl = makedict(df_custplant, 'id', 'plant_avl', custplant, 0, 0)

    # sheet mattruck
    df_mattruck['id'] = df_mattruck.apply(lambda x: (x['mat'], x['truck']), axis=1)
    truck_avl = makedict(df_mattruck, 'id', 'truck_avl', mattruck, 0, 0)

    # sheet matplant
    df_matplant['id'] = df_matplant.apply(lambda x: (x['mat'], x['plant']), axis=1)
    supply_qty = makedict(df_matplant, 'id', 'supply_qty', matplant, 0, 0)

    model_dict['opt']['prep_time'] = (datetime.datetime.now() - start).total_seconds()

    '''Start optimization'''
    start = datetime.datetime.now()
    solver = pywraplp.Solver('Fulfillment Optimization', pywraplp.Solver.CBC_MIXED_INTEGER_PROGRAMMING)

    '''Add Decision Variable'''
    # transportation to stock by lot
    trans_stk_lot = {}
    for i in cust:
        for j in mat:
            for k in truck:
                for l in plant:
                    trans_stk_lot[i, j, k, l] = solver.IntVar(0, solver.infinity(),
                                                              'trans_stk_lot[%s,%s,%s,%s]' % (i, j, k, l))

    # transportation of order lot
    trans_order_lot = {}
    for i in cust:
        for j in mat:
            for k in truck:
                for l in plant:
                    trans_order_lot[i, j, k, l] = solver.IntVar(0, solver.infinity(),
                                                                'trans_order_lot[%s,%s,%s,%s]' % (i, j, k, l))

    # transportation of order qty
    trans_order_qty = {}
    for i in cust:
        for j in mat:
            for k in truck:
                for l in plant:
                    trans_order_qty[i, j, k, l] = solver.IntVar(0, 1, 'trans_order_qty[%s,%s,%s,%s]' % (i, j, k, l))

    '''Calculated Variable'''
    trans_total_qty = {}
    for i in cust:
        for j in mat:
            for k in truck:
                for l in plant:
                    trans_total_qty[i, j, k, l] = solver.NumVar(0, solver.infinity(),
                                                                'trans_total_qty[%s,%s,%s,%s]' % (i, j, k, l))

    # order lot left
    order_lot_left = {}
    for i in cust:
        for j in mat:
            order_lot_left[i, j] = solver.NumVar(0, solver.infinity(), 'order_lot_left[%s,%s]' % (i, j))

    # order qty left
    order_qty_left = {}
    for i in cust:
        for j in mat:
            order_qty_left[i, j] = solver.NumVar(0, solver.infinity(), 'order_qty_left[%s,%s]' % (i, j))

    # day of supply
    dos = {}
    for i in cust:
        for j in mat:
            dos[i, j] = solver.NumVar(-solver.infinity(), solver.infinity(), 'dos[%s,%s]' % (i, j))

    # plant visit
    plant_visit = {}
    for i in cust:
        for k in truck:
            for l in plant:
                plant_visit[i, k, l] = solver.IntVar(0, 1, 'plant_visit[%s,%s,%s]' % (i, k, l))

    '''Objective Function'''
    # maximize transportation
    solver.Minimize(solver.Sum([dos[x[0], x[1]] for x in dos])  # Minimize dos
                    + solver.Sum([order_lot_left[x[0], x[1]] / order_priority[x[0], x[1]] for x in
                                  order_lot_left]) * order_left_penalty  # Minimize order lot left
                    + solver.Sum(
        [order_qty_left[x[0], x[1]] / order_priority[x[0], x[1]] for x in order_qty_left]) * order_left_penalty
                    # Minimize order qty left
                    )

    '''Calculate Variable'''
    # trans total
    for i in cust:
        for j in mat:
            for k in truck:
                for l in plant:
                    solver.Add(trans_stk_lot[i, j, k, l] * mat_lot[j] \
                               + trans_order_lot[i, j, k, l] * mat_lot[j] \
                               + trans_order_qty[i, j, k, l] * order_qty[(i, j)] \
                               <= trans_total_qty[i, j, k, l])
                    solver.Add(trans_stk_lot[i, j, k, l] * mat_lot[j] \
                               + trans_order_lot[i, j, k, l] * mat_lot[j] \
                               + trans_order_qty[i, j, k, l] * order_qty[(i, j)] \
                               >= trans_total_qty[i, j, k, l])

    # order lot left
    for i in cust:
        for j in mat:
            solver.Add(order_lot[(i, j)] \
                       - solver.Sum(
                [trans_order_lot[x[0], x[1], x[2], x[3]] for x in trans_order_lot if (x[0] == i and x[1] == j)]) *
                       mat_lot[j] \
                       <= order_lot_left[(i, j)])
            solver.Add(order_lot[(i, j)] \
                       - solver.Sum(
                [trans_order_lot[x[0], x[1], x[2], x[3]] for x in trans_order_lot if (x[0] == i and x[1] == j)]) *
                       mat_lot[j] \
                       >= order_lot_left[(i, j)])

    # order qty left
    for i in cust:
        for j in mat:
            solver.Add(order_qty[(i, j)] \
                       - solver.Sum(
                [trans_order_qty[x[0], x[1], x[2], x[3]] for x in trans_order_qty if (x[0] == i and x[1] == j)]) *
                       order_qty[i, j] \
                       <= order_qty_left[(i, j)])
            solver.Add(order_qty[(i, j)] \
                       - solver.Sum(
                [trans_order_qty[x[0], x[1], x[2], x[3]] for x in trans_order_qty if (x[0] == i and x[1] == j)]) *
                       order_qty[i, j] \
                       >= order_qty_left[(i, j)])

    # dos
    for i in cust:
        for j in mat:
            solver.Add((inv_qty[(i, j)] \
                        + solver.Sum(
                        [trans_stk_lot[x[0], x[1], x[2], x[3]] for x in trans_stk_lot if x[0] == i and x[1] == j]) *
                        mat_lot[j] \
                        ) / sales_qty[(i, j)] \
                       <= dos[(i, j)])
            solver.Add((inv_qty[(i, j)] \
                        + solver.Sum(
                        [trans_stk_lot[x[0], x[1], x[2], x[3]] for x in trans_stk_lot if x[0] == i and x[1] == j]) *
                        mat_lot[j] \
                        ) / sales_qty[(i, j)] \
                       >= dos[(i, j)])

    # plant visit
    for i in cust:
        for k in truck:
            for l in plant:
                solver.Add(solver.Sum([trans_stk_lot[x[0], x[1], x[2], x[3]] for x in trans_stk_lot if
                                       (x[0] == i and x[2] == k and x[3] == l)]) \
                           + solver.Sum([trans_order_lot[x[0], x[1], x[2], x[3]] for x in trans_order_lot if
                                         (x[0] == i and x[2] == k and x[3] == l)]) \
                           + solver.Sum([trans_order_qty[x[0], x[1], x[2], x[3]] for x in trans_order_qty if
                                         (x[0] == i and x[2] == k and x[3] == l)]) \
                           >= plant_visit[i, k, l] * 1)
                solver.Add(solver.Sum([trans_stk_lot[x[0], x[1], x[2], x[3]] for x in trans_stk_lot if
                                       (x[0] == i and x[2] == k and x[3] == l)]) \
                           + solver.Sum([trans_order_lot[x[0], x[1], x[2], x[3]] for x in trans_order_lot if
                                         (x[0] == i and x[2] == k and x[3] == l)]) \
                           + solver.Sum([trans_order_qty[x[0], x[1], x[2], x[3]] for x in trans_order_qty if
                                         (x[0] == i and x[2] == k and x[3] == l)]) \
                           <= plant_visit[i, k, l] * 10000)

    '''Add Constraint'''
    # replenish less than available supply
    for j in mat:
        for l in plant:
            solver.Add(solver.Sum(
                [trans_total_qty[x[0], x[1], x[2], x[3]] for x in trans_total_qty if x[1] == j and x[3] == l]) \
                       <= supply_qty[(j, l)])

    # replenish less than maximum inventory
    for i in cust:
        for j in mat:
            solver.Add(inv_qty[(i, j)]
                       + solver.Sum(
                [trans_stk_lot[x[0], x[1], x[2], x[3]] for x in trans_stk_lot if x[0] == i and x[1] == j]) * mat_lot[j]
                       >= inv_min[(i, j)])

    # plant available
    for i in cust:
        for j in mat:
            for k in truck:
                for l in plant:
                    if plant_avl[(i, l)] <= 0:
                        solver.Add(trans_total_qty[i, j, k, l] <= 0)

    # truck available
    for i in cust:
        for j in mat:
            for k in truck:
                for l in plant:
                    if truck_avl[(j, k)] <= 0:
                        solver.Add(trans_total_qty[i, j, k, l] <= 0)

    # truck min vol
    for i in cust:
        for k in truck:
            solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * mat_vol[x[1]] for x in trans_total_qty if
                                   x[0] == i and x[2] == k]) \
                       >= min_vol[(i, k)])

    # truck min weight
    for i in cust:
        for k in truck:
            solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * mat_weight[x[1]] for x in trans_total_qty if
                                   x[0] == i and x[2] == k]) \
                       >= min_weight[(i, k)])

    # truck min price
    for i in cust:
        for k in truck:
            solver.Add(solver.Sum(
                [trans_total_qty[x[0], x[1], x[2], x[3]] * custmat_price[x[0], x[1]] for x in trans_total_qty if
                 x[0] == i and x[2] == k]) \
                       >= min_price[(i, k)])

    # truck min cost
    for i in cust:
        for k in truck:
            solver.Add(solver.Sum(
                [trans_total_qty[x[0], x[1], x[2], x[3]] * custmat_cost[x[0], x[1]] for x in trans_total_qty if
                 x[0] == i and x[2] == k]) \
                       >= min_cost[(i, k)])

    # truck max volume
    for i in cust:
        for k in truck:
            solver.Add(solver.Sum([trans_total_qty[x[0], x[1], x[2], x[3]] * mat_vol[x[1]] for x in trans_total_qty if
                                   x[0] == i and x[2] == k]) \
                       <= max_vol[(i, k)])

    # truck max weight
    for i in cust:
        for k in truck:
            solver.Add(solver.Sum(
                [trans_total_qty[x[0], x[1], x[2], x[3]] * mat_weight[x[1]] for x in trans_total_qty if
                 x[0] == i and x[2] == k]) \
                       <= max_weight[(i, k)])

    # truck max transportation ratio
    for i in cust:
        for k in truck:
            solver.Add(solver.Sum(
                [trans_total_qty[x[0], x[1], x[2], x[3]] * custmat_price[x[0], x[1]] for x in trans_total_qty if
                 x[0] == i and x[2] == k]) \
                       * cost_ratio[(i, k)] \
                       >= trans_cost[(i, k)])

    # plant limit
    for i in cust:
        for k in truck:
            solver.Add(solver.Sum(
                [plant_visit[x[0], x[1], x[2]] for x in plant_visit if (x[0] == i and x[1] == k)]) <= plant_limit[(i, k)])

    '''Solve'''
    solver.Solve()
    obj_val = solver.Objective().Value()
    if obj_val > 0:
        model_dict['opt']['status'] = "OPTIMAL"
    else:
        model_dict['opt']['status'] = "INFEASIBLE"

    model_dict['opt']['opt_time'] = (datetime.datetime.now() - start).total_seconds()

    '''Result'''
    trans_stk_lot_output = dict(
        ((x[0], x[1], x[2], x[3]), trans_stk_lot[x[0], x[1], x[2], x[3]].solution_value()) for x in trans_stk_lot)
    trans_order_lot_output = dict(
        ((x[0], x[1], x[2], x[3]), trans_order_lot[x[0], x[1], x[2], x[3]].solution_value()) for x in trans_order_lot)
    trans_order_qty_output = dict(
        ((x[0], x[1], x[2], x[3]), trans_order_qty[x[0], x[1], x[2], x[3]].solution_value()) for x in trans_order_qty)
    trans_total_qty_output = dict(
        ((x[0], x[1], x[2], x[3]), trans_total_qty[x[0], x[1], x[2], x[3]].solution_value()) for x in trans_total_qty)
    order_lot_left_output = dict(((x[0], x[1]), order_lot_left[x[0], x[1]].solution_value()) for x in order_lot_left)
    order_qty_left_output = dict(((x[0], x[1]), order_qty_left[x[0], x[1]].solution_value()) for x in order_qty_left)
    dos_output = dict(((x[0], x[1]), dos[x[0], x[1]].solution_value()) for x in dos)
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
        df_trans['trans_total_order'] = (df_trans['trans_order_lot'] * df_trans['mat_lot']) + (
                    df_trans['trans_order_qty'] * df_trans['order_qty'])
        df_trans['trans_total_qty'] = (df_trans['trans_stk_lot'] * df_trans['mat_lot']) + (
                    df_trans['trans_order_lot'] * df_trans['mat_lot']) + (
                                                  df_trans['trans_order_qty'] * df_trans['order_qty'])
        df_trans['trans_total_vol'] = df_trans['trans_total_qty'] * df_trans['mat_vol']
        df_trans['trans_total_weight'] = df_trans['trans_total_qty'] * df_trans['mat_weight']
        df_trans['trans_total_price'] = df_trans['trans_total_qty'] * df_trans['custmat_price']
        df_trans['trans_total_cost'] = df_trans['trans_total_qty'] * df_trans['custmat_cost']
        df_trans = df_trans[df_trans['trans_total_qty'] > 0]
        df_trans['cost_byplant'] = df_trans.groupby(['cust', 'truck', 'plant'])['trans_total_cost'].transform(np.sum)
        df_trans['index'] = df_trans.sort_values(['cost_byplant', 'trans_total_order', 'trans_total_stk'],
                                                 ascending=[False, False, False]).groupby(
            ['cust', 'truck']).cumcount() + 1
        df_trans['id'] = df_trans['cust'] + df_trans['truck'] + df_trans['index'].astype(str)
        df_trans = df_trans[
            ['id', 'cust', 'truck', 'index', 'plant', 'mat', 'order', 'order_qty', 'mat_lot', 'trans_order_lot',
             'trans_order_qty', 'trans_stk_lot',
             'trans_total_order', 'trans_total_stk', 'trans_total_qty', 'trans_total_vol', 'trans_total_weight',
             'trans_total_price', 'trans_total_cost']]
        df_trans = df_trans.sort_values(by=['cust', 'truck', 'index']).reset_index(drop=True)

    # check by custtruck
    df_custtruck_output = df_trans.groupby(['cust', 'truck'], as_index=False).agg(
        {"trans_total_vol": "sum", "trans_total_weight": "sum", "trans_total_price": "sum", "trans_total_cost": "sum",
         "plant": "nunique"})
    df_custtruck_output = pd.merge(df_custtruck.drop('id', axis=1), df_custtruck_output, on=['cust', 'truck'],
                                   how='left')
    df_custtruck_output['trans_cost_ratio'] = df_custtruck_output['cost_ratio'] * df_custtruck_output[
        'trans_total_cost']
    df_custtruck_output['id'] = df_custtruck_output['cust'] + df_custtruck_output['truck']
    df_custtruck_output = df_custtruck_output.set_index('id').reset_index()
    df_custtruck_output = df_custtruck_output.sort_values(by=['cust', 'truck']).reset_index(drop=True)

    # check by custmat
    df_custmat_output = pd.merge(df_custmat.drop(['id', 'price', 'cost'], axis=1),
                                 df_trans.groupby(['cust', 'mat'], as_index=False).agg({"trans_total_stk": "sum"}),
                                 on=['cust', 'mat'], how='left')
    df_custmat_output['trans_total_stk'] = df_custmat_output['trans_total_stk'].apply(
        lambda x: x if pd.notnull(x) else 0)
    df_custmat_output['inv_after'] = df_custmat_output['inv_qty'] + df_custmat_output['trans_total_stk']
    df_custmat_output['dos_before'] = df_custmat_output['inv_qty'] / df_custmat_output['sales_qty']
    df_custmat_output['dos_after'] = df_custmat_output['inv_after'] / df_custmat_output['sales_qty']
    df_custmat_output['id'] = df_custmat_output['cust'] + df_custmat_output['mat']
    df_custmat_output = df_custmat_output.set_index('id').reset_index()
    df_custmat_output = df_custmat_output.sort_values(by=['cust', 'mat']).reset_index(drop=True)

    # check by order
    df_order_output = pd.merge(df_order.drop('id', axis=1),
                               df_trans.groupby(['cust', 'mat'], as_index=False).agg({"trans_total_order": "sum"}),
                               on=['cust', 'mat'], how='left')
    df_order_output['trans_total_order'] = df_order_output['trans_total_order'].apply(
        lambda x: x if pd.notnull(x) else 0)
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
    json.dump(model_dict, open(os.path.join(model_dir, "model_dict.json"), 'w'))

    return model_dict

if __name__== "__main__":
    validate_model('./test', './test')
    optimize('./test', './test')
