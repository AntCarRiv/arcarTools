import json
import os
import shutil
from glob import glob

import numpy as np
from tqdm import tqdm
from copy import deepcopy

import ngram
import pandas as pd


def get_element_father(z):
    return z.iloc[0, 0], z.iloc[0, 1]


def read_configuration(path='config.json'):
    with open(path, 'r') as ff:
        return json.load(ff)


def add_column(df, name, value):
    df[name] = value if value == 0 else np.nan


def update_rows(df, column_r, column, key, value):
    df.loc[:, column_r][df.loc[:, column] == key] = value
    return df


def compare_elements(name_1, n_grama: ngram.NGram):
    res = n_grama.search(name_1)
    for k, v in res:
        if int(v) == 1:
            return True
    return False


def get_by_split(name_product, product_in_sheet, old=None):
    split_name = str(',').join(name_product.split(',')[:-1])
    try:
        split_sheet = product_in_sheet.split(',')
        n_grama = ngram.NGram(split_sheet)
        if len(name_product.split(',')) == 2:
            split_name_2 = str(' ').join(old.split()[:3])
            r = compare_elements(split_name_2, n_grama)
            if r:
                return r
            else:
                get_by_split(split_name, product_in_sheet, old)
        elif len(name_product.split(',')) == 1:
            return False
        r = compare_elements(split_name, n_grama)
        if r:
            return r
        else:
            return get_by_split(split_name, product_in_sheet, name_product)
    except AttributeError:
        return False


def read_file_main(path='CLAVES DEL SAT .xlsx'):
    dfs = pd.read_excel(path, sheet_name=None, header=None)
    for sheet in dfs:
        dfs[sheet] = dfs[sheet].dropna(how='all', axis=1).dropna(axis=0, how='all')
        col = list(dfs[sheet].columns)
        try:
            col[0] = 'Clave'
            col[1] = 'Producto'
        except IndexError:
            pass
        dfs[sheet].columns = col
    return dfs


def read_file_invoice(path):
    dfs = pd.read_excel(path, sheet_name=None)
    for sheet in dfs:
        dfs[sheet] = dfs[sheet].dropna(how='all', axis=1).dropna(axis=0, how='all')
    return dfs


def find_product(name_prduct, books, fiability=0.9):
    for book in books:
        try:
            serie = books[book][
                books[book]['Producto'].apply(lambda x: ngram.NGram.compare(str(x), name_prduct, N=1) >= fiability)]
        except Exception:
            pass
        if not serie.empty:
            code_father = get_element_father(books[book])
            if fiability < 0.9:
                return {"Clave": code_father[0], "Descripcion": code_father[1], "Estatus": "Revisar Posible error"}
            return {"Clave": code_father[0], "Descripcion": code_father[1]}
    for book in books:
        if not books[book][books[book]['Producto'].apply(lambda x: get_by_split(name_prduct, x))].empty:
            code_father = get_element_father(books[book])
            return {"Clave": code_father[0], "Descripcion": code_father[1],
                    "Estatus": "Revisar Encontrado Parcialmente"}
    if fiability > 0.5:
        return find_product(name_prduct, 0.5)


def process_invoices(dfs, books):
    config = read_configuration()
    impuestos = config['impuestos']
    new_files = {}
    cols = config['Headers']
    for sheet in dfs:
        sheet_df = deepcopy(dfs[sheet])
        for i in cols:
            if i not in sheet_df:
                add_column(sheet_df, i, cols[i])
        items = tqdm(zip(sheet_df['Producto'], sheet_df['Producto'].isna()))
        for product, is_nan in items:
            items.set_description(f'Procesando: {sheet} Producto: {product}')
            if not is_nan:
                result = find_product(product, books)
                for k in result:
                    sheet_df = update_rows(sheet_df, k, 'Producto', product, result[k])
        new_files[sheet] = sheet_df

        sheet_df['Precio Unitario Sin Impustos'] = sheet_df['Precio Unitario'] - sheet_df['Descuento']
        sheet_df['support'] = sheet_df['Tipo de Impuesto'].apply(lambda x: impuestos.get(x, 1))
        sheet_df['Precio Unitario Sin Impustos'] = sheet_df['Precio Unitario Sin Impustos'].apply(lambda x: float(x)) / \
                                                   sheet_df['support']

        sheet_df[cols].to_excel(os.path.join(os.getcwd(), 'FacturasConCodigo',
                                             f'Encontrados_para_{sheet}.xlsx'), index=False)


def main():
    b = read_file_main()  # Book main sat

    for factura in glob(os.path.join('FacturasPendientes', '*.xlsx')):
        p_i = read_file_invoice(factura)  # Book invoice
        process_invoices(p_i, b)
        shutil.move(factura, 'Facturas')


main()
