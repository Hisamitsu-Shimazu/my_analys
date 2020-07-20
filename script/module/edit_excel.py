import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils import get_column_letter
import excel_style

def set_title(ws, title):
    ws.cell(2,2).value = title
    set_style(ws.cell(2,2), 'style_title')
    return ws

def set_style(cell, s):
    cell.font = excel_style.style[s][0]
    cell.fill = excel_style.style[s][1]
    cell.border = excel_style.style[s][2]
    cell.alignment = excel_style.style[s][3]

def put_dataframe(ws, df, start:tuple):
    # 左上のセル
    ws.cell(start[0], start[1]).value = ''
    set_style(ws.cell(start[0], start[1]), 'style_header')
#     ws.cell(start[0], start[0]).value = None
    
    # カラム代入
    for i,c in enumerate(df.columns):
        ws.cell(start[0], start[1]+1+i).value = str(c)
        set_style(ws.cell(start[0], start[1]+1+i), 'style_header')
        
    # インデックス代入
    for i,r in enumerate(df.index):
        ws.cell(start[0]+1+i, start[1]).value = str(r)
        set_style(ws.cell(start[0]+1+i, start[1]), 'style_header')
        
    # 値代入
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            if df.iat[i,j] != np.nan:
                ws.cell(start[0]+1+i, start[1]+1+j).value = str(df.iat[i,j])
            else:
                ws.cell(start[0]+1+i, start[1]+1+j).value = 'Nan'
            set_style(ws.cell(start[0]+1+i, start[1]+1+j), 'style_grid')
    return ws

def adjust_width(ws):
    for col in ws.columns:
        max_length = 0
        column = col[0].column
        
        for cell in col:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        
        adjusted_width = round((max_length + 2) * 1.2)
        ws.column_dimensions[get_column_letter(column)].width = adjusted_width