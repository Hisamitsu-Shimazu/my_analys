import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils import get_column_letter
import excel_style
from pathlib import Path
import math

# styleの設定
def set_style(cell, s):
    cell.font = excel_style.style[s][0]
    cell.fill = excel_style.style[s][1]
    cell.border = excel_style.style[s][2]
    cell.alignment = excel_style.style[s][3]

# データフレームの挿入
def put_dataframe(ws, df, start:tuple, mode='normal'):
    # 左上のセル
    set_cell(ws, cell=(start[0], start[1]), value='', style='style_header')
    
    # カラム代入
    for i,c in enumerate(df.columns):
        set_cell(ws, cell=(start[0], start[1]+1+i), value=str(c), style='style_header')
        
    # インデックス代入
    for i,r in enumerate(df.index):
        set_cell(ws, cell=(start[0]+1+i, start[1]), value=str(r), style='style_header')
        
    # 値代入
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            if math.isnan(df.iat[i,j]):
                if mode == 'describe':
                    set_cell(ws, cell=(start[0]+1+i, start[1]+1+j), value='-', style='style_grid')
                else :
                    set_cell(ws, cell=(start[0]+1+i, start[1]+1+j), value='Nan', style='style_grid')
            else :
                set_cell(ws, cell=(start[0]+1+i, start[1]+1+j), value=df.iat[i,j], style='style_grid')

# 幅調整
def adjust_width(ws):
    for col in ws.columns:
        max_length = 0
        column = col[0].column
        
        for cell in col:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        
        adjusted_width = (max_length + 2) * 1.2
        if adjusted_width > 35:
            ws.column_dimensions[get_column_letter(column)].width = 35
        else:
            ws.column_dimensions[get_column_letter(column)].width = adjusted_width

# 列番号、行番号から文字列取得
def get_letter(r, c):
    return openpyxl.utils.get_column_letter(c)+str(r)
            
# 画像貼り付け
def put_img(ws, fig, c, start:tuple):
    img_dir = Path('../output/img/')
    fig.savefig(img_dir/f'{c}_fig.png', transparent=True, bbox_inches='tight')
    img = openpyxl.drawing.image.Image(img_dir/f'{c}_fig.png')
    ws.add_image(img, get_letter(start[0], start[1]))
    
# 値とstyleを設定
def set_cell(ws, cell:tuple, value=None, style='normal'):
    ws.cell(cell[0], cell[1]).value = value
    set_style(ws.cell(cell[0], cell[1]), style)