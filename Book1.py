import xlwings as xw
import pandas as pd
import numpy as np

@xw.func
def main():
    return 'hello'

@xw.func
def join_rows(data):
    dt = np.array(data)
    rows, cols = dt.shape[0], dt.shape[1]
    if rows == 0:
        return None
    if cols != 4:
        return None
    #
    dt = []
    for row in data:
        item = []
        for col in row:
            item.append(col)
        dt.append(item)
    #
    cols = ['msp', 'name', 'unit', 'price']
    df = pd.DataFrame(dt, columns=cols)
    #
    def makeNote(msp):
        dfs = df[df['msp']==msp]
        note = ''
        for i, v in dfs.iterrows():
            price = v['price']
            unit = v['unit']
            note += f'{price}/{unit}, '
        if len(note) > 0:
            note = note[0: -2]
        return note
    df['Note'] = df['msp'].apply(makeNote)
    df = df.groupby('msp').agg({
        'msp': 'first',
        'name': 'first',
        'Note': 'first'
    })
    dt = [['msp', 'name', 'note']]
    for i, v in df.iterrows():
        dt.append([v['msp'], v['name'], v['Note']])
    return dt



