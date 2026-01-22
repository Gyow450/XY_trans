import sys,pandas as pd
from tkinter import filedialog,messagebox,Tk
from coord_convert.transform import gcj2wgs
from pathlib import Path

def do_trans(file_path:str)->Path|None:
    
    input_path = Path(file_path)
    output_path = input_path.with_name(input_path.stem+'_转换'+input_path.suffix)
    # 读 Excel
    df = pd.read_excel(input_path, engine='openpyxl')

    # 拆分经纬度
    df[['经度', '纬度']] = df['坐标'].str.split(',', expand=True).astype(float)

    # 转换
    df[['经度_wgs', '纬度_wgs']] = df.apply(
        lambda row: pd.Series(gcj2wgs(row['经度'], row['纬度'])), axis=1)

    # 拼成新坐标列
    df['坐标_wgs'] = df['经度_wgs'].astype(str) + ',' + df['纬度_wgs'].astype(str)

    # 保存
    df.to_excel(output_path, index=False, engine='openpyxl')

    return output_path

if __name__=='__main__':
    file_paths =filedialog.askopenfilenames(title="选择xlsx文件",filetypes=[('Excel文件','*.xlsx')])
    if not file_paths:
        sys.exit(0)
    for file in file_paths:
        do_trans(file)
    root = Tk()
    root.withdraw()          # 不显示空白主窗口

    # 弹出“确定”提示
    messagebox.showinfo("提示", f"完成{len(file_paths)}个文档的操作")   # 只有确定按钮

    root.destroy()           # 收掉隐藏的主窗口