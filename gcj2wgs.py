import sys,pandas as pd
from tkinter import filedialog,messagebox,Tk
from coord_convert.transform import gcj2wgs
from pathlib import Path

def do_trans()->Path|None:
    file_path =filedialog.askopenfilename(title="选择xlsx文件",filetypes=[('Excel文件','*.xlsx')])
    if file_path=='':
        sys.exit(0)
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
    do_trans()
    root = Tk()
    root.withdraw()          # 不显示空白主窗口

    # 弹出“确定”提示
    messagebox.showinfo("提示", "操作已完成！")   # 只有确定按钮

    root.destroy()           # 收掉隐藏的主窗口