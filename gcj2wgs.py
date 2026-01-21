import pandas as pd
from coord_convert.transform import gcj2wgs
from pathlib import Path

input_path = Path(r"D:\MyCode\XY_trans\试运行.xlsx")
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

print(f'已生成 {output_path.name}，预览：')
print(df[['坐标', '坐标_wgs']])