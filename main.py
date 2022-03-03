import pandas as pd

from pathlib2 import Path

import re



# 定义生成主站点表函数
def generateCSV(Originsheet_path,  Output_path):
    # 获取站名
    Station = re.match(r'.*站', Originsheet_path.name)
    # 创建生成点表目录
    Dir_name = Output_path / '主站CSV' / Station.group()
    Path(Dir_name).mkdir(parents=True, exist_ok=True)
    # 读取站端遥信表
    Origin_sheet_yx = pd.read_excel(Originsheet_path, "遥信")

    # 保存主站遥信表
    YX_path = Dir_name / 'csv_dig.csv'
    Origin_sheet_yx.to_csv(YX_path, encoding="gbk", index=None)

    # 读取主站遥测表
    Origin_sheet_yc = pd.read_excel(Originsheet_path, "遥测")

    # 保存主站遥测表
    YC_path = Dir_name / 'csv_ana.csv'
    Origin_sheet_yc.to_csv(YC_path, encoding="gbk", index=None)

    # 读取主站遥控表
    Origin_sheet_yk = pd.read_excel(Originsheet_path, "遥控")

    # 保存主站遥控表
    YK_path = Dir_name / 'csv_digctl.csv'
    Origin_sheet_yk.to_csv(YK_path, encoding="gbk", index=None)
    print('%s生成CSV文件成功' %Station.group())

# 当前项目路径
Current_path = Path.cwd()

# 原始主站点表文件路径
Originsheet_dir = Current_path / '主站点表'
for Originsheet_path in Originsheet_dir.rglob('*.xl*'):
    # 调用生成函数
    generateCSV(Originsheet_path,  Current_path)
