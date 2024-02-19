from sqlite_sql import sqlite_manager
from oss import oss_manager
from excel import excel_dfs


oss_file = "check_folder/校验配置表.xlsx"


# 创建SQLiteManager实例
db_manager = sqlite_manager()
# 创建OSS实例
oss_manager = oss_manager()

# 获取oss中excel文件的数据集
oss_data_stream = oss_manager.oss_data_stream(oss_file)
# 获取oss中excel文件的数据集
excel_dfs = excel_dfs(oss_data_stream)
# 遍历所有工作表的数据
for table_name, data_stream in excel_dfs.items():
    print("表名:", table_name)
    print("表数据:", data_stream.head())
    # 写入数据到SQLite数据库
    db_manager.to_sql(df=data_stream, table_name=table_name)
    # 执行查询并获取DataFrame结果
    query_result_df = db_manager.query(sql=f"SELECT * FROM {table_name}")
    # 输出查询结果数据的前几行
    print("sql查询结果:", query_result_df.head())

# 完成所有操作后关闭数据库连接
db_manager.close()
