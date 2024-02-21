from sharepoint import (
    sharepoint_file_list,
    sharepoint_data_stream,
    sharepoint_archive_data,
)
from oss import oss_manager
from excel import excel_dfs
from date_format import format_file_name
from io import BytesIO
from loguru import logger


oss_check_file = "check_folder/校验配置表.xlsx"
oss_archive_folder = "archive_folder"
oss_data_folder = "data_folder"
sharepoint_folder_name = "User Folder"
sharepoint_archive_folder_name = "Archive Folder"


# 创建OSS实例
oss_manager = oss_manager()

# 获取oss中校验文件的数据流
oss_data_stream = oss_manager.oss_data_stream(oss_check_file)
# 获取oss中excel文件的数据字典
custom_sheet_names = ["表名及文件格式验证"]
dfs = excel_dfs(data_stream=oss_data_stream, custom_sheet_names=custom_sheet_names)

sheet_name = "表名及文件格式验证"
# 直接通过键名获取目标数据
if sheet_name in dfs:
    target_data = dfs[sheet_name]
    print("这是目标表的数据：", target_data)
    logger.info(f"这是目标表的数据：{target_data}")
    # exception, debug, warning, error,critical
else:
    print(f"无法找到名为 '{sheet_name}' 的表")
    logger.error(f"无法找到名为 '{sheet_name}' 的表")

# 获取第一列
# 方法一：如果第一列的索引为0
first_column = target_data.iloc[:, 0].tolist()
print("first_column:", first_column)
# # 方法二：如果知道第一列的列名
# first_column = target_data["列名"]
# print("first_column:", first_column)
# 获取第一行
first_row = target_data.iloc[0].tolist()
print("first_row:", first_row)
file_names = sharepoint_file_list(sharepoint_folder_name)
files, subfile_names = format_file_name(file_names, first_column)
print(files)
print(subfile_names)

# 循环读取sharepoint文件，同时开启校验。
subfile_name = "门店信息源表-20240220.xlsx"
if subfile_name in subfile_names:
    print(subfile_name)
    data_stream = sharepoint_data_stream(sharepoint_folder_name, subfile_name)

    # 将目标文件复制到归档文件夹下，删除原文件
    sharepoint_archive_data(
        sharepoint_folder_name, sharepoint_archive_folder_name, subfile_name
    )

    # 指定x-oss-forbid-overwrite为true时，表示禁止覆盖同名Object，如果同名Object已存在，程序将报错。
    headers = {"x-oss-forbid-overwrite": "true"}
    # 说明：OSS在上传和下载文件时默认开启CRC数据校验，确保上传和下载过程的数据完整性。如果上传后文件大小与本地文件大小不一致，则报错InconsistentError。
    # 校验通过的正确文件写入oss中归档文件夹下，并写入一份parquet文件到数据文件夹下。错误文件移到错误文件夹下
    # 定义OSS上的Parquet文件名
    new_file_name = subfile_name.split(".")[0]
    parquet_object_file = f"{oss_data_folder}/{new_file_name}.parquet"
    print(parquet_object_file)

    custom_sheet_names = ["门店信息表"]
    dfs = excel_dfs(
        data_stream=data_stream,
        header=0,
        nrows=50,
        custom_sheet_names=custom_sheet_names,
    )

    sheet_name = "门店信息表"
    # 直接通过键名获取目标数据
    if sheet_name in dfs:
        target_data = dfs[sheet_name]
        print("获得目标表的数据df")
    else:
        print(f"无法找到名为 '{sheet_name}' 的表")

    # 获取列名
    column_names = target_data.columns
    # 将列名进行格式化，替换掉 \n
    formatted_column_names = column_names.str.replace("\n", "")
    print("formatted_column_names:", formatted_column_names)
    # 将格式化后的列名重新赋值给 DataFrame
    target_data.columns = formatted_column_names

    # 将DataFrame写入Parquet文件
    parquet_data_stream = BytesIO()
    target_data.to_parquet(parquet_data_stream, engine="pyarrow")  # 使用pyarrow引擎

    # 重置BytesIO的指针位置以便读取
    parquet_data_stream.seek(0)

    # 将Parquet格式的数据流上传至OSS
    oss_manager.upload_with_progress(
        object_name=parquet_object_file, data_stream=parquet_data_stream
    )

    # 将错误文件上传oss的错误文件夹下，并显示进度条
    object_file = f"{oss_archive_folder}/{subfile_name}"
    data_stream.seek(0)
    oss_manager.upload_with_progress(object_name=object_file, data_stream=data_stream)

    # # 拷贝文件
    # result = bucket.copy_object(bucket_name, check_file_object, "check_folder/校验2.xlsx")
