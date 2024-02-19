import oss2
from loguru import logger
from oss2 import Auth, Bucket
import sys
from io import BytesIO


# 设置timeout
# 公式结果、超链接
# archive_folder
# data_folder
# error_file_folder
# error_log_folder
# work_folder


class oss_manager:
    def __init__(self):
        # 请替换为你的实际AccessKey ID和AccessKey Secret
        self.access_key_id = "LTAI5tNEgBT3FdLGT9w5NWgN"
        self.access_key_secret = "rkQ4kXorILphTEGlXMJSS9es7dFK9v"
        self.endpoint = "http://oss-cn-hangzhou.aliyuncs.com"

    def bucket(self):
        self.auth = Auth(self.access_key_id, self.access_key_secret)
        # 请替换为你的OSS服务地址
        bucket = Bucket(
            self.auth, self.endpoint, "kering-cdn-prod-sharepoint"
        )  # Bucket名称

        # 使用get_bucket_acl()方法尝试获取Bucket的ACL信息，此操作会发送网络请求
        try:
            bucket.get_bucket_acl()
            print(f"\n连接和认证成功")
        except oss2.exceptions.NoSuchBucket:
            print(f"\nBucket不存在")
        except oss2.exceptions.OssError as e:
            print(f"\n连接或认证失败：{e}")
        return bucket

    # 获取目标文件

    def list_files(self, file_object, delimiter="/"):
        bucket = self.bucket()
        matching_files = []

        for obj in oss2.ObjectIterator(
            bucket=bucket, prefix=file_object, delimiter=delimiter
        ):
            if not obj.is_prefix() and obj.key.endswith(file_object):  # 根据文件名筛选
                check_file = obj.key
                matching_files.append(check_file)

        return matching_files

    def list_folders(self, folder_object, delimiter="/"):
        bucket = self.bucket()
        subfolders = []

        for obj in oss2.ObjectIterator(
            bucket, prefix=folder_object, delimiter=delimiter
        ):
            if obj.is_prefix():  # 判断obj为文件夹。
                subfolder = obj.key
                print("directory: " + subfolder)
                subfolders.append(subfolder)

        return subfolders

    # consumed_bytes表示已上传的数据量。
    # total_bytes表示待上传的总数据量。当无法确定待上传的数据长度时，total_bytes的值为None。
    def percentage(self, consumed_bytes, total_bytes):
        if total_bytes:
            rate = int(100 * (float(consumed_bytes) / float(total_bytes)))
            print(f"----- {rate}% ------", end="\r")
            sys.stdout.flush()

    # 将sharepoint原文件下载后上传到oss，并显示进度条
    def upload_with_progress(self, object_name, data_stream):
        bucket = self.bucket()
        bucket.put_object(
            object_name, data_stream.getvalue(), progress_callback=self.percentage
        )
        print(f"\n{object_name}已经上传到oss")

    # 获取oss文件数据流
    # 例：file_object = "check_folder/校验配置表.xlsx"
    def oss_data_stream(self, file_object):
        bucket = self.bucket()
        # check_file = oss_manager.list_file(bucket, file_object, delimiter)
        object_exists = bucket.object_exists(file_object)
        # 判断文件是否存在
        # 返回值为true表示文件存在，false表示文件不存在。
        # 填写Object的完整路径，Object完整路径中不能包含Bucket名称。
        if object_exists:
            print("object exist")
        else:
            print("object not exist")

        # 获取文件对象
        file = bucket.get_object(file_object)
        # read()方法将网络传输的字节数据转为二进制数据，然后BytesIO写入内存中二进制流对象，以便对数据操作
        data_stream = BytesIO(file.read())
        print("oss_data_stream()已经获取数据流!")
        return data_stream


"""
from . import credentials
def _authorize(temp_creds: credentials.Temporary) -> oss2.StsAuth:
    return oss2.StsAuth(
        temp_creds.access_key_id,
        temp_creds.access_key_secret,
        temp_creds.security_token,
    )


def get_bucket(
    temp_creds: credentials.Temporary, endpoint: str, bucket: str
) -> oss2.Bucket:
    auth = _authorize(temp_creds=temp_creds)
    return oss2.Bucket(
        auth,
        endpoint,
        bucket,
    )
"""
