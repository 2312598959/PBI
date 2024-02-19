from shareplum import Site, Office365
from shareplum.site import Version
from io import BytesIO


# SharePoint 用户名和密码
sharepointUsername = "toddy.zhou-ext@kering.com"
sharepointPassword = "K@ring96828"

# SharePoint 网站的URL
website = "https://kering.sharepoint.com"
# SharePoint 文档库的URL
sharepointSite = "https://kering.sharepoint.com"

# SharePoint 文档库
library_name = "Document Library/KE – DIGITAL WARRANTY CARD BI"
# sharepoint_directory: ["User Folder", "Archive Folder", "test folder"]

# 首先进行身份验证并获取 SharePoint 网站的授权 cookie
authcookie = Office365(
    website, username=sharepointUsername, password=sharepointPassword
).GetCookies()
# print(f"Auth cookie: {authcookie}")


try:
    site = Site(sharepointSite, version=Version.v365, authcookie=authcookie)
    # 输出网站的标题和描述
    print(f"SharePoint:{site.site_url}")
    print("连接成功!")
    print("site.version")


except Exception as e:
    print(f"连接或认证失败：{e}")


# 获取数据文件列表
def sharepoint_file_list(folder_name):
    folder_path = f"{library_name}/{folder_name}"

    # 创建 Folder 对象
    folder_all = site.Folder(folder_path)
    print("sharepoint_directory:", folder_all.folders)
    # 获取files字典
    file_dictionary = folder_all.files

    # for file_info in file_dictionary:
    #     file_name = file_info["Name"]
    #     print(file_name)
    file_names = [file_info["Name"] for file_info in file_dictionary]
    print("sharepoint_file_list()已经获取数据文件列表!", file_names)

    return file_names


# 获取sharepoint文件数据流
def sharepoint_data_stream(folder_name, subfile_name):
    folder_path = f"{library_name}/{folder_name}"

    # 创建 Folder 对象
    folder_all = site.Folder(folder_path)

    # 获取文件对象
    file = folder_all.get_file(subfile_name)

    # 获取文件夹对象
    # folder = folder_all.folders
    # print("sharepoint_directory:", folder)

    # 将二进制数据写入内存中二进制流对象，以便对数据操作
    data_stream = BytesIO(file)
    print("sharepoint_data_stream()已经获取数据流!")
    return data_stream


def sharepoint_archive_data(folder_name, archive_folder_name, subfile_name):
    folder_path = f"{library_name}/{folder_name}"

    # 创建 Folder 对象
    folder_all = site.Folder(folder_path)

    # 获取文件内容
    file = folder_all.get_file(subfile_name)

    archive_folder_path = f"{library_name}/{archive_folder_name}"

    # 创建 Folder 对象
    archive_folder_all = site.Folder(archive_folder_path)
    # 上传文件内容到归档文件夹
    archive_folder_all.upload_file(file, subfile_name)
    # folder_all.delete_file(subfile_name)
    print("sharepoint_archive_data()已经成功归档文件!")


# folder_name = "User Folder"
# sharepoint_file_list(folder_name)
