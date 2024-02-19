from datetime import datetime


def is_valid_date_format(date_string, min_year=2020, max_year=2099):
    length_dict = {8: "%Y%m%d", 6: "%Y%m", 4: "%Y"}

    fmt = length_dict.get(len(date_string))

    if fmt:
        try:
            dt = datetime.strptime(date_string, fmt)

            # 对于 %Y%m%d 和 %Y%m 格式，检查月份的有效性（对于 %Y%m 也需检查日）
            if fmt == "%Y%m%d":
                month = dt.month
                day = dt.day

                if (
                    1 <= month <= 12
                    and 1 <= day <= 31
                    and min_year <= dt.year <= max_year
                ):
                    return (True, "yyyymmdd")
                else:
                    return (False, None)

            elif fmt == "%Y%m":
                month = dt.month

                if 1 <= month <= 12 and min_year <= dt.year <= max_year:
                    return (True, "yyyymm")

            # 对于只有年份的情况（%Y），检查是否在指定范围内
            elif fmt == "%Y":
                if min_year <= dt.year <= max_year:
                    return (True, "yyyy")
                else:
                    return (False, None)

        except ValueError:
            pass

    return (False, None)


def format_file_name(file_names, first_column):
    files = []
    subfile_names = []
    for file_name in file_names:
        # 检查文件扩展名是否为 "xlsx" 或 "xlsb"
        if file_name.endswith((".xlsx", ".xlsb")):
            # 检查文件名是否包含至少一个杠 '-'
            if "-" in file_name:
                split_parts = file_name.split("-")

                # 确保有足够的部分进行分割
                if len(split_parts) > 1:
                    split_suffix = split_parts[1].split(".")[0]
                    split_prefix = split_parts[0]

                    is_valid, date_format = is_valid_date_format(
                        date_string=split_suffix
                    )

                    if is_valid:
                        format_file_name = f"{split_prefix}-{date_format}"
                        print("format_file_name:", format_file_name)

                        # 假设 first_column 是一个列表，包含有效的前缀部分
                        if format_file_name in first_column:
                            files.append(format_file_name)
                            subfile_names.append(file_name)

                    else:
                        print(f"{file_name}, 不符合日期格式要求")

            else:
                print(f"{file_name}，不符合命名规范（缺少下划线 '-'）")

        else:
            print(f"{file_name}，不符合文件类型要求（应为 .xlsx 或 .xlsb）")

    return files, subfile_names
