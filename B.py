import os
from typing import Sequence, Tuple, Union, Optional, Dict
from pathlib import Path

import xlwings as xw
from xlwings import Sheet


def opt(
        file_path: Union[str, Path], sheet: str, x_row: str, f_row: str, r: str
):
    b = xw.Book(Path(file_path))
    s: Sheet = b.sheets[sheet]

    line_end = -1
    for i in range(2, 200):
        if s[f"{x_row}{i}"].value is None:
            line_end = i
            break

    try:
        return [
            (_, r.format(**get_xh_name_class(s[f"{x_row}{i}"].value)) + "." + _.split(".")[-1])
            for i in range(2, line_end) if (_ := s[f"{f_row}{i}"].value)
        ]
    finally:
        b.close()


def renames(new_folder: str, fj_path: Union[str, Path], file_names: Sequence[Tuple[str, str]]):
    p = Path(fj_path)
    try:
        if os.path.exists(p / new_folder):
            print("文件夹已存在~")
        else:
            os.mkdir(p / new_folder)
    except Exception as e:
        print(f"创建新文件夹失败:{e}", e)
    for raw_file_name, new_file_name in file_names:
        try:
            os.rename(p / raw_file_name, p / new_folder / new_file_name)
        except Exception as e:
            print(f"重命名失败 {raw_file_name} {new_file_name}", type(e), e)


def get_xh_name_class(txt: str) -> Optional[Dict[str, str]]:
    for i in txt.split():
        if i.startswith("学生："):
            _1, _2, _3 = i[3:].split("-")
            return {
                "姓名": _1,
                "学号": _2,
                "班级": _3
            }
    return None


def get_xlsx(path_name: Path):
    for i in os.listdir(path_name):
        if i.endswith(".xlsx"):
            return i
    return None


if __name__ == '__main__':
    n: Path = Path(input("请输入文件夹完整路径："))

    if (sn := get_xlsx(n)) is None:
        print("该路径下没有收集结果！")
        exit(-1)

    xr = input("请输入学生列号：")
    fr = input("请输入附件列号：")

    new_folder_name: str = input("请输入保存的文件夹名称：")
    r: str = input("请输入重命名规则(例如:{学号}_{姓名}_{班级}_)：")
    ll = opt(
        n / sn, sn.split(".xlsx")[0], xr, fr, r
    )
    print("重命名结果->")
    for j in ll:
        print(j[0], j[1])
    if input("确认要重命名吗?(yes/other)：") == "yes":
        renames(new_folder_name, n / "附件", ll)
        print("重命名结束~")
    else:
        print("取消~")
