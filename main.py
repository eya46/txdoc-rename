import os
from typing import Sequence, Tuple, Union
from pathlib import Path

import xlwings as xw
from xlwings import Sheet


def opt(
        file_path: Union[str, Path], sheet: str, rows: Sequence[str], line_start: int, line_end: int,
        rule: str, file_row: str
):
    """
    :param file_path: 文件路径
    :param rows: A~...
    :param sheet: 表
    :param rule: {A}_{B}.jpg
    :param line_start: 行开始
    :param line_end: 行结束
    :param file_row: 附件名称的列
    :return: [(raw_name,new_name),...]
    """
    b = xw.Book(Path(file_path))
    s: Sheet = b.sheets[sheet]
    try:
        return [
            (s[f"{file_row}{i}"].value, rule.format(**{row: s[f"{row}{i}"].value for row in rows}))
            for i in range(line_start, line_end + 1)
        ]
    finally:
        b.close()


def renames(fj_path: Union[str, Path], file_names: Sequence[Tuple[str, str]]):
    p = Path(fj_path)
    for raw_file_name, new_file_name in file_names:
        os.rename(p / raw_file_name, p / new_file_name)


if __name__ == '__main__':
    n: Path = Path(input("请输入文件完整路径："))
    sn: str = n.name.split(".xlsx")[0]
    rs: list[str] = input("请输入要使用的列号(例如:A,B,C)：").split(",")
    ls: int = int(input("请输入数据开始的行号："))
    le: int = int(input("请输入数据结束的行号："))
    r: str = input("请输入重命名规则(例如:{A}_{B}_{C}.jpg)：")
    fr: str = input("请输入附件名所在地列号：")
    ll = opt(
        n, sn, rs,
        ls, le, r, fr
    )
    print("重命名结果->")
    for j in ll:
        print(j[0], j[1])
    if input("确认要重命名吗?(yes/other)：") == "yes":
        renames(n.parent / "附件", ll)
        print("重命名结束~")
    else:
        print("取消~")
