# _*_ coding:utf-8 _*_
"""
:Author: DarkLii
:Date  : 2025/09/12
"""
import copy
import shutil

import pandas as pd


def convert_testcase_format(testcases: list[dict]) -> pd.DataFrame:
    """
    将测试用例从原始格式转换为目标格式的DataFrame

    参数:
        testcase: 原始格式的测试用例字典，包含:
            - casetitle: 用例标题
            - priority: 优先级
            - precondition: 前置条件
            - steps: 步骤数组，每个元素包含step和expect

    返回:
        包含转换后测试用例的DataFrame，列包括:
            - 用例标题
            - 优先级
            - 所属模块 (默认为空)
            - 前置条件
            - 步骤: 合并后的步骤字符串
            - 预期: 合并后的预期结果字符串
    """
    converted_cases = []
    for testcase in testcases:
        converted = {
            "用例标题": testcase.get("casetitle", ""),
            "优先级": testcase.get("priority", 2),
            "所属模块": testcase.get("func_model", ""),  # 默认为空
            "前置条件": testcase.get("precondition", ""),
            "步骤": "\n".join(
                f"{i + 1}.{step['step']}"
                for i, step in enumerate(testcase.get("steps", []))
            ),
            "预期": "\n".join(
                f"{i + 1}.{step['expect']}"
                for i, step in enumerate(testcase.get("steps", []))
            )
        }
        converted_cases.append(copy.deepcopy(converted))
    return pd.DataFrame(converted_cases)


def _invoke_dataframe_with_template(df, output_path):
    # df.replace("<br>", r"\n", regex=True, inplace=True)
    df_dct = df.to_dict()
    # case_template = r"D:\hugh\code\pz\convert_testcase_format\test_case_tmp.xlsx"
    case_template = r"D:\hugh\code\pz\convert_testcase_format\case_template.xlsx"
    shutil.copy(case_template, output_path)
    df2 = pd.read_excel(case_template)
    dct = df2.to_dict()
    for key in df_dct.keys():
        dct[key] = df_dct[key]
    df3 = pd.DataFrame.from_dict(dct)
    with pd.ExcelWriter(output_path, engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
        df3.to_excel(writer, sheet_name='用例', index=False)


def export_testcases_to_excel(testcases: list[dict], output_path: str):
    """
    将测试用例导出到Excel文件

    参数:
        testcases: 原始格式的测试用例列表
        output_path: 输出Excel文件路径
    """
    df = convert_testcase_format(testcases)
    _invoke_dataframe_with_template(df, output_path)
    print("测试用例已导出到Excel文件")


if __name__ == "__main__":
    data = [
  {
    "casetitle": "Layout选中单个器件后Logic高亮并居中显示",
    "priority": 1,
    "precondition": "1. 打开原理图编辑器和Layout编辑器\n2. 建立Logic与Layout的连接\n3. 确保原理图和Layout中存在对应器件（如Q1）",
    "steps": [
      {"step": "在Layout中选中器件Q1", "expect": "器件Q1在Layout中高亮显示"},
      {"step": "观察Logic中的原理图", "expect": "Logic中对应器件Q1高亮显示并自动缩放至画布中央"}
    ]
  },
  {
    "casetitle": "Layout选中器件管脚后Logic高亮并居中显示所属器件",
    "priority": 1,
    "precondition": "1. 打开原理图编辑器和Layout编辑器\n2. 建立Logic与Layout的连接\n3. 确保原理图和Layout中存在对应器件及其管脚",
    "steps": [
      {"step": "在Layout中选中器件Q1的一个管脚", "expect": "管脚在Layout中高亮显示"},
      {"step": "观察Logic中的原理图", "expect": "Logic中管脚所属器件Q1高亮显示并自动缩放至画布中央"}
    ]
  },
  {
    "casetitle": "Layout选中多个器件后Logic缩放容纳所有器件",
    "priority": 2,
    "precondition": "1. 打开原理图编辑器和Layout编辑器\n2. 建立Logic与Layout的连接\n3. 确保原理图和Layout中存在多个对应器件（如Q1、Q2）",
    "steps": [
      {"step": "在Layout中同时选中器件Q1和Q2", "expect": "器件Q1和Q2在Layout中高亮显示"},
      {"step": "观察Logic中的原理图", "expect": "Logic中Q1和Q2高亮显示，画布缩放至刚好容纳所有选中器件，器件居中显示"}
    ]
  },
  {
    "casetitle": "Layout通过网表导入后选中器件Logic正确响应",
    "priority": 2,
    "precondition": "1. 打开原理图编辑器和Layout编辑器\n2. 在Logic中生成并发送网表到Layout\n3. Layout成功导入网表生成PCB文件",
    "steps": [
      {"step": "在Layout中选中器件Q1", "expect": "器件Q1在Layout中高亮显示"},
      {"step": "观察Logic中的原理图", "expect": "Logic中对应器件Q1高亮显示并自动缩放至画布中央"}
    ]
  },
  {
    "casetitle": "未建立Logic和Layout连接时选中器件无响应",
    "priority": 3,
    "precondition": "1. 打开原理图编辑器和Layout编辑器\n2. 未建立Logic与Layout的连接\n3. 确保Layout中存在器件（如Q1）",
    "steps": [
      {"step": "在Layout中选中器件Q1", "expect": "器件Q1在Layout中高亮显示"},
      {"step": "观察Logic中的原理图", "expect": "Logic中无任何器件高亮或缩放，画布状态不变"}
    ]
  },
  {
    "casetitle": "Layout选中不存在于Logic的器件时无响应",
    "priority": 3,
    "precondition": "1. 打开原理图编辑器和Layout编辑器\n2. 建立Logic与Layout的连接\n3. Layout中存在Logic原理图中不存在的器件（如Q3）",
    "steps": [
      {"step": "在Layout中选中器件Q3", "expect": "器件Q3在Layout中高亮显示"},
      {"step": "观察Logic中的原理图", "expect": "Logic中无任何器件高亮或缩放，画布状态不变"}
    ]
  },
  {
    "casetitle": "Layout快速连续选中多个器件Logic正确响应",
    "priority": 3,
    "precondition": "1. 打开原理图编辑器和Layout编辑器\n2. 建立Logic与Layout的连接\n3. 确保原理图和Layout中存在多个对应器件（如Q1、Q2）",
    "steps": [
      {"step": "在Layout中快速连续选中器件Q1和Q2", "expect": "器件Q1和Q2在Layout中依次高亮显示"},
      {"step": "观察Logic中的原理图", "expect": "Logic中Q1和Q2依次高亮显示，画布动态缩放至刚好容纳选中器件并居中"}
    ]
  },
  {
    "casetitle": "Layout选中器件后Logic中非当前sheet器件跳转",
    "priority": 4,
    "precondition": "1. 打开原理图编辑器和Layout编辑器\n2. 建立Logic与Layout的连接\n3. 确保原理图包含多个sheet页，器件Q1位于非当前sheet",
    "steps": [
      {"step": "在Layout中选中器件Q1", "expect": "器件Q1在Layout中高亮显示"},
      {"step": "观察Logic中的原理图", "expect": "Logic自动切换到Q1所在sheet页，Q1高亮显示并缩放至画布中央"}
    ]
  }
]
    output_path = r"D:\hugh\code\pz\convert_testcase_format\test_case.xlsx"
    export_testcases_to_excel(data, output_path)
