# 公式插入API指南

公式插入API源代码为 [insert.py](api/insert.py)，包含Latex代码转Word对象、平均值/标准差/不确定度计算公式插入、最小二乘法结果插入三大功能，调用方法可参考 [exp5.py](module/exp5.py)。

## Latex 代码转 Word 对象

- 调用方法：
  `latex_to_word(latex_code)`
- 参数：
  latex_code：Latex 代码（字符串）
- 返回：
  一个 Word 对象
- 在 Word 中插入公式示例：
  `docu.add_paragraph()._element.append(latex_to_word(r"\frac{a}{b}"))`

## 平均值/标准差/不确定度计算公式插入

- 调用方法：
  `insert_data(docu, name, data, option)`
- 参数：
  docu：文档对象
  name：该物理量的名称和符号
  data：该物理量的相关数据（包含平均值/标准差/不确定度的元组）
  option：插入选项，`"word"`表示插入 Word 公式，`"latex"`表示插入 Latex 公式

## 最小二乘法结果插入

- 调用方法：
  `insert_data_lsm(docu, data, option)`
- 参数：
  docu：文档对象
  data：该物理量的相关数据（包含平均值/标准差/不确定度的元组）
  option：插入选项，`"word"`表示插入 Word 公式，`"latex"`表示插入 Latex 公式