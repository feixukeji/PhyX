# 蜗壳大雾实验工具开发说明

:stuck_out_tongue: You are welcomed to contribute, as long as you can read Chinese!

如果你只是用户，请直接访问[蜗壳大雾实验工具](https://dawu.feixu.site/)网站。

## 目录结构说明

```text
│ head.py（实验数据处理程序万能头）
│ main.py（主程序）
│ modulelist.py（模块清单）
│ requirements.txt（需要预先安装的包）
│ SourceHanSansSC-Regular.otf（作图字体文件）
│
├─api（API代码）
│      calc.py（数据处理API）
│      insert.py（公式插入API）
│      transformer.py（表达式转换）
│      MML2OMML.XSL（转Word对象）
│
├─module（存放各实验的数据处理程序）
│
├─static
│ ├─experiment（存放各实验的静态文件）
│ ├─layui（前端框架）
│ ├─pdf.js（来自Mozilla，用于PDF文件在线预览）
│ └─uncertainty（不确定度相关表格）
│
├─templates（前端代码）
│      404.html（错误页面）
│      experiment.html（具体实验处理页面）
│      index.html（网站首页）
│      uncertainty.html（不确定度相关表格页面）
│      viewer.html（PDF文件在线预览页面）
│
└─usrdata（存放程序运行时产生的文件）
```

## 需要预先安装的包

- chardet
- Flask
- latex2mathml
- lxml
- matplotlib
- numpy
- openpyxl
- pandas
- python-docx
- scipy
- sympy
- uncertainties

可以通过 `pip` 批量安装：

```bash
pip install -r requirements.txt
```

或者通过 `conda` 配置环境：

```bash
conda create -n phyx -c conda-forge python pip chardet flask lxml matplotlib numpy openpyxl pandas python-docx scipy sympy uncertainties
conda activate phyx
pip install latex2mathml
```

## 每个数据处理程序的开发流程

1. 编写数据处理程序，置于 [module](module) 文件夹内。
2. 将实验指导pdf（原讲义即可）、示例数据csv（示例数据主要体现格式，而数据不必正确合理）与png（生成png文件的的方法：在Excel中全选复制，在QQ聊天框或Windows画图软件中粘贴，然后另存为；若数据过多，请使用省略号，参考 exp23d；如有特别需要说明的，请用红色字体标注，参考 exp18）置于 static/experiment/expID/ 目录下，文件名**必须**与数据处理程序中的 `name()` 函数返回值一致。
3. 修改模块清单 [modulelist.py](modulelist.py)。
4. 修改前端代码 [templates/index.html](templates/index.html)。
5. 本地启动 [main.py](main.py) 测试运行。

（注：已统一写了前端及前后端衔接部分，故不必对每个实验编写网页）

## 开发规范与说明

- 建议使用 Python 3.9.6
- 文件名：expID.py，如 exp1.py（若一个实验有多个小实验，则在ID后加a/b/...，如 exp1a.py）
- 函数
  - `name` 返回实验名称
  - `handle` 处理数据并生成文档
- 变量名规范：由多个单词组成时，请使用下划线_（如 tight_layout），而不是大写（如 tightLayout,TightLayout）。
- 代码运行顺序：计算→作图→在Word中插入结果，不要一边计算一边在Word中插入结果。
- 单个数据相关的计算（如 `pi,sqrt,log`）使用 `math` 库函数，一组数据相关的计算（如 `mean,max,min`）优先考虑成员函数（如 `data.max()`），其次是 `numpy` 函数。
- 使用之前未曾使用过的库/函数（如 scipy.signal.savgol_filter），**必须**在注释中说明其功能。
- 提交的代码不应有多余的输出。
- 作图及线性拟合（参考 [exp5.py](module/exp5.py)）
  - 面向绘图对象作图（`fig, ax = plt.subplots()`）。
  - 设置副刻度为主刻度的一半，但保持主刻度为默认。
  - 刻度朝内（已在 main.py 中统一设置）。
  - 若一张图只有一组点线，则点使用红色（`color='r'`），线使用蓝色（`color='b'`），且线覆盖在点的上面；若一张图有多组点线，则同一组点线的颜色应当相同，并依次使用蓝(b)、红(r)、绿(g)、紫(m)、橙(orange)、青(c)。
  - 点的类型使用实心圆（`"o"`），若一张图有多组点线，则依次使用实心圆(o)、正方形(s)、上三角(^)、菱形(D)、下三角(v)、星号(*)。
  - 线条粗细使用 `linewidth=1.5`，点的大小使用 `markersize=3`，可视数据量、数据组数适当调整，但应保持统一性。
  - 绘制双y轴图参考 [exp15c.py](module/exp15c.py)。
  - 只有一组点线的图，一般不显示图例。
  - 图像字体使用 SourceHanSansSC-Regular.otf。
  - 轴标签和标题支持 Latex。
- 不确定度计算及线性回归参见 [数据处理API指南](数据处理API指南.md)。
- 非线性拟合一般可通过坐标变换转换为线性拟合，若无法转换，则使用 `scipy.optimize.curve_fit`（参考 [exp15c.py](module/exp15c.py)）
- 关于文档
  - 字体使用微软雅黑。
  - 文档第一行是实验名称（即 `name()` 函数返回值），随后注明“【Latex 代码在下面，请向下翻阅】”。
  - 计算过程的 Word 公式放在文档的前半部分，Latex 公式放在文档的后半部分。公式插入API可参见 [公式插入API指南](公式插入API指南.md)。
  - 内容跨度较大的段落之间应当用一个空行。
  - 文档中插入的数据一般保留4或5位有效数字（`'%.5g' % x`），线性拟合的相关系数$r$建议保留8位有效数字。
  - 若某张图片正好在第2页开头，而第1页尾部有很多空白区域，为避免误解，应在第1页的最后一个段落之后注明“【本文档不只有一页，请向下翻阅】”。
  - 插入表格参考 [exp5.py](module/exp5.py)。

## 协作方法

在 Github 上协作非常简单，只需进行以下四个步骤。

### 建立 fork

本页右上角有一个 `Fork` 按钮，点击它会出现如下界面：

![Mew fork](https://s2.loli.net/2022/08/15/5FskUI1WhOql3n8.png)

直接点击 `Create fork` 即可。

### 作出修改

`Create fork` 会在你的账户下创建一个 Repository，其内容与本处 Repository 的内容一样，但你拥有一切权限。这时，你就可以在你的这个 Repository 里自由地进行修改了。

![Commit changes](https://s2.loli.net/2022/08/15/wKltBaYsIj8ASpW.png)

### 准备 Pull request

当你作完了修改后，便可将这些修改提交予我们，以改进本项目。在你的 Repository 主页按照下面图片操作：

![Open pull request](https://s2.loli.net/2022/08/15/TbqXjed3lOhA4Jv.png)

### 提交 Pull request

在 `Open a pull requst` 的表单中，直接将内容提交至我们的 `main` base 即可。填写完表单后点击 `Create pull request`。

![Create pull request](https://s2.loli.net/2022/08/15/4krCp8MSNnehH7T.png)

**至此，你已经成功提交了你的修改。**

随后，在本项目的 `Pull requests` 选项卡中会出现你的提交，我们会心怀感激地接纳你的修改，或与你进一步讨论。

![Merge pull request](https://s2.loli.net/2022/08/15/s3CrZJvXItwyxgn.png)

或者，如果你有好的想法，也欢迎[提出issue](https://github.com/feixukeji/PhyX/issues)。

> 参考：[Fork a repo](https://docs.github.com/en/get-started/quickstart/fork-a-repo), [Pull requests](https://docs.github.com/en/pull-requests), [Creating an issue](https://docs.github.com/en/issues/tracking-your-work-with-issues/creating-an-issue#creating-an-issue-from-a-repository)

## 学习参考

- Python 3: [Documentation](https://docs.python.org/3/), [Tutorial](https://docs.python.org/3/tutorial/), [教程](https://www.runoob.com/python3/python3-tutorial.html)
- NumPy: [Documentation](https://numpy.org/doc/), [Learn](https://numpy.org/learn/), [教程](https://www.runoob.com/numpy/numpy-tutorial.html)
- Matplotlib: [Documentation](https://matplotlib.org/stable/index.html), [Tutorial](https://matplotlib.org/stable/tutorials/index.html), [教程](https://www.runoob.com/matplotlib/matplotlib-tutorial.html)
- Pandas: [Documentation](https://pandas.pydata.org/docs/), [User Guide](https://pandas.pydata.org/docs/user_guide/index.html), [教程](https://www.runoob.com/pandas/pandas-tutorial.html)
- Python-docx: [Documentation](https://python-docx.readthedocs.io/en/latest/), [Quickstart](https://python-docx.readthedocs.io/en/latest/user/quickstart.html)

## To Do

1. 大学物理二级、三级、四级实验数据处理程序的开发。
2. 网页端直接输入和显示（目前只支持上传、下载文件）。
3. 手写表格数字识别：把手写的实验数据转换成 Excel（csv）文件。
4. PDF在线预览（目前在部分手机浏览器上无法预览）。

## Contributors

- 组织策划&前端&前后端衔接程序&公式插入API&数据处理程序示例&开发文档编写&代码审查：孙旭磊
- 数据处理API：孙旭磊、张学涵、周旭冉、尹冠霖
- 技术与安全支持：赵奕
- 各实验对应的数据处理程序：
  |ID|实验|分类|开发|
  |-|-|-|-|
  |0|通用工具|通用|孙旭磊|
  |1|重力加速度的测量|力热|秦沁|
  |2|表面张力|力热|赵奕|
  |3|黏滞系数|力热|孙旭磊|
  |4|质量和密度的测量|力热|鲍政廷|
  |5|钢丝杨氏模量|力热|孙旭磊|
  |6|切变模量|力热|鲍政廷|
  |7|固体比热|力热|张学涵|
  |8|匀加速运动|力热|孙旭磊|
  |9|声速测量|力热|张学涵|
  |10|磁力摆|力热|（待开发）|
  |11|半导体温度计|电磁|张学涵|
  |12|示波器的使用|电磁|鲍政廷|
  |13|整流滤波|电磁|赵奕|
  |14|直流电源特性|电磁|张学涵|
  |15|硅光电池|电磁|夏熙林|
  |16|RGB配色|电磁|秦沁|
  |17|数字体温计|电磁|鲍政廷|
  |18|分光计|光学|赵奕|
  |19|干涉法测微小量|光学|赵奕|
  |20|透镜参数测量|光学|孙旭磊|
  |21|显微镜使用|光学|鲍政廷|
  |22|衍射实验|光学|鲍政廷|
  |23|光电效应|近代|张学涵|
  |24|密立根油滴|近代|秦沁|
  |25|生活中的物理实验|生活|秦沁|

## License

PhyX is released under the [AGPL-3.0 license](./LICENSE).
