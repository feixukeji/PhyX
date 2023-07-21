from head import *  # 导入万能头


def name():  # 返回实验名称
    return "最小二乘法线性回归"


def handle(workpath, extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf")  # 设置图像中的文字字体

        excelpath = workpath+name()+'.'+extension  # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension == 'csv':
            with open(excelpath, 'rb') as f:
                encode = chardet.detect(f.read())["encoding"]  # 判断编码格式
            data = pd.read_csv(excelpath, header=None, names=["x", "y", "unit_x", "unit_y", "unit_m"], encoding=encode)  # 读取csv文件
        else:
            data = pd.read_excel(excelpath, header=None, names=["x", "y", "unit_x", "unit_y", "unit_m"])  # 读取xls/xlsx文件

        os.remove(excelpath)  # 读取Excel数据后删除文件

        x = pd.Series([float(x) for x in data["x"][1:]])
        y = pd.Series([float(y) for y in data["y"][1:]])
        res = analyse_lsm(x, y, data["x"][0], data["y"][0], data["unit_m"][1], data["unit_y"][1])  # 最小二乘多项式拟合之线性回归

        ax.clear()  # 新建绘图对象

        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(x, y, "o", color='r', markersize=3)  # 绘制数据点
        ax.plot(x, res.b + res.m * x, color='b', linewidth=1.5)  # 拟合直线
        # 作图，详见 https://www.runoob.com/matplotlib/matplotlib-marker.html 和 https://www.runoob.com/matplotlib/matplotlib-line.html
        ax.set_xlabel(data["x"][0] + "/" + data["unit_x"][1], fontproperties=zhfont)
        ax.set_ylabel(data["y"][0] + "/" + data["unit_y"][1], fontproperties=zhfont)
        # 添加标题和轴标签，详见 https://www.runoob.com/matplotlib/matplotlib-label.html

        imgpath = workpath + "img.jpg"
        fig.savefig(imgpath, dpi=300, bbox_inches='tight')  # 保存生成的图像

        docu = Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')  # 设置Word文档字体

        docu.add_paragraph(name())  # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()

        docu.add_picture(imgpath)

        insert_data_lsm(docu, res, "word")

        docu.add_paragraph("【Latex代码】")
        insert_data_lsm(docu, res, "latex")

        docu.save(workpath+name()+".docx")  # 保存Word文档，注意文件名必须与name()函数返回值一致

        os.remove(imgpath)  # 删除刚才保存的图像

        return 0  # 若成功，返回0
    except:
        traceback.print_exc()  # 打印错误
        return 1  # 若失败，返回1
