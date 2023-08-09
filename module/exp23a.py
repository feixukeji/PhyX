from head import * # 导入万能头

def name(): # 返回实验名称
    return "光电效应测普朗克常量"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["v","U0"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["v","U0"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        res=analyse_lsm(data["v"], data["U0"], 'ν', 'U', 'V/Thz', 'V')
        fig, ax = plt.subplots()  # 新建绘图对象
        ax.plot(data["v"], data["U0"], "o", color='r', markersize=3)
        ax.plot(data["v"], res.b + res.m*data["v"], color='b', linewidth=1.5)
        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.set_title("遏止电压与单色光频率的关系", fontproperties=zhfont)
        ax.set_xlabel("Frequency (THz)")
        ax.set_ylabel("Stopping Potential (V)")

        imgpath=workpath+"img.jpg"
        fig.savefig(imgpath, dpi=300, bbox_inches='tight')
        plt.close()

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

        docu.add_paragraph(name())
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()
        e = 1.602e-19
        h0 = 6.626e-34
        docu.add_paragraph("取电子的电荷量 e = "+str(e)+" C，普朗克常量的公认值 h0 = "+str(h0)+" J⋅s")
        docu.add_picture(imgpath)
        insert_data_lsm(docu, res, option="word")
        h = abs(res.m)*e*10**-12
        docu.add_paragraph("普朗克常量 h = e|m| = "+'%.5g'%h+" J⋅s")
        docu.add_paragraph("相对误差 E = |h-h0|/h0 = "+'%.5g'%(abs(h-h0)/h0*100)+"%")
        docu.add_paragraph("与 x 轴截距为 -b/m = "+'%.5g'%(-res.b/res.m)+" Thz，即为红限")
        docu.add_paragraph("逸出功 A = e|b| = "+'%.5g'%abs(e*res.b)+" J")
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码】")
        insert_data_lsm(docu, res, option="latex")

        docu.save(workpath+name()+".docx")

        os.remove(imgpath)

        return 0
    except:
        traceback.print_exc() # 打印错误
        return 1