from head import * # 导入万能头

def name(): # 返回实验名称
    return "单缝衍射"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["L","d","y","k","x"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["L","d","y","k","x"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        data["x"]-=data["y"][0]

        res_lsm=analyse_lsm(data["k"], data["x"], "k", "x_k", "mm", "mm") # 线性回归

        fig, ax = plt.subplots()  # 新建绘图对象

        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(data["k"], data["x"], "o", color='r', markersize=3) # 绘制数据点
        ax.plot(data["k"], res_lsm.b + res_lsm.m*data["k"], color='b', linewidth=1.5) # 拟合直线
        # 作图，详见 https://www.runoob.com/matplotlib/matplotlib-marker.html 和 https://www.runoob.com/matplotlib/matplotlib-line.html
        ax.set_title("$k$ 和 $x_k$ 的关系曲线", fontproperties=zhfont) # 若有中文，需加fontproperties=zhfont
        ax.set_xlabel("$k$", fontproperties=zhfont)
        ax.set_ylabel("$x_k/mm$", fontproperties=zhfont)
        # 添加标题和轴标签，详见 https://www.runoob.com/matplotlib/matplotlib-label.html

        imgpath=workpath+"img.jpg"
        fig.savefig(imgpath, dpi=300, bbox_inches='tight')
        plt.close()


        res_a=analyse_com("a=L*λ/m",(),(("m",res_lsm.m),("L",data["L"][0]),("λ",0.6328)),"mm")
        if(res_a.ans-data["d"][0]>=0):
            res_b=analyse_com("b=(a-d)/d*100",(),(("d",data["d"][0]),("a",res_a.ans)),"%")
        else:
            res_b=analyse_com("b=(d-a)/d*100",(),(("d",data["d"][0]),("a",res_a.ans)),"%")

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph("单缝衍射") # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_picture(imgpath) # 在Word文档中添加图片
        docu.add_paragraph()
        docu.add_paragraph("斜率")
        docu.add_paragraph()._element.append(latex_to_word(res_lsm.mx2))
        docu.add_paragraph("截距")
        docu.add_paragraph()._element.append(latex_to_word(res_lsm.bx2))


        docu.add_paragraph("测得缝宽")
        docu.add_paragraph()._element.append(latex_to_word(res_a.ansx2))
        docu.add_paragraph("相对误差")
        docu.add_paragraph()._element.append(latex_to_word(res_b.ansx2))
        docu.add_paragraph()

        docu.add_paragraph("【Latex代码】")
        docu.add_paragraph()
        docu.add_paragraph("斜率")
        docu.add_paragraph(res_lsm.mx)
        docu.add_paragraph("截距")
        docu.add_paragraph(res_lsm.bx)


        docu.add_paragraph("测得缝宽")
        docu.add_paragraph(res_a.ansx)
        docu.add_paragraph("相对误差")
        docu.add_paragraph(res_b.ansx)
        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致

        os.remove(imgpath)

        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1