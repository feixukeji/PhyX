from head import * # 导入万能头

def name(): # 返回实验名称
    return "用转动定律测量物体的质量"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["r","t","L","r0","t0","m"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["r","t","L","r0","t0","m"]) # 读取xls/xlsx文件
        
        os.remove(excelpath) # 读取Excel数据后删除文件

        data["r"]/=100
        data["t"]/=30
        data["L"]/=100
        data["r0"]/=100
        data["t0"]/=30
        data["m"]/=1000

        for i in range (0,5)  :
            x=data["r"][i]
            y=data["t"][i]
            data["r"][i]=x*x
            data["t"][i]=9.8*x*y*y/(4*math.pi*math.pi)

        res_lsm=analyse_lsm(data["r"], data["t"], "r", "t", "m^{-1}·s^{-2}", "m·s^{-2}") # 线性回归

        fig, ax=plt.subplots() # 新建绘图对象

        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(data["r"], data["t"], "o", color='r', markersize=3) # 绘制数据点
        ax.plot(data["r"], res_lsm.b + res_lsm.m*data["r"], color='b', linewidth=1.5) # 拟合直线
        # 作图，详见 https://www.runoob.com/matplotlib/matplotlib-marker.html 和 https://www.runoob.com/matplotlib/matplotlib-line.html
        ax.set_title("$rT^2$ 和 $r^2$ 的关系曲线", fontproperties=zhfont) # 若有中文，需加fontproperties=zhfont
        ax.set_xlabel("$r^2/m^2$", fontproperties=zhfont)
        ax.set_ylabel("$rT^2/(m\cdot s^2)$", fontproperties=zhfont)
        # 添加标题和轴标签，详见 https://www.runoob.com/matplotlib/matplotlib-label.html
        
        imgpath=workpath+"img.jpg"
        fig.savefig(imgpath, dpi=300, bbox_inches='tight') # 保存生成的图像

        
        res_ic=analyse_com("IC=b*m",(),(("b",res_lsm.b),("m",data["m"][0])),"kg·m^2")
        res_m=analyse_com("M=IC/(g*r*t*t/(4*pi*pi)-L*L/12-r*r)",(),(("IC",res_ic.ans),("r",data["r0"][0]),("t",data["t0"][0]),("L",data["L"][0]),("g",9.8)),"kg")

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph("用转动定律测量物体的质量") # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_picture(imgpath) # 在Word文档中添加图片
        docu.add_paragraph()
        docu.add_paragraph("斜率")
        docu.add_paragraph()._element.append(latex_to_word(res_lsm.mx2))
        docu.add_paragraph("截距")
        docu.add_paragraph()._element.append(latex_to_word(res_lsm.bx2))

        
        docu.add_paragraph("转动惯量")
        docu.add_paragraph()._element.append(latex_to_word(res_ic.ansx2))
        docu.add_paragraph("物体的质量")
        docu.add_paragraph()._element.append(latex_to_word(res_m.ansx2))
        docu.add_paragraph()

        docu.add_paragraph("【Latex代码】")
        docu.add_paragraph()
        docu.add_paragraph("斜率")
        docu.add_paragraph(res_lsm.mx)
        docu.add_paragraph("截距")
        docu.add_paragraph(res_lsm.bx)

        
        docu.add_paragraph("转动惯量")
        docu.add_paragraph(res_ic.ansx)
        docu.add_paragraph("物体的质量")
        docu.add_paragraph(res_m.ansx)

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致

        os.remove(imgpath)
    
        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1