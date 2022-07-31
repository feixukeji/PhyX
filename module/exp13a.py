from head import * # 导入万能头

def name(): # 返回实验名称
    return "整流滤波（信号源频率对滤波效果的影响）"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["f.Hz","pi.DC","pi.AC","whole.DC","whole.AC"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["f.Hz","pi.DC","pi.AC","whole.DC","whole.AC"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        fig, ax=plt.subplots() # 新建绘图对象
        imgpath=workpath+"1.jpg"

        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(data["f.Hz"], data["pi.DC"], 'o-', color='r', markersize=3, label="π型RC电路") # 绘制数据点
        ax.plot(data["f.Hz"], data["whole.DC"], 'o-', color='b', markersize=3, label="全波整流电路") # 绘制数据点
        ax.set_title("直流电压随频率的变化曲线", fontproperties=zhfont)
        ax.set_xlabel("频率 $f/Hz$", fontproperties=zhfont)
        ax.set_ylabel("$U/V$") # 支持latex
        ax.legend(prop=zhfont)
        fig.savefig(imgpath, dpi=300, bbox_inches='tight')

        fig, ax=plt.subplots() # 新建绘图对象
        imgpath=workpath+"2.jpg"

        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(data["f.Hz"], data["pi.AC"], 'o-', color='r', markersize=3, label="π型RC电路") # 绘制数据点
        ax.plot(data["f.Hz"], data["whole.AC"], 'o-', color='b', markersize=3, label="全波整流电路") # 绘制数据点
        ax.set_title("交流电压随频率的变化曲线", fontproperties=zhfont)
        ax.set_xlabel("频率 $f/Hz$", fontproperties=zhfont)
        ax.set_ylabel("$U/V$") # 支持latex
        ax.legend(prop=zhfont)
        fig.savefig(imgpath, dpi=300, bbox_inches='tight')

        fig, ax=plt.subplots() # 新建绘图对象
        imgpath=workpath+"3.jpg"
        
        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(data["f.Hz"], data["pi.AC"] / data["pi.DC"] * 100, 'o-', color='r', markersize=3, label="π型RC电路") # 绘制数据点
        ax.plot(data["f.Hz"], data["whole.AC"] / data["whole.DC"] * 100, 'o-', color='b', markersize=3, label="全波整流电路") # 绘制数据点
        ax.set_title("纹波系数随频率的变化曲线", fontproperties=zhfont)
        ax.set_xlabel("频率 $f/Hz$", fontproperties=zhfont)
        ax.set_ylabel("纹波系数$\kappa(\%)$", fontproperties=zhfont) # 支持latex
        ax.legend(prop=zhfont)
        fig.savefig(imgpath, dpi=300, bbox_inches='tight')

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph("整流滤波（信号源频率对滤波效果的影响）")
        docu.add_paragraph("内容不只一页，请向下翻阅！\n")
        docu.add_paragraph("κ值（π型RC电路）：" + ', '.join(['%.2lf'%x for x in data["pi.AC"] / data["pi.DC"] * 100]))
        docu.add_paragraph("κ值（全波整流电路）：" + ', '.join(['%.2lf'%x for x in data["whole.AC"] / data["whole.DC"] * 100]))
        docu.add_picture(workpath+"1.jpg") # 在Word文档中添加图片
        docu.add_picture(workpath+"2.jpg") # 在Word文档中添加图片
        docu.add_picture(workpath+"3.jpg") # 在Word文档中添加图片
        

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致
        
        os.remove(imgpath) # 删除刚才保存的图像
    
        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1