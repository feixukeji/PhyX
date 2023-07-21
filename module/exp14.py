from head import * # 导入万能头

def name(): # 返回实验名称
    return "直流电源特性"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["R", "U_DC", "U_AC"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["R", "U_DC", "U_AC"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        data["P"] = data["U_DC"]*data["U_DC"]/data["R"]*1000
        data["K"] = data["U_AC"]/data["U_DC"]

        fig_P, ax_P = plt.subplots()
        ax_P.plot(data["R"], data["P"], "o", color='r', markersize=3)
        ax_P.plot(data["R"], data["P"], color='b', linewidth=1.5)
        ax_P.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax_P.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax_P.set_title("输出功率和负载 $P$–$R$ 曲线", fontproperties=zhfont)
        ax_P.set_xlabel("Load Resistance ($\Omega$)")
        ax_P.set_ylabel("Output Power (mW)")

        imgpath_P=workpath+"Power.jpg"
        fig_P.savefig(imgpath_P, dpi=300, bbox_inches='tight')

        fig_K, ax_K=plt.subplots()
        ax_K.plot(data["R"], data["K"], "o", color='r', markersize=3)
        ax_K.plot(data["R"], data["K"], color='b', linewidth=1.5)
        ax_K.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax_K.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax_K.set_title("纹波系数和负载 $K$–$R$ 曲线", fontproperties=zhfont)
        ax_K.set_xlabel("Load Resistance ($\Omega$)")
        ax_K.set_ylabel("Ripple Factor")

        imgpath_K=workpath+"ripple.jpg"
        fig_K.savefig(imgpath_K, dpi=300, bbox_inches='tight')

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【本文档不只有一页，请向下翻阅】")
        docu.add_paragraph()
        docu.add_paragraph("输出功率：")
        docu.add_picture(imgpath_P)
        docu.add_paragraph()
        docu.add_paragraph("纹波系数：")
        docu.add_picture(imgpath_K)

        docu.save(workpath+name()+".docx")

        os.remove(imgpath_P)
        os.remove(imgpath_K)

        return 0
    except:
        traceback.print_exc() # 打印错误
        return 1