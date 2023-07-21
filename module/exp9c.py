from head import * # 导入万能头

def name(): # 返回实验名称
    return "时差法测有机玻璃棒和黄铜棒中声速"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["L_plexiglass", "t_plexiglass", "L_brass", "t_brass"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["L_plexiglass", "t_plexiglass", "L_brass", "t_brass"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        res_plexiglass=analyse_lsm(data["t_plexiglass"], data["L_plexiglass"], 't', 'L', 'cm/μs', 'cm')
        ax.clear()
        fig_plexiglass, ax_plexiglass = fig, ax
        ax_plexiglass.plot(data["t_plexiglass"], data["L_plexiglass"], "o", color='r', markersize=3)
        ax_plexiglass.plot(data["t_plexiglass"], res_plexiglass.b + res_plexiglass.m*data["t_plexiglass"], color='b', linewidth=1.5)
        ax_plexiglass.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax_plexiglass.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax_plexiglass.set_title("时差法测有机玻璃棒中声速的最小二乘法拟合图", fontproperties=zhfont)
        ax_plexiglass.set_xlabel("Propagation Time of Sound Waves in Plexiglass Rod (µs)")
        ax_plexiglass.set_ylabel("Length of Plexiglass Rod (cm)")

        imgpath_plexiglass=workpath+"plexiglass.jpg"
        fig_plexiglass.savefig(imgpath_plexiglass, dpi=300, bbox_inches='tight')

        res_brass=analyse_lsm(data["t_brass"], data["L_brass"], 't', 'L', 'cm/μs', 'cm')
        ax.clear()
        fig_brass, ax_brass = fig, ax
        ax_brass.plot(data["t_brass"], data["L_brass"], "o", color='r', markersize=3)
        ax_brass.plot(data["t_brass"], res_brass.b + res_brass.m*data["t_brass"], color='b', linewidth=1.5)
        ax_brass.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax_brass.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax_brass.set_title("时差法测黄铜棒中声速的最小二乘法拟合图", fontproperties=zhfont)
        ax_brass.set_xlabel("Propagation Time of Sound Waves in Brass Rod (µs)")
        ax_brass.set_ylabel("Length of Brass Rod (cm)")

        imgpath_brass=workpath+"brass.jpg"
        fig_brass.savefig(imgpath_brass, dpi=300, bbox_inches='tight')

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()

        docu.add_paragraph("时差法测有机玻璃棒中声速")
        docu.add_picture(imgpath_plexiglass)
        insert_data_lsm(docu, res_plexiglass, option="word")
        docu.add_paragraph("v_有机玻璃棒 = 斜率m = "+'%.5g'%(10000*res_plexiglass.m)+" m/s")
        docu.add_paragraph()
        docu.add_paragraph("时差法测黄铜棒中声速")
        docu.add_picture(imgpath_brass)
        insert_data_lsm(docu, res_brass, option="word")
        docu.add_paragraph("v_黄铜棒 = 斜率m = "+'%.5g'%(10000*res_brass.m)+" m/s")
        docu.add_paragraph()

        docu.add_paragraph("【Latex代码】")

        docu.add_paragraph("时差法测有机玻璃棒中声速")
        insert_data_lsm(docu, res_plexiglass, option="latex")
        docu.add_paragraph("时差法测黄铜棒中声速")
        insert_data_lsm(docu, res_brass, option="latex")

        docu.save(workpath+name()+".docx")

        os.remove(imgpath_plexiglass)
        os.remove(imgpath_brass)

        return 0
    except:
        traceback.print_exc() # 打印错误
        return 1