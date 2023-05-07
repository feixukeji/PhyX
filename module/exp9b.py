from head import * # 导入万能头

def name(): # 返回实验名称
    return "相位比较法测水中声速"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["L","f"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["L","f"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        data["n"]=range(1, data["L"].size+1)
        res_lsm=analyse_lsm(data["n"], data["L"], "n", "L", "cm", "cm")

        res_f=analyse(data["f"], 0.001, 10, "f", "Hz")

        fig, ax=plt.subplots() # 新建绘图对象

        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(data["n"], data["L"], "o", color='r', markersize=3)
        ax.plot(data["n"], res_lsm.b + res_lsm.m * data["n"], color='b', linewidth=1.5)
        ax.set_title("共振干涉法测空气中声速的最小二乘法拟合图", fontproperties=zhfont)
        ax.set_xlabel("n")
        ax.set_ylabel("The nth Position (cm)")
        
        imgpath=workpath+"img.jpg"
        fig.savefig(imgpath, dpi=300, bbox_inches='tight') # 保存生成的图像

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()

        docu.add_paragraph("最小二乘法拟合：")
        docu.add_picture(imgpath) # 在Word文档中添加图片

        insert_data_lsm(docu, res_lsm, option="word")
        lamda=analyse_com("λ=2*m", (("m", abs(res_lsm.m), res_lsm.u_m),), (), "cm")
        docu.add_paragraph("波长λ")
        docu.add_paragraph()._element.append(latex_to_word(lamda.ansx2))
        docu.add_paragraph("波长λ的延伸不确定度")
        docu.add_paragraph()._element.append(latex_to_word(lamda.uncx2))
        docu.add_paragraph()
        
        docu.add_paragraph("谐振频率的不确定度")
        docu.add_paragraph()._element.append(latex_to_word(res_f.delta_bx2))
        docu.add_paragraph()
        v=analyse_com("v=f*λ", (("λ", lamda.ans/100, lamda.unc/100), ("f", res_f.average, res_f.unc)), (), "m/s")
        docu.add_paragraph("声速v")
        docu.add_paragraph()._element.append(latex_to_word(v.ansx2))
        docu.add_paragraph("声速v的延伸不确定度")
        docu.add_paragraph()._element.append(latex_to_word(v.uncx2))
        docu.add_paragraph("声速v最终结果")
        docu.add_paragraph()._element.append(latex_to_word(v.finalx2))
        docu.add_paragraph()
        
        docu.add_paragraph("【Latex代码】")

        insert_data_lsm(docu, res_lsm, option="latex")
        docu.add_paragraph("波长λ")
        docu.add_paragraph(lamda.ansx)
        docu.add_paragraph("波长λ的延伸不确定度")
        docu.add_paragraph(lamda.uncx)
        docu.add_paragraph()
        docu.add_paragraph("谐振频率的不确定度")
        docu.add_paragraph(res_f.delta_bx)
        docu.add_paragraph()
        docu.add_paragraph("声速v")
        docu.add_paragraph(v.ansx)
        docu.add_paragraph("声速v的延伸不确定度")
        docu.add_paragraph(v.uncx)
        docu.add_paragraph("声速v最终结果")
        docu.add_paragraph(v.finalx)
        docu.add_paragraph()
        
        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致
        
        os.remove(imgpath) # 删除刚才保存的图像
    
        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1