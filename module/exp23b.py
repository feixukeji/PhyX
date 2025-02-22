from head import * # 导入万能头

def name(): # 返回实验名称
    return "光电效应验证饱和光电流与光强（光阑孔面积）成正比"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["I_436","I_546"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["I_436","I_546"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        data["Phi_squared"] = [4, 16, 64, 205.9225]

        res1=analyse_lsm(data["Phi_squared"], data["I_436"], 'Φ', 'I', 'nA/mm', 'nA')
        res2=analyse_lsm(data["Phi_squared"], data["I_546"], 'Φ', 'I', 'nA/mm', 'nA')
        fig, ax = plt.subplots()  # 新建绘图对象
        ax.plot(data["Phi_squared"], data["I_436"], "^", color='g', markersize=3)
        ax.plot(data["Phi_squared"], res1.b + res1.m*data["Phi_squared"], color='g', markersize=1.5, label="λ=435.8nm")
        ax.plot(data["Phi_squared"], data["I_546"], "v", color='m', markersize=3)
        ax.plot(data["Phi_squared"], res2.b + res2.m*data["Phi_squared"], color='m', markersize=1.5, label="λ=546.1nm")
        ax.legend()
        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.set_title("饱和光电流与光阑孔直径的平方的关系", fontproperties=zhfont)
        ax.set_xlabel("Square of Length of Diaphragm's Diameter (mm$^2$)")
        ax.set_ylabel("Saturation Current (nA)")

        imgpath=workpath+"img.jpg"
        fig.savefig(imgpath, dpi=300, bbox_inches='tight')
        plt.close()

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph("光电效应验证饱和光电流与光强（光阑孔面积）成正比") # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_picture(imgpath) # 在Word文档中添加图片
        docu.add_paragraph("光波长为 435.8 nm 的线性拟合相关系数 r = %.8g"%res1.r)
        docu.add_paragraph()._element.append(latex_to_word(res1.rx2))
        docu.add_paragraph("[Latex 代码]")
        docu.add_paragraph(res1.rx)
        docu.add_paragraph("光波长为 546.1 nm 的线性拟合相关系数 r = %.8g"%res2.r)
        docu.add_paragraph()._element.append(latex_to_word(res2.rx2))
        docu.add_paragraph("[Latex 代码]")
        docu.add_paragraph(res2.rx)

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致

        os.remove(imgpath)

        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1