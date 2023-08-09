from head import * # 导入万能头

def name(): # 返回实验名称
    return "匀变速运动中速度与加速度的测量"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["b","s","t1","t2","t3"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["b","s","t1","t2","t3"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        delta_s=float(data["b"][0])
        h=float(data["b"][2])
        L=float(data["b"][4])
        data.dropna(inplace = True)
        data["两倍距离2s(m)"]=data["s"]/50
        data["速度平方v^2(m^2/s^2)"]=(delta_s/((data["t1"]+data["t2"]+data["t3"])/3))**2

        res_lsm=analyse_lsm(data["两倍距离2s(m)"], data["速度平方v^2(m^2/s^2)"], 'X', 'Y', 'm/s^2', 'm^2/s^2') # 线性回归

        res_g=analyse_com("g=m*L/h",(),(("m",res_lsm.m),("L",L),("h",h)),"m/s^2")

        fig, ax = plt.subplots()  # 新建绘图对象

        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(data["两倍距离2s(m)"], data["速度平方v^2(m^2/s^2)"], "o", color='r', markersize=3) # 绘制数据点
        ax.plot(data["两倍距离2s(m)"], res_lsm.b + res_lsm.m*data["两倍距离2s(m)"], color='b', linewidth=1.5) # 拟合直线
        # 作图，详见 https://www.runoob.com/matplotlib/matplotlib-marker.html 和 https://www.runoob.com/matplotlib/matplotlib-line.html
        ax.set_title("速度平方与两倍距离 $v^2-2s$ 关系曲线", fontproperties=zhfont) # 若有中文，需加fontproperties=zhfont
        ax.set_xlabel("两倍距离 $2s\\rm{(m)}$", fontproperties=zhfont)
        ax.set_ylabel("速度平方 $v^2\\rm{(m^2/s^2)}$", fontproperties=zhfont)
        # 添加标题和轴标签，详见 https://www.runoob.com/matplotlib/matplotlib-label.html

        imgpath=workpath+"img.jpg"
        fig.savefig(imgpath, dpi=300, bbox_inches='tight')
        plt.close()

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()

        docu.add_paragraph("速度平方v^2与两倍距离2s的关系：")
        table = docu.add_table(rows=len(data["两倍距离2s(m)"])+1, cols=2, style="Table Grid") # 在Word文档中插入表格
        table.cell(0,0).text = '两倍距离2s(m)'
        table.cell(0,1).text = '速度平方v^2(m^2/s^2)'
        for i in range(len(data["两倍距离2s(m)"])):
            table.cell(i+1,0).text = ('%.5g' % data["两倍距离2s(m)"][i])
            table.cell(i+1,1).text = ('%.5g' % data["速度平方v^2(m^2/s^2)"][i])
        docu.add_paragraph()

        docu.add_paragraph("最小二乘法拟合：")
        docu.add_picture(imgpath) # 在Word文档中添加图片
        insert_data_lsm(docu, res_lsm, "word")

        docu.add_paragraph("重力加速度")
        docu.add_paragraph()._element.append(latex_to_word(res_g.ansx2))
        docu.add_paragraph()

        docu.add_paragraph("【Latex代码】")

        insert_data_lsm(docu, res_lsm, "latex")
        docu.add_paragraph("重力加速度")
        docu.add_paragraph(res_g.ansx)

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致

        os.remove(imgpath) # 删除刚才保存的图像

        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1