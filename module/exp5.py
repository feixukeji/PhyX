from head import * # 导入万能头

def name(): # 返回实验名称
    return "用拉伸法测量钢丝的杨氏模量"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["D","l","L","d","n","b1","b2"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["D","l","L","d","n","b1","b2"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        res_D=analyse(data["D"], 0.12, 0.05, "D", "cm")
        res_l=analyse(data["l"], 0.12, 0.05, "l", "cm")
        res_L=analyse(data["L"], 0.12, 0.05, "L", "cm")
        res_d=analyse(data["d"], 0.004, 0.005, "d", "mm")

        data["砝码总质量m/g"]=[n*500 for n in data["n"]]
        data["金属丝受拉力F/N"]=[m/1000*9.8 for m in data["砝码总质量m/g"]]
        data["标尺读数平均值b/cm"]=[(data["b1"][i]+data["b2"][i])/2 for i in range(len(data["b1"]))]

        res_lsm=analyse_lsm(data["金属丝受拉力F/N"], data["标尺读数平均值b/cm"], "F", "b", "cm/N", "cm") # 线性回归

        res_E=analyse_com("E=(8*D*L)/(pi*d**2*l*m)",(("D",res_D.average,res_D.unc),("L",res_L.average,res_L.unc),("d",res_d.average/10,res_d.unc/10),("l",res_l.average,res_l.unc),("m",abs(res_lsm.m),res_lsm.u_m)),(),"N/cm^2")

        ax.clear()

        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(data["金属丝受拉力F/N"], data["标尺读数平均值b/cm"], "o", color='r', markersize=3) # 绘制数据点
        ax.plot(data["金属丝受拉力F/N"], res_lsm.b + res_lsm.m*data["金属丝受拉力F/N"], color='b', linewidth=1.5) # 拟合直线
        # 作图，详见 https://www.runoob.com/matplotlib/matplotlib-marker.html 和 https://www.runoob.com/matplotlib/matplotlib-line.html
        ax.set_title("标尺读数与金属丝受拉力 b-F 关系曲线", fontproperties=zhfont) # 若有中文，需加fontproperties=zhfont
        ax.set_xlabel("金属丝受拉力 F/N", fontproperties=zhfont)
        ax.set_ylabel("标尺读数 b/cm", fontproperties=zhfont)
        # 添加标题和轴标签，详见 https://www.runoob.com/matplotlib/matplotlib-label.html

        imgpath=workpath+"img.jpg"
        fig.savefig(imgpath, dpi=300, bbox_inches='tight') # 保存生成的图像

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()

        insert_data(docu, "标尺到平面镜的距离D", res_D, "word")
        insert_data(docu, "光杠杆的臂长l", res_l, "word")
        insert_data(docu, "钢丝原长L", res_L, "word")
        insert_data(docu, "钢丝直径d", res_d, "word")

        docu.add_paragraph("金属丝受拉力F与标尺读数b的关系：")
        table = docu.add_table(rows=len(data["砝码总质量m/g"])+1, cols=3, style="Table Grid") # 在Word文档中插入表格
        table.cell(0,0).text = '砝码总质量m/g'
        table.cell(0,1).text = '金属丝受拉力F/N'
        table.cell(0,2).text = '标尺读数平均值b/cm'
        for i in range(len(data["砝码总质量m/g"])):
            table.cell(i+1,0).text = ('%.5g' % data["砝码总质量m/g"][i])
            table.cell(i+1,1).text = ('%.5g' % data["金属丝受拉力F/N"][i])
            table.cell(i+1,2).text = ('%.5g' % data["标尺读数平均值b/cm"][i])
        docu.add_paragraph()

        docu.add_paragraph("最小二乘法拟合：")
        docu.add_picture(imgpath) # 在Word文档中添加图片
        insert_data_lsm(docu, res_lsm, "word")

        docu.add_paragraph("杨氏模量")
        docu.add_paragraph()._element.append(latex_to_word(res_E.ansx2))
        docu.add_paragraph("杨氏模量E的延伸不确定度")
        docu.add_paragraph()._element.append(latex_to_word(res_E.uncx2))
        docu.add_paragraph("杨氏模量最终结果")
        docu.add_paragraph()._element.append(latex_to_word(res_E.finalx2))
        docu.add_paragraph()

        docu.add_paragraph("【Latex代码】")

        insert_data(docu, "标尺到平面镜的距离D", res_D, "latex")
        insert_data(docu, "光杠杆的臂长l", res_l, "latex")
        insert_data(docu, "钢丝原长L", res_L, "latex")
        insert_data(docu, "钢丝直径d", res_d, "latex")

        docu.add_paragraph("金属丝受拉力F与标尺读数b的关系：")
        table = docu.add_table(rows=len(data["砝码总质量m/g"])+1, cols=3, style="Table Grid") # 在Word文档中插入表格
        table.cell(0,0).text = '砝码总质量m/g'
        table.cell(0,1).text = '金属丝受拉力F/N'
        table.cell(0,2).text = '标尺读数平均值b/cm'
        for i in range(len(data["砝码总质量m/g"])):
            table.cell(i+1,0).text = ('%.5g' % data["砝码总质量m/g"][i])
            table.cell(i+1,1).text = ('%.5g' % data["金属丝受拉力F/N"][i])
            table.cell(i+1,2).text = ('%.5g' % data["标尺读数平均值b/cm"][i])
        docu.add_paragraph()

        docu.add_paragraph("最小二乘法拟合：")
        docu.add_picture(imgpath) # 在Word文档中添加图片
        insert_data_lsm(docu, res_lsm, "latex")

        docu.add_paragraph("杨氏模量")
        docu.add_paragraph(res_E.ansx)
        docu.add_paragraph("杨氏模量E的延伸不确定度")
        docu.add_paragraph(res_E.uncx)
        docu.add_paragraph("杨氏模量最终结果")
        docu.add_paragraph(res_E.finalx)
        docu.add_paragraph()

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致

        os.remove(imgpath) # 删除刚才保存的图像

        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1
