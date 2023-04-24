from head import * # 导入万能头

def name(): # 返回实验名称
    return "表面张力（测量弹簧劲度系数）"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["砝码总质量m/g", "弹簧长度l/cm"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["砝码总质量m/g", "弹簧长度l/cm"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        g = 9.7947
        data["砝码重量G/N"]=[n/1000*g for n in data["砝码总质量m/g"]]
        data["弹簧长度l/m"]=[m/100 for m in data["弹簧长度l/cm"]]

        res_lsm=analyse_lsm(data["弹簧长度l/m"], data["砝码重量G/N"], "F", "l", "N/m", "N") # 线性回归

        fig, ax=plt.subplots() # 新建绘图对象

        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(data["弹簧长度l/m"], data["砝码重量G/N"], "o", color='r', markersize=3) # 绘制数据点
        ax.plot(data["弹簧长度l/m"], res_lsm.b + res_lsm.m*data["弹簧长度l/m"], color='b', linewidth=1.5) # 拟合直线
        # 作图，详见 https://www.runoob.com/matplotlib/matplotlib-marker.html 和 https://www.runoob.com/matplotlib/matplotlib-line.html
        ax.set_title("砝码重量和弹簧长度 G-l 关系曲线", fontproperties=zhfont) # 若有中文，需加fontproperties=zhfont
        ax.set_xlabel("弹簧长度 l/m", fontproperties=zhfont)
        ax.set_ylabel("砝码重量 G/N", fontproperties=zhfont)
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

        docu.add_paragraph("砝码总质量 m 和弹簧长度 l 的关系：")
        table = docu.add_table(rows=len(data["砝码总质量m/g"])+1, cols=3, style="Table Grid") # 在Word文档中插入表格
        table.cell(0,0).text = '弹簧长度l/cm'
        table.cell(0,1).text = '砝码总质量m/g'
        table.cell(0,2).text = '砝码重量G/N'
        for i in range(len(data["砝码总质量m/g"])):
            table.cell(i+1,0).text = ('%.5g' % data["弹簧长度l/cm"][i])
            table.cell(i+1,1).text = ('%.5g' % data["砝码总质量m/g"][i])
            table.cell(i+1,2).text = ('%.5g' % data["砝码重量G/N"][i])
        docu.add_paragraph()

        docu.add_paragraph("最小二乘法拟合：")
        docu.add_picture(imgpath) # 在Word文档中添加图片
        insert_data_lsm(docu, res_lsm, "word")

        docu.add_paragraph("弹簧的劲度系数为：" + '%.5g' % res_lsm.m + ' N/m')
        docu.add_paragraph()

        docu.add_paragraph("【Latex代码】")

        docu.add_paragraph("最小二乘法拟合：")
        docu.add_picture(imgpath) # 在Word文档中添加图片
        insert_data_lsm(docu, res_lsm, "latex")

        docu.add_paragraph("弹簧的劲度系数为：" + '%.5g' % res_lsm.m + ' N/m')

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致
        
        os.remove(imgpath) # 删除刚才保存的图像
    
        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1
