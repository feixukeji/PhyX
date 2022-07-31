from head import * # 导入万能头

def name(): # 返回实验名称
    return "匀加速运动验证牛顿第二定律"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["b","mn","t1","t2","t3"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["b","mn","t1","t2","t3"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        m=float(data["b"][0])
        me=float(data["b"][2])
        s=float(data["b"][4])/100
        delta_s=float(data["b"][6])
        data.dropna(inplace = True)
        data["拉力F/N"]=data["mn"]/1000*9.8
        data["通过光电门时的瞬时速度v(m/s)"]=delta_s/((data["t1"]+data["t2"]+data["t3"])/3)
        data["通过光电门时的瞬时加速度a(m/s^2)"]=data["通过光电门时的瞬时速度v(m/s)"]**2/(2*s)
        
        res=analyse_lsm(data["通过光电门时的瞬时加速度a(m/s^2)"], data["拉力F/N"], 'a', 'F', 'kg', 'N') # 最小二乘多项式拟合之线性回归

        fig, ax=plt.subplots() # 新建绘图对象

        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(data["通过光电门时的瞬时加速度a(m/s^2)"], data["拉力F/N"], "o", color='r', markersize=3) # 绘制数据点
        ax.plot(data["通过光电门时的瞬时加速度a(m/s^2)"], res.b + res.m*data["通过光电门时的瞬时加速度a(m/s^2)"], color='b', linewidth=1.5) # 拟合直线
        # 作图，详见 https://www.runoob.com/matplotlib/matplotlib-marker.html 和 https://www.runoob.com/matplotlib/matplotlib-line.html
        ax.set_title("拉力与通过光电门时的瞬时加速度 $F_n-a_n$ 关系曲线", fontproperties=zhfont) # 若有中文，需加fontproperties=zhfont
        ax.set_xlabel("通过光电门时的瞬时加速度 $a_n\\rm{(m/s^2)}$", fontproperties=zhfont)
        ax.set_ylabel("拉力 $F_n\\rm{(N)}$", fontproperties=zhfont)
        # 添加标题和轴标签，详见 https://www.runoob.com/matplotlib/matplotlib-label.html
        
        imgpath=workpath+"img.jpg"
        fig.savefig(imgpath, dpi=300, bbox_inches='tight') # 保存生成的图像

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        
        docu.add_paragraph("拉力与通过光电门时的瞬时加速度的关系：")
        table = docu.add_table(rows=len(data["拉力F/N"])+1, cols=3, style="Table Grid") # 在Word文档中插入表格
        table.cell(0,0).text = '拉力F/N'
        table.cell(0,1).text = '通过光电门时的瞬时速度v(m/s)'
        table.cell(0,2).text = '通过光电门时的瞬时加速度a(m/s^2)'
        for i in range(len(data["拉力F/N"])):
            table.cell(i+1,0).text = ('%.5g' % data["拉力F/N"][i])
            table.cell(i+1,1).text = ('%.5g' % data["通过光电门时的瞬时速度v(m/s)"][i])
            table.cell(i+1,2).text = ('%.5g' % data["通过光电门时的瞬时加速度a(m/s^2)"][i])
        docu.add_paragraph()

        docu.add_picture(imgpath)
        docu.add_paragraph("斜率 M="+('%.5g' % res.m)+" kg="+('%.5g' % (res.m*1000))+" g")

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致
        
        os.remove(imgpath) # 删除刚才保存的图像
    
        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1