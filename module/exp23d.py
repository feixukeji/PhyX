from head import * # 导入万能头
from scipy.signal import savgol_filter # 数据平滑去噪

def name(): # 返回实验名称
    return "光电效应测伏安特性曲线"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["U_365","I_365","U_405","I_405","U_436","I_436","U_546","I_546","U_577","I_577"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["U_365","I_365","U_405","I_405","U_436","I_436","U_546","I_546","U_577","I_577"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        fig, ax = plt.subplots()  # 新建绘图对象
        ax.plot(data["U_365"], data["I_365"], "o", color='b', markersize=3)
        data["I_365"] = savgol_filter(data["I_365"], 7, 4)
        ax.plot(data["U_365"], data["I_365"], color='b', markersize=1.5, label="λ=365.0nm")
        ax.plot(data["U_405"], data["I_405"], "s", color='r', markersize=3)
        data["I_405"] = savgol_filter(data["I_405"], 7, 4)
        ax.plot(data["U_405"], data["I_405"], color='r', markersize=1.5, label="λ=404.7nm")
        ax.plot(data["U_436"], data["I_436"], "^", color='g', markersize=3)
        data["I_436"] = savgol_filter(data["I_436"], 7, 4)
        ax.plot(data["U_436"], data["I_436"], color='g', markersize=1.5, label="λ=435.8nm")
        ax.plot(data["U_546"], data["I_546"], "D", color='m', markersize=3)
        data["I_546"] = savgol_filter(data["I_546"], 7, 4)
        ax.plot(data["U_546"], data["I_546"], color='m', markersize=1.5, label="λ=546.1nm")
        ax.plot(data["U_577"], data["I_577"], "v", color='orange', markersize=3)
        data["I_577"] = savgol_filter(data["I_577"], 7, 4)
        ax.plot(data["U_577"], data["I_577"], color='orange', markersize=1.5, label="λ=577.0nm")
        ax.legend()
        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(5))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(5))

        ax.set_title("光电管的伏安特性曲线", fontproperties=zhfont)
        ax.set_xlabel("Collector Voltage (V)")
        ax.set_ylabel("Current in Phototube (nA)")

        imgpath=workpath+"1.jpg"
        fig.savefig(imgpath, dpi=300, bbox_inches='tight')
        plt.close()

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph("光电管的伏安特性曲线：")
        docu.add_picture(imgpath)
        docu.add_paragraph("从上图中观察出5条曲线的拐点，记录其对应的电压与电流，再利用“近代：光电效应测普朗克常量”实验工具即可获得用“拐点法”测得的普朗克常数。")

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致

        os.remove(imgpath)

        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1