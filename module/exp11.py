from head import * # 导入万能头
from scipy.signal import savgol_filter # 使用 Savitzky-Golay 平滑去噪

def name(): # 返回实验名称
    return "半导体温度计"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["I"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["I"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        data["T"] = [20, 25, 30, 35, 40, 45, 50, 55, 60, 65, 70]

        fig, ax = plt.subplots()  # 新建绘图对象
        ax.plot(data["T"], data["I"], "o", color='r', markersize=3, label="Bridge Current")
        data["I"] = savgol_filter(data["I"], 5, 2)
        ax.plot(data["T"], data["I"], color='b', markersize=1.5, label="5 pts SG smooth of \"Bridge Current\"")
        ax.legend(loc="upper left")
        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))

        ax.set_title("半导体温度计电桥电流与温度的关系", fontproperties=zhfont)
        ax.set_xlabel("Temperature (°C)")
        ax.set_ylabel("Bridge Current (μA)")

        imgpath=workpath+"img.jpg"
        fig.savefig(imgpath, dpi=300, bbox_inches='tight')
        plt.close()

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph("半导体温度计电桥电流与温度的关系：")
        docu.add_picture(imgpath)
        docu.add_paragraph()

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致

        os.remove(imgpath)

        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1