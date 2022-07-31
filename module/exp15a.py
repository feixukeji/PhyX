from head import *


def name():
    return '硅光电池暗伏安特性测量'


def handle(workpath, extension):
    try:
        excelpath = workpath+name()+'.'+extension

        if extension == 'csv':
            with open(excelpath, 'rb') as f:
                encode = chardet.detect(f.read())["encoding"]
            data = pd.read_csv(excelpath, header=0, names=["U", "I"], encoding=encode)
        else:
            data = pd.read_excel(excelpath, header=0, names=["U", "I"])

        os.remove(excelpath)

        U = [data['U'][i] for i in range(len(data['U']))]
        I = [data['I'][i] for i in range(len(data['I']))]

        #   拟合曲线
        def func(x, a, b):
            return a*(np.exp(b*x)-1)
        popt, pcov = scipy.optimize.curve_fit(func, U, I, maxfev=1000)
        a = popt[0]
        b = popt[1]
        x_fit = np.linspace(np.min(U), np.max(U))
        y_fit = func(x_fit, a, b)

        #   开始绘制I-V曲线
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf")
        
        imgpath = workpath + "1.jpg"

        fig, ax = plt.subplots()  # 新建绘图对象

        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(data['U'], data['I'], "o", color='r', markersize=3)
        ax.plot(x_fit, y_fit, color='b', linewidth=1.5)
        ax.set_title('无光照正向偏压时硅光电池的 I-U 特性曲线', fontproperties=zhfont)
        ax.set_xlabel("U/V")
        ax.set_ylabel('I/mA')
        fig.savefig(imgpath, dpi=300, bbox_inches='tight')

        docu = Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

        docu.add_paragraph('无光照正向偏压时硅光电池的 I-U 特性曲线')
        docu.add_picture(imgpath)

        docu.save(workpath+name()+".docx")

        os.remove(imgpath)

        return 0  # succeeded
    except:
        traceback.print_exc()  # 打印错误
        return 1  # failed
