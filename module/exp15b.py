from head import *


def name():
    return '硅光电池输出特性测量'


def handle(workpath, extension):
    try:
        excelpath = workpath+name()+'.'+extension  # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension == 'csv':
            with open(excelpath, 'rb') as f:
                encode = chardet.detect(f.read())["encoding"]  # 判断编码格式
            data = pd.read_csv(excelpath, header=0, names=["R", "d1", "d2", "d3", "d4"], encoding=encode)  # 读取csv文件
        else:
            data = pd.read_excel(excelpath, header=0, names=["R", "d1", "d2", "d3", "d4"])  # 读取xls/xlsx文件

        os.remove(excelpath)  # 读取Excel数据后删除文件

        n = len(data["R"])
        #   数据处理
        R = data["R"]
        index = ["d1", "d2", "d3", "d4"]
        U = np.asarray([[data[index[i]][j] for j in range(n)] for i in range(4)])
        L = ['250', '111.1', '62.5', '40']
        I = np.asarray([[U[i][j] / R[j] for j in range(n)] for i in range(4)])
        P = U * I
        data = []
        index = []

        #   开始拟合P—R图线

        #   开始画图I-U
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf")

        imgpath1 = workpath + "1.jpg"

        fig, ax = plt.subplots()  # 新建绘图对象

        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(I[0], U[0], 'o-', label='L = 250 lx', markersize=3, linewidth=1.5)
        ax.plot(I[1], U[1], 's-', label='L = 111.1 lx', markersize=3, linewidth=1.5)
        ax.plot(I[2], U[2], '^-', label='L = 62.5 lx', markersize=3, linewidth=1.5)
        ax.plot(I[3], U[3], 'D-', label='L = 40 lx', markersize=3, linewidth=1.5)

        ax.set_title('I - U', fontproperties=zhfont)
        ax.set_xlabel('U/mV')
        ax.set_ylabel('I/mA')
        ax.legend()
        fig.savefig(imgpath1, dpi=300, bbox_inches='tight')
        plt.close()

        #   开始画图P-Rl
        imgpath2 = workpath + '2.jpg'

        P *= 1000  # 换单位

        fig, ax = plt.subplots()  # 新建绘图对象

        ax.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax.plot(R, P[0], 'o-', label='L = 250 lx', markersize=3, linewidth=1.5)
        ax.plot(R, P[1], 's-', label='L = 111.1 lx', markersize=3, linewidth=1.5)
        ax.plot(R, P[2], '^-', label='L = 62.5 lx', markersize=3, linewidth=1.5)
        ax.plot(R, P[3], 'D-', label='L = 40 lx', markersize=3, linewidth=1.5)

        ax.set_title('P - R', fontproperties=zhfont)
        ax.set_xlabel('P/mW')
        ax.set_ylabel('R/' + r'$\Omega$')
        ax.legend()
        fig.savefig(imgpath2, dpi=300, bbox_inches='tight')
        plt.close()

        # 文本内容
        # 电流
        content_I = ''
        for i in range(4):
            content_I = content_I + 'L = ' + L[i] + 'lx :\n'
            for j in range(n):
                content_I += '  %.5g'%I[i][j]
            content_I += '\n'
        # 功率
        content_P = ''
        for i in range(4):
            content_P = content_P + 'L = ' + L[i] + 'lx :\n'
            for j in range(n):
                content_P += '  %.5g'%P[i][j]
            content_P = content_P + '\n'
        # FF因子
        P_max = [np.max(i) for i in P]
        FF = [P_max[i]/1000/I[i][0]/U[i][-1] for i in range(4)]
        content_FF = ''
        for i in range(4):
            content_FF = content_FF + 'L = ' + \
                L[i] + 'lx , FF = %.5g'%FF[i] + '\n'

        #   写入文件
        docu = Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

        docu.add_paragraph('不同光照条件下 I - U 曲线')
        docu.add_picture(imgpath1)

        docu.add_paragraph('对应的I/mA为：')
        docu.add_paragraph(content_I)

        docu.add_paragraph('不同光照条件下 P - R 曲线')
        docu.add_picture(imgpath2)

        docu.add_paragraph('对应的P/mW为：')
        docu.add_paragraph(content_P)
        docu.add_paragraph('填充因子FF：')
        docu.add_paragraph(content_FF)

        docu.save(workpath+name()+".docx")

        os.remove(imgpath1)
        os.remove(imgpath2)

        return 0  # succeeded
    except:
        traceback.print_exc()  # 打印错误
        return 1  # failed
