from head import *


def name():
    return '硅光电池开路电压、短路电流与光照特性测量'


def handle(workpath, extension):
    try:
        excelpath = workpath+name()+'.'+extension
        if extension == 'csv':
            with open(excelpath, 'rb') as f:
                encode = chardet.detect(f.read())["encoding"]
            data = pd.read_csv(excelpath, header=0, names=["U", "U_of_I"], encoding=encode)
        else:
            data = pd.read_excel(excelpath, header=0, names=["U", "U_of_I"])
        os.remove(excelpath)

        #   数据处理
        U = np.asarray([data['U'][i] for i in range(len(data['U']))])
        U_of_I = np.asarray([data['U_of_I'][i] for i in range(len(data['U_of_I']))])
        I = U_of_I/50
        L = np.asarray([250, 160, 111.1, 81.6, 62.5, 49.4, 40])
        d = np.asarray([20, 25, 30, 35, 40, 45, 50])

        #   拟合曲线
        def func_U(x, a, b):
            return a*np.log(b*x+1)
        popt, pcov = scipy.optimize.curve_fit(func_U, L, U, maxfev=1000)
        a_U = popt[0]
        b_U = popt[1]
        L_fit = np.linspace(np.min(L) - 10, np.max(L) + 10)
        U_fit = func_U(L_fit, a_U, b_U)

        def func_I(x, a, b):
            return a * x + b
        popt, pcov = scipy.optimize.curve_fit(func_I, L, I, maxfev=1000)
        a_I = popt[0]
        b_I = popt[1]
        I_fit = func_I(L_fit, a_I, b_I)

        #   开始绘制曲线
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf")
        imgpath = workpath + "1.jpg"

        fig, ax = plt.subplots()  # 新建绘图对象
        ax_U = ax

        ax_U.xaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        ax_U.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半

        ax_U.plot(L, U, 'o', color='b', markersize=3)
        ax_U.plot(L_fit, U_fit, color='b', linewidth=1.5, label=r'$U_{oc}$' + ' - L')
        ax_U.set_title('硅光电池开路电压、短路电流与光照特性曲线', fontproperties=zhfont)
        ax_U.set_xlabel("L/lx")
        ax_U.set_ylabel('U/V')
        ax_U.legend(loc='upper left')

        #   绘制双轴
        ax_I = ax_U.twinx()
        ax_I.yaxis.set_minor_locator(matplotlib.ticker.AutoMinorLocator(2))
        # 设置副刻度为主刻度的一半
        ax_I.plot(L, I, 'x', color='r', markersize=4)
        ax_I.plot(L_fit, I_fit, color='r', linewidth=1.5, label=r'$I_{sc}$' + ' - L')
        ax_I.set_ylabel('I/mA')
        ax_I.legend(loc='lower right')

        fig.savefig(imgpath, dpi=300, bbox_inches='tight')
        plt.close()
        # ax_I.remove()  # 删除双轴，否则影响以后的图

        #   写入文件
        docu = Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

        docu.add_paragraph('硅光电池开路电压、短路电流与光照特性曲线')
        docu.add_picture(imgpath)

        docu.add_paragraph('开路电压 ' + 'Uoc' + '/V 与光照强度 L/lx 的函数关系为：')
        docu.add_paragraph('U = ' + '{:.5f}'.format(a_U) + ' · ln( ' + '{:.4f}'.format(b_U) + ' · L + 1 )')
        docu.add_paragraph('短路电流 ' + 'Isc' + '/mA 与光照强度 L/lx 的函数关系为：')
        docu.add_paragraph('I = {:.5f} · L'.format(a_I))

        docu.add_paragraph('光照强度L的大小和距离d的关系：')
        table = docu.add_table(rows=2, cols=8, style="Table Grid")  # 在Word文档中插入表格
        table.cell(0, 0).text = 'd/cm'
        for i in range(7):
            table.cell(0, 1 + i).text = str(d[i])
        table.cell(1, 0).text = 'L/lx'
        for i in range(7):
            table.cell(1, 1 + i).text = str(L[i])

        docu.save(workpath+name()+".docx")

        os.remove(imgpath)

        return 0  # succeeded
    except:
        traceback.print_exc()  # 打印错误
        return 1  # failed
