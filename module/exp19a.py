from head import * # 导入万能头

def name(): # 返回实验名称
    return "干涉法测微小量（测曲率半径）"

def array_to_str(arr, eps = 3):
    return '[' + ', '.join('%.4lf'%a for a in arr) + ']'

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["name","d5","d10","d15","d20","d25","d30"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["name","d5","d10","d15","d20","d25","d30"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        dataTurn = [data["d5"], data["d10"], data["d15"], data["d20"], data["d25"], data["d30"]]
        d1_R = np.array([x[0] for x in dataTurn])
        d1_L = np.array([x[1] for x in dataTurn])
        d2_R = np.array([x[2] for x in dataTurn])
        d2_L = np.array([x[3] for x in dataTurn])
        d3_R = np.array([x[4] for x in dataTurn])
        d3_L = np.array([x[5] for x in dataTurn])

        lmbda = 589.3 * 1e-6 # 毫米
        cnt = 6

        D1, D2, D3 = d1_L - d1_R, d2_L - d2_R, d3_L - d3_R
        D = (D1 + D2 + D3) / 3
        # print(D)
        D2 = D ** 2
        # print(D2) 
        n = cnt // 2
        R = (D2[:-n] - D2[n:]) / (4 * (5 * n) * lmbda)
        # print(R, sum(R) / (cnt - n))

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph("干涉法测微小量（测曲率半径）") # 在Word文档中添加文字
        docu.add_paragraph("平均直径 D_{5,10,15,20,25,30}="+array_to_str(D)+" mm")
        docu.add_paragraph("求出的曲率半径 (5~20, 10~25, 15~30)="+array_to_str(R)+" mm")
        docu.add_paragraph("平均曲率半径 "+'%.4lf'%(sum(R) / (cnt - n))+" mm")

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致
    
        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1