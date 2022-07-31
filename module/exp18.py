from numpy import average
from head import * # 导入万能头

def name(): # 返回实验名称
    return "分光计的调节与使用"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["#1",'t1','t2','t1-','t2-','#2'], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["#1",'t1','t2','t1-','t2-','#2']) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        # data['t1'] = np.array(list(data['t1']))
        # data['t2'] = np.array(list(data['t2']))
        # data['t1-'] = np.array(list(data['t1-']))
        # data['t2-'] = np.array(list(data['t2-']))
        data = pd.DataFrame({
            't1': np.array(data['t1']),
            't2': np.array(data['t2']),
            't1-': np.array(data['t1-']),
            't2-': np.array(data['t2-']),
        }) # 由于第一列不是数字，转换成的 DataFrame 带了前缀，暂未找到可以在read_csv/read_excel参数上调整的方法，因此这样转换。
        data['t1'] = np.floor(data['t1']) + (data['t1'] - np.floor(data['t1'])) * 100 / 60
        data['t2'] = np.floor(data['t2']) + (data['t2'] - np.floor(data['t2'])) * 100 / 60
        data['t1-'] = np.floor(data['t1-']) + (data['t1-'] - np.floor(data['t1-'])) * 100 / 60
        data['t2-'] = np.floor(data['t2-']) + (data['t2-'] - np.floor(data['t2-'])) * 100 / 60
        if data['t2'][0] < data['t1'][0]:
            data['t2'] += 360
        if data['t1-'][0] < data['t1'][0]:
            data['t1-'] += 360
        if data['t2-'][0] < data['t2'][0]:
            data['t2-'] += 360
        if average(180 - (data['t1-'][:3] + data['t2-'][:3] - data['t1'][:3] - data['t2'][:3]) / 2) < 0:
            data['t1'], data['t1-'] = data['t1-'], data['t1']
            data['t2'], data['t2-'] = data['t2-'], data['t2']
            if data['t2'][0] < data['t1'][0]:
                data['t2'] += 360
            if data['t1-'][0] < data['t1'][0]:
                data['t1-'] += 360
            if data['t2-'][0] < data['t2'][0]:
                data['t2-'] += 360
        data['A'] = 180 - (data['t1-'][:3] + data['t2-'][:3] - data['t1'][:3] - data['t2'][:3]) / 2
        data['green'] = 180 - (data['t1-'][3:] + data['t2-'][3:] - data['t1'][3:] - data['t2'][3:]) / 2

        res_A=analyse(data["A"], 1/60, 0.5/60, "A", "°", confidence_C=3**.5)
        res_green=analyse(data["green"], 1/60, 0.5/60, "{\delta_{\min}}", "°", confidence_C=3**.5)
        res_n=analyse_com("n=sin((A+green)/2)/sin(A/2)",(("A",res_A.average*np.pi/180,res_A.unc*np.pi/180),("green",res_green.average*np.pi/180,res_green.unc*np.pi/180)),())

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()

        docu.add_paragraph("顶角：" + ', '.join(['%.5g °' % x for x in data["A"][:3]]))
        insert_data(docu, "顶角", res_A, "word")
        docu.add_paragraph("绿光最小偏向角：" + ', '.join(['%.5g °' % x for x in data["green"][3:]]))
        insert_data(docu, "绿光最小偏向角", res_green, "word")

        docu.add_paragraph("绿光下玻璃三棱镜折射率n")
        docu.add_paragraph()._element.append(latex_to_word(res_n.ansx2))
        docu.add_paragraph("绿光下玻璃三棱镜折射率n的延伸不确定度")
        docu.add_paragraph()._element.append(latex_to_word(res_n.uncx2))
        docu.add_paragraph("绿光下玻璃三棱镜折射率n最终结果")
        docu.add_paragraph()._element.append(latex_to_word(res_n.finalx2))
        docu.add_paragraph()

        docu.add_paragraph("【Latex代码】")

        docu.add_paragraph("顶角：" + ', '.join(['%.5g °' % x for x in data["A"][:3]]))
        insert_data(docu, "顶角", res_A, "latex")
        docu.add_paragraph("绿光最小偏向角：" + ', '.join(['%.5g °' % x for x in data["green"][3:]]))
        insert_data(docu, "绿光最小偏向角", res_green, "latex")

        docu.add_paragraph("绿光下玻璃三棱镜折射率n")
        docu.add_paragraph(res_n.ansx)
        docu.add_paragraph("绿光下玻璃三棱镜折射率n的延伸不确定度")
        docu.add_paragraph(res_n.uncx)
        docu.add_paragraph("绿光下玻璃三棱镜折射率n最终结果")
        docu.add_paragraph(res_n.finalx)
        docu.add_paragraph()

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致
        
        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1
