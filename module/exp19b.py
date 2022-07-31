from head import * # 导入万能头

def name(): # 返回实验名称
    return "干涉法测微小量（测头发丝直径）"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["name","1","2","3"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["name","1","2","3"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        dataTurn = [data["1"], data["2"], data["3"]]
        data = pd.DataFrame({
            'l2': np.array([x[0] for x in dataTurn]),
            'l1': np.array([x[1] for x in dataTurn]),
            'L0': np.array([x[2] for x in dataTurn]),
            'Ln': np.array([x[3] for x in dataTurn]),
        }) # 在转置的同时，由于第一列不是数字，转换成的 DataFrame 带了前缀，暂未找到可以在read_csv/read_excel参数上调整的方法，因此这样转换。

        lmbda = 589.3 * 1e-6 # 毫米

        dl = data['l2'] - data['l1']
        arrL = data['Ln'] - data['L0']
        res_dl = analyse(dl, 0.0012, 0.0005, '{l_{20\\text{条}}}', 'mm', 3 ** .5, 0.68)
        res_L = analyse(arrL, 0.0012, 0.0005, '{L_{\\text{总长}}}', 'mm', 3 ** .5, 0.68)
        # n = 20 / res_dl.average
        # d = res_L.average * n * lmbda / 2
        res_d = analyse_com("D=L*(20/dl)*λ/2",(("dl",res_dl.average,res_dl.unc), ("L",res_L.average,res_L.unc)),(("λ",lmbda),),('mm'),0.68)
        # print(d) 

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()

        insert_data(docu, "20根干涉条纹长度", res_dl, "word")
        insert_data(docu, "总干涉条纹长度长度", res_L, "word")
        docu.add_paragraph("头发丝直径长度")
        docu.add_paragraph()._element.append(latex_to_word(res_d.ansx2))
        docu.add_paragraph("头发丝直径长度的延伸不确定度")
        docu.add_paragraph()._element.append(latex_to_word(res_d.uncx2))
        docu.add_paragraph("头发丝直径长度最终结果")
        docu.add_paragraph()._element.append(latex_to_word(res_d.finalx2))
        docu.add_paragraph()

        docu.add_paragraph("【Latex代码】")

        insert_data(docu, "20根干涉条纹长度", res_dl, "latex")
        insert_data(docu, "总干涉条纹长度长度", res_L, "latex")
        docu.add_paragraph("头发丝直径长度")
        docu.add_paragraph(res_d.ansx)
        docu.add_paragraph("头发丝直径长度的延伸不确定度")
        docu.add_paragraph(res_d.uncx)
        docu.add_paragraph("头发丝直径长度最终结果")
        docu.add_paragraph(res_d.finalx)
        docu.add_paragraph()

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致
    
        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1