from head import * # 导入万能头

def name(): # 返回实验名称
    return "表面张力（用金属丝测量）"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["k", "d", "l0", "l"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["k", "d", "l0", "l"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        k = data["k"][0]
        l0 = data["l0"][0]
        data["l"] -= l0
        res_d=analyse(data["d"], 0.01, 0.005, "d", "cm")
        res_l=analyse(data["l"], 0.01, 0.005, "Δl", "cm")

        res_sigma=analyse_com("σ=(k*Δl)/(2*d)",(("Δl",res_l.average,res_l.unc),("d",res_d.average,res_d.unc)),(("k",k),),"N/m")

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()

        insert_data(docu, "铁丝长度d", res_d, "word")
        insert_data(docu, "弹簧伸长量Δl", res_l, "word")

        docu.add_paragraph("表面张力σ")
        docu.add_paragraph()._element.append(latex_to_word(res_sigma.ansx2))
        docu.add_paragraph("表面张力σ的延伸不确定度")
        docu.add_paragraph()._element.append(latex_to_word(res_sigma.uncx2))
        docu.add_paragraph("表面张力σ最终结果")
        docu.add_paragraph()._element.append(latex_to_word(res_sigma.finalx2))

        docu.add_paragraph()
        docu.add_paragraph("【Latex代码】")
        docu.add_paragraph()
        
        insert_data(docu, "铁丝长度d", res_d, "latex")
        insert_data(docu, "弹簧伸长量Δl", res_l, "latex")

        docu.add_paragraph("表面张力σ")
        docu.add_paragraph(res_sigma.ansx)
        docu.add_paragraph("表面张力σ的延伸不确定度")
        docu.add_paragraph(res_sigma.uncx)
        docu.add_paragraph("表面张力σ最终结果")
        docu.add_paragraph(res_sigma.finalx)
        docu.add_paragraph()

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致
    
        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1
