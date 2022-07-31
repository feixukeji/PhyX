from head import * # 导入万能头

def name(): # 返回实验名称
    return "平衡测量法测油滴带电量"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["t1","U1","t2","U2","t3","U3"], encoding=encode) # 读取csv文件

        else:
            data=pd.read_excel(excelpath, header=0, names=["t1","U1","t2","U2","t3","U3"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        p=0.95
        res_U1=analyse(data["U1"],1,0,'U','V')
        res_t1=analyse(data["t1"],0.2,0,'t','s')
        res_U2=analyse(data["U2"],1,0,'U','V')
        res_t2=analyse(data["t2"],0.2,0,'t','s')
        res_U3=analyse(data["U3"],1,0,'U','V')
        res_t3=analyse(data["t3"],0.2,0,'t','s')

        res_q1=analyse_com("q=1.429*10**-14/(U*(t*(1+0.0196*sqrt(t)))**1.5)",(("U",res_U1.average,res_U1.unc),("t",res_t1.average,res_t1.unc)),(),"C")
        res_q2=analyse_com("q=1.429*10**-14/(U*(t*(1+0.0196*sqrt(t)))**1.5)",(("U",res_U2.average,res_U2.unc),("t",res_t2.average,res_t2.unc)),(),"C")
        res_q3=analyse_com("q=1.429*10**-14/(U*(t*(1+0.0196*sqrt(t)))**1.5)",(("U",res_U3.average,res_U3.unc),("t",res_t3.average,res_t3.unc)),(),"C")

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()

        docu.add_paragraph("第一个油滴")

        insert_data(docu, "平衡电压", res_U1, "word")
        insert_data(docu, "下落时间", res_t1, "word")

        docu.add_paragraph("带电量")
        docu.add_paragraph()._element.append(latex_to_word(res_q1.ansx2))
        docu.add_paragraph("带电量")
        docu.add_paragraph()._element.append(latex_to_word(res_q1.uncx2))
        docu.add_paragraph("带电量")
        docu.add_paragraph()._element.append(latex_to_word(res_q1.finalx2))
        docu.add_paragraph()

        docu.add_paragraph("第二个油滴")
        
        insert_data(docu, "平衡电压", res_U2, "word")
        insert_data(docu, "下落时间", res_t2, "word")

        docu.add_paragraph("带电量")
        docu.add_paragraph()._element.append(latex_to_word(res_q2.ansx2))
        docu.add_paragraph("带电量")
        docu.add_paragraph()._element.append(latex_to_word(res_q2.uncx2))
        docu.add_paragraph("带电量")
        docu.add_paragraph()._element.append(latex_to_word(res_q2.finalx2))
        docu.add_paragraph()

        docu.add_paragraph("第三个油滴")
        insert_data(docu, "平衡电压", res_U3, "word")
        insert_data(docu, "下落时间", res_t3, "word")

        docu.add_paragraph("带电量")
        docu.add_paragraph()._element.append(latex_to_word(res_q3.ansx2))
        docu.add_paragraph("带电量")
        docu.add_paragraph()._element.append(latex_to_word(res_q3.uncx2))
        docu.add_paragraph("带电量")
        docu.add_paragraph()._element.append(latex_to_word(res_q3.finalx2))
        docu.add_paragraph()

        docu.add_paragraph("【Latex代码】")

        docu.add_paragraph("第一个油滴")

        insert_data(docu, "平衡电压", res_U1, "latex")
        insert_data(docu, "下落时间", res_t1, "latex")

        docu.add_paragraph("带电量")
        docu.add_paragraph(res_q1.ansx2)
        docu.add_paragraph("带电量")
        docu.add_paragraph(res_q1.uncx2)
        docu.add_paragraph("带电量")
        docu.add_paragraph(res_q1.finalx2)
        docu.add_paragraph()

        docu.add_paragraph("第二个油滴")
        
        insert_data(docu, "平衡电压", res_U2, "latex")
        insert_data(docu, "下落时间", res_t2, "latex")

        docu.add_paragraph("带电量")
        docu.add_paragraph(res_q2.ansx2)
        docu.add_paragraph("带电量")
        docu.add_paragraph(res_q2.uncx2)
        docu.add_paragraph("带电量")
        docu.add_paragraph(res_q2.finalx2)
        docu.add_paragraph()

        docu.add_paragraph("第三个油滴")
        insert_data(docu, "平衡电压", res_U3, "latex")
        insert_data(docu, "下落时间", res_t3, "latex")

        docu.add_paragraph("带电量")
        docu.add_paragraph(res_q3.ansx2)
        docu.add_paragraph("带电量")
        docu.add_paragraph(res_q3.uncx2)
        docu.add_paragraph("带电量")
        docu.add_paragraph(res_q3.finalx2)
        docu.add_paragraph()

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致
    
        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1
