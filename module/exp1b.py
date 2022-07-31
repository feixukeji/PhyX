from head import * # 导入万能头

def name(): # 返回实验名称
    return "单摆法测重力加速度"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["l","d","T","n"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["l","d","T","n"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        data["T"]/=data["n"][0] # Pandas Series支持整体运算，相当于数组中的每个元素都做同样的操作

        res_l=analyse(data["l"],0.2,0.05,'l','cm') # 摆线长度的相关值计算
        res_d=analyse(data["d"],0.02,0,'d','mm',confidence_C=3**0.5) # 摆球直径的相关值计算
        res_T=analyse(data["T"],0.01/data["n"][0],0.2/data["n"][0],'T','s') # 周期的相关值计算

        res_L=analyse_com("L=l+d",(("l",res_l.average,res_l.unc),("d",res_d.average/20,res_d.unc/20)),(),"cm")
        res_g=analyse_com("g=4*pi**2*L/T**2",(("L",res_L.ans/100,res_L.unc/100),("T",res_T.average,res_T.unc)),(),"m/s^2")
    
        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()

        docu.add_paragraph("钢卷尺的最大允差为0.2cm，游标卡尺的最大允差为0.002cm，秒表的最大允差为0.01s")
        docu.add_paragraph("钢卷尺和游标卡尺的估计误差为最小分度值的一半，分别为0.05cm和0.001cm")
        docu.add_paragraph("秒表的估计误差为0.2s")
        docu.add_paragraph()

        insert_data(docu, "摆线长度l", res_l, "word")
        insert_data(docu, "摆球直径d", res_d, "word")

        docu.add_paragraph("摆长L")
        docu.add_paragraph()._element.append(latex_to_word(res_L.ansx2))
        docu.add_paragraph("摆长L的延伸不确定度")
        docu.add_paragraph()._element.append(latex_to_word(res_L.uncx2))

        insert_data(docu, "周期T", res_T, "word")

        docu.add_paragraph("重力加速度g")
        docu.add_paragraph()._element.append(latex_to_word(res_g.ansx2))
        docu.add_paragraph("重力加速度g的延伸不确定度")
        docu.add_paragraph()._element.append(latex_to_word(res_g.uncx2))
        docu.add_paragraph("重力加速度g最终结果")
        docu.add_paragraph()._element.append(latex_to_word(res_g.finalx2))
        docu.add_paragraph()

        docu.add_paragraph("【Latex代码】")
        
        insert_data(docu, "摆线长度l", res_l, "latex")
        insert_data(docu, "摆球直径d", res_d, "latex")

        docu.add_paragraph("摆长L")
        docu.add_paragraph(res_L.ansx)
        docu.add_paragraph("摆长L的延伸不确定度")
        docu.add_paragraph(res_L.uncx)

        insert_data(docu, "周期T", res_T, "latex")

        docu.add_paragraph("重力加速度g")
        docu.add_paragraph(res_g.ansx)
        docu.add_paragraph("重力加速度g的延伸不确定度")
        docu.add_paragraph(res_g.uncx)
        docu.add_paragraph("重力加速度g最终结果")
        docu.add_paragraph(res_g.finalx)

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致

        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1
