from head import * # 导入万能头

def name(): # 返回实验名称
    return "切变模量"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["D","d1","d2","L","m","T0","n0","T1","n1"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["D","d1","d2","L","m","T0","n0","T1","n1"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        data["T0"]/=data["n0"][0]
        data["T1"]/=data["n1"][0]

        res_d=analyse(data["D"],0.01,0.005,"d","mm")
        res_d1=analyse(data["d1"],0.02,0,"d1","mm",3**0.5)
        res_d2=analyse(data["d2"],0.02,0,"d2","mm",3**0.5)
        res_L=analyse(data["L"],0.1,0.05,"L","cm")
        res_m=analyse(data["m"],1,0.5,"m","g")
        res_t0=analyse(data["T0"],0.0005,0.01,"T0","s")
        res_t1=analyse(data["T1"],0.0005,0.01,"T1","s")

        res_D=analyse_com("D=(pi**2*m*(d1**2+d2**2))/(2*(t1**2-t0**2))",(("m",res_m.average/1000,res_m.unc/1000),("d1",res_d1.average/1000,res_d1.unc/1000),("d2",res_d2.average/1000,res_d2.unc/1000),("t1",res_t1.average,res_t1.unc),("t0",res_t0.average,res_t0.unc)),(),"kg·m^2/s^2")
        res_G=analyse_com("G=(16*pi*L*m*(d1*d1+d2*d2))/((d**4)*(t1*t1-t0*t0))",(("L",res_L.average/100,res_L.unc/100),("m",res_m.average/1000,res_m.unc/1000),("d1",res_d1.average/1000,res_d1.unc/1000),("d2",res_d2.average/1000,res_d2.unc/1000),("d",res_d.average/1000,res_d.unc/1000),("t1",res_t1.average,res_t1.unc),("t0",res_t0.average,res_t0.unc)),(),"kg/(m·s^2)")

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()

        insert_data(docu, "钢丝直径d", res_d, "word")
        insert_data(docu, "环内直径d1", res_d1, "word")
        insert_data(docu, "环外直径d2", res_d2, "word")
        insert_data(docu, "钢丝长度L", res_L, "word")
        insert_data(docu, "圆环质量m", res_m, "word")
        insert_data(docu, "周期T0", res_t0, "word")
        insert_data(docu, "周期T1", res_t1, "word")

        docu.add_paragraph("扭转模量")
        docu.add_paragraph()._element.append(latex_to_word(res_D.ansx2))
        docu.add_paragraph("扭转模量D的延伸不确定度")
        docu.add_paragraph()._element.append(latex_to_word(res_D.uncx2))
        docu.add_paragraph("扭转模量最终结果")
        docu.add_paragraph()._element.append(latex_to_word(res_D.finalx2))
        docu.add_paragraph()

        docu.add_paragraph("切变模量")
        docu.add_paragraph()._element.append(latex_to_word(res_G.ansx2))
        docu.add_paragraph("切变模量G的延伸不确定度")
        docu.add_paragraph()._element.append(latex_to_word(res_G.uncx2))
        docu.add_paragraph("切变模量最终结果")
        docu.add_paragraph()._element.append(latex_to_word(res_G.finalx2))
        docu.add_paragraph()

        docu.add_paragraph("【Latex代码】")

        insert_data(docu, "钢丝直径d", res_d, "latex")
        insert_data(docu, "环内直径d1", res_d1, "latex")
        insert_data(docu, "环外直径d2", res_d2, "latex")
        insert_data(docu, "钢丝长度L", res_L, "latex")
        insert_data(docu, "圆环质量m", res_m, "latex")
        insert_data(docu, "周期T0", res_t0, "latex")
        insert_data(docu, "周期T1", res_t1, "latex")

        docu.add_paragraph("扭转模量")
        docu.add_paragraph(res_D.ansx)
        docu.add_paragraph("扭转模量D的延伸不确定度")
        docu.add_paragraph(res_D.uncx)
        docu.add_paragraph("扭转模量最终结果")
        docu.add_paragraph(res_D.finalx)
        docu.add_paragraph()

        docu.add_paragraph("切变模量")
        docu.add_paragraph(res_G.ansx)
        docu.add_paragraph("切变模量G的延伸不确定度")
        docu.add_paragraph(res_G.uncx)
        docu.add_paragraph("切变模量最终结果")
        docu.add_paragraph(res_G.finalx)
        docu.add_paragraph()

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致

        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1
