from head import * # 导入万能头

def name(): # 返回实验名称
    return "测量弹簧参数"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["L","x","y"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["L","x","y"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        l1=len(data["x"])
        l2=len(data["y"])
        a=[0]*(l1-1)
        b=[0]*(l2-1)
        i=0
        for i in range (0,l1-1):
            a[i]=data["x"][i+1]-data["x"][i]
        
        for i in range (0,l2-1):
            b[i]=data["y"][i+1]-data["y"][i]

        res_L=analyse(data["L"],1,0.5,"L","mm")
        res_x=analyse(data["x"],0.01,0.004,"d1","mm")
        res_y=analyse(data["y"],0.01,0.004,"d2","mm")

        res_dx=analyse_com("d1=L*λ/x",(("L",res_L.average,res_L.unc),("x",res_x.average,res_x.unc)),(("λ",0.6328),),"μm")
        res_dy=analyse_com("d2=L*λ/y",(("L",res_L.average,res_L.unc),("y",res_y.average,res_y.unc)),(("λ",0.6328),),"μm")

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()

        insert_data(docu, "距离L", res_L, "word")
        insert_data(docu, "较窄亮条纹宽度x", res_x, "word")
        insert_data(docu, "较宽亮条纹宽度y", res_y, "word")

        docu.add_paragraph("弹簧细丝直径")
        docu.add_paragraph()._element.append(latex_to_word(res_dx.ansx2))
        docu.add_paragraph("弹簧细丝直径d1的延伸不确定度")
        docu.add_paragraph()._element.append(latex_to_word(res_dx.uncx2))
        docu.add_paragraph("弹簧细丝直径最终结果")
        docu.add_paragraph()._element.append(latex_to_word(res_dx.finalx2))
        docu.add_paragraph()

        docu.add_paragraph("弹簧螺距")
        docu.add_paragraph()._element.append(latex_to_word(res_dy.ansx2))
        docu.add_paragraph("弹簧螺距d2的延伸不确定度")
        docu.add_paragraph()._element.append(latex_to_word(res_dy.uncx2))
        docu.add_paragraph("弹簧螺距最终结果")
        docu.add_paragraph()._element.append(latex_to_word(res_dy.finalx2))
        docu.add_paragraph()

        docu.add_paragraph("【Latex代码】")
        
        insert_data(docu, "距离L", res_L, "latex")
        insert_data(docu, "较窄亮条纹宽度x", res_x, "latex")
        insert_data(docu, "较宽亮条纹宽度y", res_y, "latex")

        docu.add_paragraph("弹簧细丝直径")
        docu.add_paragraph(res_dx.ansx)
        docu.add_paragraph("弹簧细丝直径d1的延伸不确定度")
        docu.add_paragraph(res_dx.uncx)
        docu.add_paragraph("弹簧细丝直径最终结果")
        docu.add_paragraph(res_dx.finalx)
        docu.add_paragraph()

        docu.add_paragraph("弹簧螺距")
        docu.add_paragraph(res_dy.ansx)
        docu.add_paragraph("弹簧螺距d2的延伸不确定度")
        docu.add_paragraph(res_dy.uncx)
        docu.add_paragraph("弹簧螺距最终结果")
        docu.add_paragraph(res_dy.finalx)
        docu.add_paragraph()

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致
        
        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1
