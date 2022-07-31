from head import * # 导入万能头

def name(): # 返回实验名称
    return "模拟太空失重环境用动力学方法测量物体的质量"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["t1","t2","t3","m"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["t1","t2","t3","m"]) # 读取xls/xlsx文件
        
        os.remove(excelpath) # 读取Excel数据后删除文件

        res_m=analyse_com("M=m/(t2*t2/(t1*t1)-1)*(t3*t3/(t1*t1)-1)-m",(),(("m",data["m"][0]),("t1",data["t1"][0]),("t2",data["t2"][0]),("t3",data["t3"][0])),"g")

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph("模拟太空失重环境用动力学方法测量物体的质量") # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("质量")
        docu.add_paragraph()._element.append(latex_to_word(res_m.ansx2))
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码】")
        docu.add_paragraph(res_m.ansx)

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致
    
        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1