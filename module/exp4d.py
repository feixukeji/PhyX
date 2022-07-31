from head import * # 导入万能头

def name(): # 返回实验名称
    return "液体法测量物体密度"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["m","m0","p"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["m","m0","p"]) # 读取xls/xlsx文件
        
        os.remove(excelpath) # 读取Excel数据后删除文件

        res_m=analyse_com("ρ=p*m/M",(),(("p",data["p"][0]),("m",data["m"][0]),("M",data["m0"][0])),"g/cm^3")

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph("液体法测量物体密度") # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("密度")
        docu.add_paragraph()._element.append(latex_to_word(res_m.ansx2))
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码】")
        docu.add_paragraph(res_m.ansx)

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致

        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1