from head import * # 导入万能头

def name(): # 返回实验名称
    return "测量金属圆柱体的密度"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["d","h","m"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["d","h","m"]) # 读取xls/xlsx文件
        
        os.remove(excelpath) # 读取Excel数据后删除文件

        res_E=analyse_com("ρ=4*m/(pi*d*d*h)",(),(("m",data["m"][0]),("d",data["d"]),("h",data["h"])),"g/cm^3")

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph("测量金属圆柱体的密度") # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("密度")
        docu.add_paragraph()._element.append(latex_to_word(res_E.ansx2))
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码】")
        docu.add_paragraph(res_E.ansx)

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致

        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1