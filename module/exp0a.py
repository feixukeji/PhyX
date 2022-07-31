from head import *  # 导入万能头


def name():  # 返回实验名称
    return "标准差和不确定度计算"


def handle(workpath, extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        excelpath = workpath+name()+'.'+extension  # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension == 'csv':
            with open(excelpath, 'rb') as f:
                encode = chardet.detect(f.read())["encoding"]  # 判断编码格式
            data = pd.read_csv(excelpath, header=0, names=["x", "delta_b1", "delta_b2", "confidence_C", "confidence_P"], encoding=encode)  # 读取csv文件
        else:
            data = pd.read_excel(excelpath, header=0, names=["x", "delta_b1", "delta_b2", "confidence_C", "confidence_P"])  # 读取xls/xlsx文件

        os.remove(excelpath)  # 读取Excel数据后删除文件

        res = analyse(data["x"], data["delta_b1"][0], data["delta_b2"][0], confidence_C=data["confidence_C"][0], confidence_P=data["confidence_P"][0])

        docu = Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')  # 设置Word文档字体

        docu.add_paragraph(name())  # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()

        insert_data(docu, "x", res, "word")

        docu.add_paragraph("【Latex代码】")
        insert_data(docu, "x", res, "latex")

        docu.save(workpath+name()+".docx")  # 保存Word文档，注意文件名必须与name()函数返回值一致

        return 0  # 若成功，返回0
    except:
        traceback.print_exc()  # 打印错误
        return 1  # 若失败，返回1
