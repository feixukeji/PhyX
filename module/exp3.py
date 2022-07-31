from head import * # 导入万能头

def name(): # 返回实验名称
    return "落球法测定液体的粘度"

def handle(workpath,extension):
    # 处理数据并生成文档，workpath为工作文件夹路径（本程序涉及到的所有文件只能保存在此文件夹内），extension为扩展名（csv/xls/xlsx）
    try:
        zhfont = matplotlib.font_manager.FontProperties(fname="SourceHanSansSC-Regular.otf") # 设置图像中的文字字体

        excelpath=workpath+name()+'.'+extension # Excel文件名（含路径），文件名与name()函数返回值一致

        if extension=='csv':
            with open(excelpath,'rb') as f:
                encode=chardet.detect(f.read())["encoding"] # 判断编码格式
            data=pd.read_csv(excelpath, header=0, names=["rho","rho0","g","h","l","D","d","t"], encoding=encode) # 读取csv文件
        else:
            data=pd.read_excel(excelpath, header=0, names=["rho","rho0","g","h","l","D","d","t"]) # 读取xls/xlsx文件

        os.remove(excelpath) # 读取Excel数据后删除文件

        rho=data["rho"][0]*1000
        rho0=data["rho0"][0]*1000
        g=data["g"][0]
        p=0.95
        
        res_h=analyse(data["h"],0.02,0.05,'h','cm')
        res_l=analyse(data["l"],0.02,0.05,'l','cm')
        res_D=analyse(data["D"],0.02,0,'D','mm',confidence_C=3**0.5)
        res_d=analyse(data["d"],0.02,0,'d','mm',confidence_C=3**0.5)
        res_t=analyse(data["t"],0.01,0.2,'t','s')

        res_v=analyse_com("v=l/t",(),(("l",res_l.average/100),("t",res_t.average)),"m/s")
        v=res_v.ans

        d=res_d.average/1000
        D=res_D.average/1000
        h=res_h.average/100

        res_eta0=analyse_com("eta0=1/18*(rho-rho0)*g*d**2/(v*(1+2.4*d/D)*(1+3.3*d/2/h))",(),(("rho",rho),("rho0",rho0),("g",g),("d",d),("v",v),("D",D),("h",h)),"Pa·s")
        eta0=res_eta0.ans

        res_Re=analyse_com("Re=v*rho0*d/eta0",(),(("v",v),("rho0",rho0),("d",d),("eta0",eta0)),"")
        Re=res_Re.ans

        res_eta1=analyse_com("eta1=eta0-3/16*d*v*rho0",(),(("eta0",eta0),("d",d),("v",v),("rho0",rho0)),"Pa·s")
        eta1=res_eta1.ans

        res_eta2=analyse_com("eta2=1/2*eta1*(1+sqrt(1+19/270*(d*v*rho0/eta1)**2))",(),(("eta1",eta1),("d",d),("v",v),("rho0",rho0),("eta1",eta1)),"Pa·s")
        eta2=res_eta2.ans

        docu=Document()
        docu.styles['Normal'].font.name = '微软雅黑'
        docu.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑') # 设置Word文档字体

        docu.add_paragraph(name()) # 在Word文档中添加文字
        docu.add_paragraph()
        docu.add_paragraph("【Latex代码在下面，请向下翻阅】")
        docu.add_paragraph()
        
        insert_data(docu, "液面高度h", res_h, "word")
        insert_data(docu, "匀速下降区l", res_l, "word")
        insert_data(docu, "量筒直径D", res_D, "word")
        insert_data(docu, "小球直径d", res_d, "word")
        insert_data(docu, "下落时间t", res_t, "word")

        docu.add_paragraph("小球下落速度")
        docu.add_paragraph()._element.append(latex_to_word(res_v.ansx2))
        docu.add_paragraph("粘度的零级近似值")
        docu.add_paragraph()._element.append(latex_to_word(res_eta0.ansx2))
        docu.add_paragraph("雷诺数")
        docu.add_paragraph()._element.append(latex_to_word(res_Re.ansx2))

        if Re<0.1:
            docu.add_paragraph("Re<0.1，无需修正")
            docu.add_paragraph("粘度 η=η0="+('%.5g'%eta0)+" Pa·s")
        else:
            if Re<0.5:
                docu.add_paragraph("0.1<Re<0.5，进行一级修正：")
                docu.add_paragraph("粘度")
                docu.add_paragraph()._element.append(latex_to_word(res_eta1.ansx2))
            else:
                docu.add_paragraph("Re>0.5，进行二级修正：")
                docu.add_paragraph()._element.append(latex_to_word(res_eta1.ansx2))
                docu.add_paragraph("粘度")
                docu.add_paragraph()._element.append(latex_to_word(res_eta2.ansx2))

        docu.add_paragraph()
        
        docu.add_paragraph("【Latex代码】")

        insert_data(docu, "液面高度h", res_h, "latex")
        insert_data(docu, "匀速下降区l", res_l, "latex")
        insert_data(docu, "量筒直径D", res_D, "latex")
        insert_data(docu, "小球直径d", res_d, "latex")
        insert_data(docu, "下落时间t", res_t, "latex")

        docu.add_paragraph("小球下落速度")
        docu.add_paragraph(res_v.ansx)
        docu.add_paragraph("粘度的零级近似值")
        docu.add_paragraph(res_eta0.ansx)
        docu.add_paragraph("雷诺数")
        docu.add_paragraph(res_Re.ansx)
        
        if Re<0.1:
            docu.add_paragraph("Re<0.1，无需修正")
            docu.add_paragraph("粘度 η=η0="+('%.5g'%eta0)+" Pa·s")
        else:
            if Re<0.5:
                docu.add_paragraph("0.1<Re<0.5，进行一级修正：")
                docu.add_paragraph("粘度")
                docu.add_paragraph(res_eta1.ansx)
            else:
                docu.add_paragraph("Re>0.5，进行二级修正：")
                docu.add_paragraph(res_eta1.ansx)
                docu.add_paragraph("粘度")
                docu.add_paragraph(res_eta2.ansx)

        docu.save(workpath+name()+".docx") # 保存Word文档，注意文件名必须与name()函数返回值一致
    
        return 0 # 若成功，返回0
    except:
        traceback.print_exc() # 打印错误
        return 1 # 若失败，返回1