import math
import scipy.stats
from sympy import *
from collections import namedtuple
from api.transformer import *
from uncertainties import ufloat

t_P = {0.68: {2: 1.84, 3: 1.32, 4: 1.20, 5: 1.14, 6: 1.11, 7: 1.09, 8: 1.08, 9: 1.07, 10: 1.06, 11: 1.05, 12: 1.05, 13: 1.04, 14: 1.04, 15: 1.04, 16: 1.03, 17: 1.03, 18: 1.03, 19: 1.03, 20: 1.03, 21: 1.03, 26: 1.02, 31: 1.02, 36: 1.01, 41: 1.01, 46: 1.01, 51: 1.01, 101: 1.005, 'inf': 1},
       0.90: {2: 6.31, 3: 2.92, 4: 2.35, 5: 2.13, 6: 2.02, 7: 1.94, 8: 1.89, 9: 1.86, 10: 1.83, 11: 1.81, 12: 1.80, 13: 1.78, 14: 1.77, 15: 1.76, 16: 1.75, 17: 1.75, 18: 1.74, 19: 1.73, 20: 1.73, 21: 1.72, 26: 1.71, 31: 1.70, 36: 1.70, 41: 1.68, 46: 1.68, 51: 1.68, 101: 1.660, 'inf': 1.645},
       0.95: {2: 12.71, 3: 4.30, 4: 3.18, 5: 2.78, 6: 2.57, 7: 2.45, 8: 2.36, 9: 2.31, 10: 2.26, 11: 2.23, 12: 2.20, 13: 2.18, 14: 2.16, 15: 2.14, 16: 2.13, 17: 2.12, 18: 2.11, 19: 2.10, 20: 2.09, 21: 2.09, 26: 2.06, 31: 2.04, 36: 2.03, 41: 2.02, 46: 2.01, 51: 1.68, 101: 1.984, 'inf': 1.960},
       0.99: {2: 63.66, 3: 9.92, 4: 5.84, 5: 4.60, 6: 4.03, 7: 3.71, 8: 3.50, 9: 3.36, 10: 3.25, 11: 3.17, 12: 3.11, 13: 3.05, 14: 3.01, 15: 2.98, 16: 2.95, 17: 2.92, 18: 2.90, 19: 2.88, 20: 2.86, 21: 2.85, 26: 2.79, 31: 2.75, 36: 2.72, 41: 2.70, 46: 2.69, 51: 2.68, 101: 2.626, 'inf': 2.576}}

k_P = {0.50: 0.675, 0.68: 1, 0.90: 1.65, 0.95: 1.96, 0.955: 2, 0.99: 2.58, 0.997: 3}


def adjust_char(latex_code):
    # 将特殊字符转换为Latex代码
    latex_code = latex_code.replace("μ", r"\mu ")
    latex_code = latex_code.replace("ν", r"\nu ")
    latex_code = latex_code.replace("Φ", r"\Phi ")
    latex_code = latex_code.replace("Ω", r"\Omega ")
    latex_code = latex_code.replace("λ", r"\lambda ")
    latex_code = latex_code.replace("℃", r"^{\circ}C")
    latex_code = latex_code.replace("°", r"^{\circ}")
    latex_code = latex_code.replace("⋅", r"\cdot ")
    latex_code = latex_code.replace("・", r"\cdot ")
    latex_code = latex_code.replace("%", r"\% ")
    latex_code = latex_code.replace("δ", r"\delta ")
    latex_code = latex_code.replace("Δ", r"\Delta ")
    latex_code = latex_code.replace("△", r"\Delta ")
    return latex_code


def str_for_data(x):
    # 数据字符化（负数加括号）
    if x >= 0:
        return '%.5g' % x
    else:
        return '(%.5g)' % x


def numlatex(num, prec=5):
    # 科学计数法输出

    # 参数：
    # num：数字
    # prec：有效数字位数（缺省值：5）

    return latex(sympify(num).evalf(prec), min=-3, max=3, mul_symbol="times")


def analyse(data, delta_b1=0, delta_b2=0, symbol='x', unit='', confidence_C=3, confidence_P=0.95):
    # 平均值、标准差、不确定度计算

    # 参数：
    # data：实验数据（一维数组）
    # delta_b1：仪器最大允差Δ_仪
    # delta_b2：估读最大允差Δ_估
    # symbol：该物理量的符号
    # unit：该物理量的单位
    # confidence_C：置信系数C（缺省值：3）
    # confidence_P：置信概率P（缺省值：0.95）

    # 返回：

    # average：平均值
    # sigma：标准差σ
    # delta_b：B类不确定度Δ_B
    # unc：延伸不确定度U

    # averagex：平均值的Latex代码
    # sigmax：标准差σ的Latex代码
    # delta_bx：B类不确定度Δ_B的Latex代码
    # uncx：延伸不确定度U的Latex代码

    # averagex2：平均值的面向MathML的Latex代码
    # sigmax2：标准差σ的面向MathML的Latex代码
    # delta_bx2：B类不确定度Δ_B的面向MathML的Latex代码
    # uncx2：延伸不确定度U的面向MathML的Latex代码

    symbol = adjust_char(symbol)
    unit = adjust_char(unit)

    data.dropna(inplace=True)
    n = len(data)

    average = data.mean()
    averagex = "$$\n"+r"\overline{"+symbol+r"}=\frac{1}{n}\sum_{i=1}^{n}"+symbol+r"_i=\frac{"+'+'.join(str_for_data(x) for x in data)+r"}{"+str(n)+r"}\,\mathrm{"+unit+r"}="+('%.5g' % average)+r"\,\mathrm{"+unit+"}\n$$"
    averagex2 = r"\bar{"+symbol+r"}=\frac{1}{n}\sum\limits_{i=1}^{n}{"+symbol+r"_i}=\frac{"+'+'.join(str_for_data(x) for x in data)+r"}{"+str(n)+r"}\ \mathrm{"+unit+r" }="+('%.5g' % average)+r"\ \mathrm{"+unit+" }"
    # average：平均值

    delta_b = (delta_b1 ** 2 + delta_b2 ** 2) ** 0.5
    if delta_b1 != 0 and delta_b2 != 0:
        delta_bx = "$$\n"+r"\Delta_{B,"+symbol+r"}=\sqrt{\Delta_\text{仪}^2+\Delta_\text{估}^2}=\sqrt{"+('%.5g' % delta_b1)+r"^2+"+('%.5g' % delta_b2)+r"^2}\,\mathrm{"+unit+"}="+('%.5g' % delta_b)+r"\,\mathrm{"+unit+"}\n$$"
        delta_bx2 = r"\Delta_{B,"+symbol+r"}=\sqrt{\Delta_\text{仪}^2+\Delta_\text{估}^2}=\sqrt{"+('%.5g' % delta_b1)+r"^2+"+('%.5g' % delta_b2)+r"^2}\ \mathrm{"+unit+" }="+('%.5g' % delta_b)+r"\ \mathrm{"+unit+" }"
    else:
        delta_bx = "$$\n"+r"\Delta_{B,"+symbol+"}="+('%.5g' % delta_b)+r"\,\mathrm{"+unit+"}\n$$"
        delta_bx2 = r"\Delta_{B,"+symbol+"}="+('%.5g' % delta_b)+r"\ \mathrm{"+unit+" }"
    # delta_b：B类不确定度Δ_B
    
    u_b = delta_b / confidence_C
    # u_b：B类标准不确定度u_B
    unc_b = k_P[confidence_P] * u_b
    # unc_b：k_p*Δ_B/C

    if math.isclose(confidence_C, 3 ** 0.5):
        str_confidence_C = r"\sqrt{3}"
    else:
        str_confidence_C = '%.5g' % confidence_C

    AnalyseData = namedtuple('AnalyseData', ['average', 'averagex', 'averagex2', 'sigma', 'sigmax', 'sigmax2', 'delta_b', 'delta_bx', 'delta_bx2', 'unc', 'uncx', 'uncx2'])

    if n == 1:
        unc = unc_b
        uncx = "$$\nU_{"+symbol+r"}="+r"k_P\frac{\Delta_{B,"+symbol+r"}}{C}="+str(k_P[confidence_P])+r"\times\frac{"+('%.5g' % delta_b)+"}{"+str_confidence_C+r"}\,\mathrm{"+unit+"}="+numlatex(unc)+r"\,\mathrm{"+unit+r"},P="+str(confidence_P)+"\n$$"
        uncx2 = "U_{"+symbol+r"}="+r"k_P\frac{\Delta_{B,"+symbol+r"}}{C}="+str(k_P[confidence_P])+r"\times\frac{"+('%.5g' % delta_b)+"}{"+str_confidence_C+r"}\ \mathrm{"+unit+" }="+numlatex(unc)+r"\ \mathrm{"+unit+r" },\ P="+str(confidence_P)
        # unc：延伸不确定度U
        return AnalyseData(average, averagex, averagex2, 0, "NULL", r"\text{NULL}", delta_b, delta_bx, delta_bx2, unc, uncx, uncx2)

    sigma = math.sqrt(sum([(i - average) ** 2 for i in data])/(n - 1))
    sigmax = "$$\n"+r"\begin{aligned}"+"\n"+r"\sigma_{"+symbol+r"}&=\sqrt{\frac{1}{n-1}\sum_{i=1}^n\left("+symbol+r"_i-\overline{"+symbol+r"}\right)^2}\\"+"\n"+r"&=\sqrt{\frac{"+'+'.join('('+str_for_data(x)+'-'+str_for_data(average)+')^2' for x in data)+"}{"+str(n)+r"-1}}\,\mathrm{"+unit+r"}\\"+"\n"+r"&="+('%.5g' % sigma)+r"\,\mathrm{"+unit+"}\n"+r"\end{aligned}"+"\n$$"
    sigmax2 = r"\sigma_{"+symbol+r"}=\sqrt{\frac{1}{n-1}\sum\limits_{i=1}^n{\left("+symbol+r"_i-\bar{"+symbol+r"}\right)^2}}"+r"=\sqrt{\frac{"+'+'.join('('+str_for_data(x)+'-'+str_for_data(average)+')^2' for x in data)+"}{"+str(n)+r"-1}}\ \mathrm{"+unit+r" }"+r"="+('%.5g' % sigma)+r"\ \mathrm{"+unit+" }"
    # sigma：标准差σ

    u_a = sigma / math.sqrt(n)
    # u_a：A类标准不确定度u_A
    unc_a = t_P[confidence_P][n] * u_a
    # unc_a：A类不确定度t_P*u_A

    unc = math.sqrt(unc_a ** 2 + unc_b ** 2)
    uncx = "$$\n"+r"\begin{aligned}"+"\n"+r"U_{"+symbol+r",P}&=\sqrt{\left(t_P\frac{\sigma_{"+symbol+r"}}{\sqrt{n}}\right)^2+\left(k_P\frac{\Delta_{B,"+symbol+r"}}{C}\right)^2}\\"+"\n"+r"&=\sqrt{\left("+str(t_P[confidence_P][n])+r"\times\frac{"+('%.5g' % sigma)+r"}{\sqrt{"+str(n)+r"}}\right)^2+\left("+str(k_P[confidence_P])+r"\times\frac{"+('%.5g' % delta_b)+"}{"+str_confidence_C+r"}\right)^2}\,\mathrm{"+unit+r"}\\"+"\n&="+numlatex(unc)+r"\,\mathrm{"+unit+r"},P="+str(confidence_P)+"\n"+r"\end{aligned}"+"\n$$"
    uncx2 = "U_{"+symbol+r",P}=\sqrt{\left(t_P\frac{\sigma_{"+symbol+r"}}{\sqrt{n}}\right)^2+\left(k_P\frac{\Delta_{B,"+symbol+r"}}{C}\right)^2}"+r"=\sqrt{\left("+str(t_P[confidence_P][n])+r"\times\frac{"+('%.5g' % sigma)+r"}{\sqrt{"+str(n)+r"}}\right)^2+\left("+str(k_P[confidence_P])+r"\times\frac{"+('%.5g' % delta_b)+"}{"+str_confidence_C+r"}\right)^2}\ \mathrm{"+unit+" }="+numlatex(unc)+r"\ \mathrm{"+unit+r" },\ P="+str(confidence_P)
    # unc：延伸不确定度U

    return AnalyseData(average, averagex, averagex2, sigma, sigmax, sigmax2, delta_b, delta_bx, delta_bx2, unc, uncx, uncx2)


def analyse_lsm(data_X, data_Y, symbol_X='X', symbol_Y='Y', unit_m='', unit_b='', confidence_P=0.95):
    # 最小二乘法线性回归

    # 参数：
    # data_X：X轴数据（一维数组）
    # data_Y：Y轴数据（一维数组）
    # symbol_X：X轴数据物理量的符号
    # symbol_Y：Y轴数据物理量的符号
    # unit_m：斜率的单位
    # unit_b：截距的单位

    # 返回：

    # m：斜率m
    # b：截距b
    # r：线性拟合的相关系数r
    # u_m：斜率的延伸不确定度u_m
    # u_b：截距的延伸不确定度u_b

    # mx：斜率m的Latex代码
    # bx：截距b的Latex代码
    # rx：线性拟合的相关系数r的Latex代码
    # u_mx：斜率的延伸不确定度u_m的Latex代码
    # u_bx：截距的延伸不确定度u_b的Latex代码

    # mx2：斜率m的面向MathML的Latex代码
    # bx2：截距b的面向MathML的Latex代码
    # rx2：线性拟合的相关系数r的面向MathML的Latex代码
    # u_mx2：斜率的延伸不确定度u_m的面向MathML的Latex代码
    # u_bx2：截距的延伸不确定度u_b的面向MathML的Latex代码

    n = len(data_X)
    symbol_X = adjust_char(symbol_X)
    symbol_Y = adjust_char(symbol_Y)
    unit_m = adjust_char(unit_m)
    unit_b = adjust_char(unit_b)

    lsm_res = scipy.stats.linregress(data_X, data_Y)
    # 线性回归并计算统计误差

    mx = "$$\nm="+('%.5g' % lsm_res.slope)+r"\,\mathrm{"+unit_m+"}\n$$"
    mx2 = "m="+('%.5g' % lsm_res.slope)+r"\ \mathrm{"+unit_m+" }"
    # m：斜率m
    bx = "$$\nb="+('%.5g' % lsm_res.intercept)+r"\,\mathrm{"+unit_b+"}\n$$"
    bx2 = "b="+('%.5g' % lsm_res.intercept)+r"\ \mathrm{"+unit_b+" }"
    # b：截距b
    rx = "$$\n"+r"r=\frac{\overline{"+symbol_X+symbol_Y+r"}-\overline{"+symbol_X+r"}\cdot\overline{"+symbol_Y+r"}}{\sqrt{\left(\overline{"+symbol_X+r"^2}-\overline{"+symbol_X+r"}^2\right)\left(\overline{"+symbol_Y+r"^2}-\overline{"+symbol_Y+r"}^2\right)}}="+('%.8g' % lsm_res.rvalue)+"\n$$"
    rx2 = r"r=\frac{\bar{"+symbol_X+symbol_Y+r"}-\bar{"+symbol_X+r"}\cdot\bar{"+symbol_Y+r"}}{\sqrt{\left(\bar{"+symbol_X+r"^2}-\bar{"+symbol_X+r"}^2\right)\left(\bar{"+symbol_Y+r"^2}-\bar{"+symbol_Y+r"}^2\right)}}="+('%.8g' % lsm_res.rvalue)
    # r：线性拟合的相关系数r
    u_m = t_P[confidence_P][n-1] * lsm_res.stderr
    u_mx = "$$\n"+r"u_m=t_P\cdot\lvert m\rvert\cdot\sqrt{\left(\frac{1}{r^2}-1\right)/(n-2)}="+('%.5g' % u_m)+r"\,\mathrm{"+unit_m+"},P="+str(confidence_P)+"\n$$"
    u_mx2 = r"u_m=t_P\cdot\lvert m\rvert\cdot\sqrt{\left(\frac{1}{r^2}-1\right)/(n-2)}="+('%.5g' % u_m)+r"\ \mathrm{"+unit_m+r" },\ P="+str(confidence_P)
    # u_m：斜率的延伸不确定度 u_m
    u_b = t_P[confidence_P][n-1] * lsm_res.intercept_stderr
    u_bx = "$$\n"+r"u_b=u_m\cdot\sqrt{\overline{"+symbol_X+r"^2}}="+('%.5g' % u_b)+r"\,\mathrm{"+unit_b+"},P="+str(confidence_P)+"\n$$"
    u_bx2 = r"u_b=u_m\cdot\sqrt{\bar{"+symbol_X+r"^2}}="+('%.5g' % u_b)+r"\ \mathrm{"+unit_b+" },\ P="+str(confidence_P)
    # u_b：截距的延伸不确定度 u_b

    AnalyseLsmData = namedtuple('AnalyseLsmData', ['m', 'mx', 'mx2', 'b', 'bx', 'bx2', 'r', 'rx', 'rx2', 'u_m', 'u_mx', 'u_mx2', 'u_b', 'u_bx', 'u_bx2'])

    return AnalyseLsmData(lsm_res.slope, mx, mx2, lsm_res.intercept, bx, bx2, lsm_res.rvalue, rx, rx2, u_m, u_mx, u_mx2, u_b, u_bx, u_bx2)


def analyse_com(exp, varr=(), constt=(), unit='', confidence_P=0.95):
    # 表达式及合成不确定度计算

    # 参数：
    # exp：计算表达式（字符串）
    # varr：变量（元组），元组的每个元素均为元组，该子元组的第1个元素为变量名，第2个元素为变量值，第3个元素为其不确定度
    # constt：常量（元组），元组的每个元素均为元组，该子元组的第1个元素为常量名，第2个元素为常量值
    # unit：该物理量的单位
    # confidence_P：置信概率P（缺省值：0.95）
    # 注：若仅计算表达式，不计算合成不确定度，则传入varr=()，并将所有变量信息放在constt元组中

    # 返回：

    # ans：表达式计算值
    # unc：合成不确定度计算值

    # ansx：表达式计算的Latex代码
    # uncx：合成不确定度计算的Latex代码
    # finalx：最终结果的Latex代码

    # ansx2：表达式计算的面向MathML的Latex代码
    # uncx2：合成不确定度计算的面向MathML的Latex代码
    # finalx2：最终结果的面向MathML的Latex代码

    varE = symbols('E')
    pos_for_equal = exp.find('=')
    expr = sympify(exp[(pos_for_equal+1):], locals={'E': varE})
    symbol = latex(sympify(exp[:pos_for_equal], locals={'E': varE}))

    eval_value = {}
    for tuplee in constt:
        eval_value[symbols(tuplee[0])] = tuplee[1]
    for tuplee in varr:
        eval_value[symbols(tuplee[0])] = tuplee[1]

    ans = expr.evalf(subs=eval_value)
    # 表达式计算

    latex_of_expr = latex(expr)
    ansx = "$$\n"+symbol+"="+latex_of_expr+"="
    ansx2 = symbol + "=" + latex_of_expr + "="

    eval_value = {}
    for tuplee in varr:
        eval_value[latex(sympify(tuplee[0]))] = tuplee[1]
    for tuplee in constt:
        eval_value[latex(sympify(tuplee[0]))] = tuplee[1]
    try:
        value_expr = AST2TeX(replaceVariable(TeX2AST(latex_of_expr), eval_value))
        ansx = ansx + value_expr+r"\,\mathrm{"+unit+"}="
        ansx2 = ansx2 + value_expr+r"\ \mathrm{"+unit+" }="
    except:
        pass
    # 表达式数值代入

    ansx = ansx+numlatex(ans)+r"\,\mathrm{"+unit+"}\n$$"
    ansx2 = ansx2+ numlatex(ans) + r"\ \mathrm{"+unit+" }"

    ansx = adjust_char(ansx)
    ansx2 = adjust_char(ansx2)

    if varr == ():
        AnalyseComData = namedtuple('AnalyseComData', ['ans', 'ansx', 'ansx2'])
        return AnalyseComData(ans, ansx, ansx2)
        # 仅计算表达式，不计算合成不确定度

    uncx = "$$\n"+r"\begin{aligned}"+"\nU_{"+symbol+r",P}&="
    uncx2 = "U_{"+symbol+r",P}="
    unc_expr = r"\sqrt{"

    for tuplee in varr:
        unc_expr = unc_expr+r"\left(\frac{\partial "+symbol+r"}{\partial "+tuplee[0]+"}U_{"+tuplee[0]+",P}"+r"\right)^2+"
    unc_expr = unc_expr[:-1]+"}"
    # 合成不确定度公式（偏导形式）

    uncx = uncx+unc_expr+r"\\"+"\n&="
    uncx2 = uncx2 + unc_expr + "="

    for tuplee in varr:
        unc_calc_expr_var = latex(diff(expr, symbols(tuplee[0])))
        unc_expr = unc_expr.replace(r"\frac{\partial "+symbol+r"}{\partial "+tuplee[0]+"}", unc_calc_expr_var)
    # 合成不确定度公式（偏导已求出）

    uncx = uncx+unc_expr+r"\\"+"\n&="
    uncx2 = uncx2 + unc_expr + "="

    eval_uncertainty = {}
    for tuplee in varr:
        eval_uncertainty[latex(sympify(tuplee[0]))] = tuplee[1]
        eval_uncertainty["U_{"+tuplee[0]+",P}"] = tuplee[2]
    for tuplee in constt:
        eval_uncertainty[latex(sympify(tuplee[0]))] = tuplee[1]
    try:
        unc_value_expr = AST2TeX(replaceVariable(TeX2AST(unc_expr), eval_uncertainty))
        uncx = uncx + unc_value_expr+r"\,\mathrm{"+unit+"}" + r"\\"+"\n&="
        uncx2 = uncx2 + unc_value_expr+r"\ \mathrm{"+unit+" }" + "="
    except:
        pass
    # 合成不确定度公式（数值代入）

    unc_calc_expr = sympify("0")
    for tuplee in varr:
        unc_calc_expr = unc_calc_expr + (diff(expr, symbols(tuplee[0])) * symbols("U" + tuplee[0])) ** 2
    unc_calc_expr = sqrt(unc_calc_expr)

    eval_uncertainty = {}
    for tuplee in varr:
        eval_uncertainty[symbols(tuplee[0])] = tuplee[1]
        eval_uncertainty[symbols("U" + tuplee[0])] = tuplee[2]
    for tuplee in constt:
        eval_uncertainty[symbols(tuplee[0])] = tuplee[1]

    unc = unc_calc_expr.evalf(subs=eval_uncertainty)
    # 合成不确定度计算

    uncx = uncx+numlatex(unc)+r"\,\mathrm{"+unit+"},P="+str(confidence_P)+"\n"+r"\end{aligned}"+"\n$$"
    uncx2 = uncx2 + numlatex(unc) + r"\ \mathrm{"+unit+r" },\ P="+str(confidence_P)

    uncx = adjust_char(uncx)
    uncx2 = adjust_char(uncx2)

    unc_first_digit = 0
    for unc_digit in str(unc):
        if "0" < unc_digit <= "9":
            unc_first_digit = unc_digit
            break
    if unc_first_digit == "1" or unc_first_digit == "2": # 判断不确定度的第一个有效数位
        final = "{:.2uL}".format(ufloat(ans, unc))
    else:
        final = "{:.1uL}".format(ufloat(ans, unc))
    if final[:5] != r"\left":
        final = r"\left(" + final + r"\right)"
    finalx = "$$\n" + symbol + "=" + final + r"\,\mathrm{"+unit+"}\n$$"
    finalx2 = symbol + "=" + final + r"\ \mathrm{"+unit+" }"

    finalx = adjust_char(finalx)
    finalx2 = adjust_char(finalx2)

    AnalyseComData = namedtuple('AnalyseComData', ['ans', 'unc', 'ansx', 'uncx', 'finalx', 'ansx2', 'uncx2', 'finalx2'])

    return AnalyseComData(ans, unc, ansx, uncx, finalx, ansx2, uncx2, finalx2)
