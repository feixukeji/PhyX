from flask import Flask, request, render_template, send_from_directory
import os
import shutil
import time
import threading
import random
import matplotlib
from modulelist import *

app = Flask(__name__)

basepath = os.path.dirname(__file__)  # 当前目录路径

matplotlib.rcParams['xtick.direction'] = 'in'
matplotlib.rcParams['ytick.direction'] = 'in'  # matplotlib图像刻度线朝内


def removedir(dirpath):  # 约5分钟后删除用户数据
    time.sleep(320)
    shutil.rmtree(dirpath)


@app.route('/', methods=['GET'])  # 首页
def index():
    return render_template("index.html")


@app.route('/<string:num>', methods=['GET'])  # 单个实验处理页面
def experiment_index(num):
    if num == "uncertainty":
        return render_template("uncertainty.html")  # 不确定度概观及常用表格页面
    if num in numlist:
        return render_template("experiment.html", num=num, name=eval(num + '.name()'))
    else:
        return render_template("404.html")


@app.route('/pdf-view', methods=['GET'])  # PDF文件在线预览
def view_pdf():
    return render_template("viewer.html")


@app.route('/<string:num>/exam', methods=['GET'])  # 预习测、出门测答案
def exam(num):
    if num in numlist:
        name = eval(num + '.name()')
        try:
            f = open(basepath + '/static/experiment/' + num + '/' + name + '（测试答案）.txt', 'r', encoding='utf8')
            return '<br>'.join(f.readlines())
        except:
            return render_template("404.html")
    else:
        return render_template("404.html")


@app.route('/<string:num>/handle', methods=['POST'])  # 接收上传的Excel文件并处理
def handle(num):
    if num in numlist:
        if 'file' in request.files:
            file = request.files['file']
            filename = file.filename
            if '.' in filename:
                extension = filename.rsplit('.', 1)[1].lower()
                if extension in {'csv', 'xls', 'xlsx'}:
                    fileid = str(random.randrange(1000000))
                    workpath = basepath + "/usrdata/" + num + "/" + fileid + '/'
                    os.makedirs(workpath)
                    file.save(workpath + eval(num + '.name()') + '.' + extension)
                    threading.Thread(target=removedir, args=(workpath,)).start()
                    if eval(num + '.handle(workpath, extension)') == 0:
                        return {
                            "code": 0,
                            "data": fileid
                        }
                    else:
                        return {
                            "code": 1,
                            "data": "文档生成失败，请检查数据文件格式！<br>你可以点击“示例数据”按钮下载示例 Excel 文件，然后将自己的实验数据填入。<br>注意表格底部不要有多余的空行，右侧不要有多余的空列，如有，请删除这些行（列）。"
                        }
            return {
                "code": 1,
                "data": "只支持 csv/xls/xlsx 类型的文件！"
            }
    else:
        return "Not Found"


@app.route('/download/<string:num>/<string:fileid>.docx', methods=['GET'])  # Word文档下载
def download_docx(num, fileid):
    if num in numlist:
        directory = basepath + '/usrdata/' + num + "/" + fileid + '/'
        filename = eval(num + '.name()') + '.docx'
        return send_from_directory(directory, path=filename, as_attachment=True)
    else:
        return render_template("404.html")


@app.route('/baidu_verify_code-Y7syXC9ypc.html', methods=['GET'])
def baidu_verify():
    return render_template("baidu_verify_code-Y7syXC9ypc.html")


@app.route('/BingSiteAuth.xml', methods=['GET'])
def bing_verify():
    return render_template("BingSiteAuth.xml")


@app.errorhandler(404)  # 错误页面
def ERROR_404(e):
    return render_template("404.html")


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
