from flask import Flask,render_template,request
import util,os

app=Flask(__name__,template_folder='static')

@app.route('/')
def start():
    return render_template('html/index.html')

@app.route('/parse_excel',methods=['post'])
def parse_excel():
    file=request.files['excel']
    file_name='test'+os.path.splitext(file.filename)[1]
    file.save(os.path.join(os.path.dirname(__file__),'res//'+file_name))
    #获得excel解析数据
    parse_result=util.read_original_excel(file_name)
    return render_template('html/index.html',parse_result=parse_result)



if __name__=='__main__':
    app.run(port=80,debug=True)
