from flask import Flask, render_template, send_file
import os

app = Flask(__name__)

@app.route('/')
def home():
    return render_template("home.html")
# @app.route('/home')
@app.route('/user/<user_name>/<int:user_id>')
def user(user_name, user_id):
    return f'Hello, {user_name}({user_id})!'
    # http://127.0.0.1:5000/user/gy/2017732002



@app.route('/download')
def download():
    files_list = os.listdir('./Server')
    print(files_list)
    return render_template("download.html", path_data=files_list)

@app.route('/english')
def english():
    files_list = os.listdir('./Server/english')
    return render_template("english_repo.html", path_data=files_list)

@app.route('/download_ENG')
def download_ENG():
    files_list = './Server/english/test/test01.psd'
    return send_file(files_list,
                     mimetype='application/psd',
                     attachment_filename='test01.psd',# 다운받아지는 파일 이름.
                     as_attachment=True)

@app.route('/japan')
def japan():
    files_list = os.listdir('./Server/japan')
    return render_template("japan_repo.html", path_data=files_list)
@app.route('/korea')
def korea():
    files_list = os.listdir('./Server/korea')
    return render_template("korea_repo.html", path_data=files_list)
if __name__ == '__main__':
    app.run(host='0.0.0.0', port='5001', debug=True)
    # debug=True : 해당 파일의 코드를 수정할 때마다 Flask가 변경된 것을 인식하고 다시 시작한다.