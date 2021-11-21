from flask import Flask, render_template
import os

app = Flask(__name__)

@app.route('/')
@app.route('/home')
def home():
    return '''
    <h1>이건 h1 제목</h1>
    <p>이건 p 본문 </p>
    <a href="https://flask.palletsprojects.com">Flask 홈페이지 바로가기</a>
    '''
@app.route('/user/<user_name>/<int:user_id>')
def user(user_name, user_id):
    return f'Hello, {user_name}({user_id})!'
    # http://127.0.0.1:5000/user/gy/2017732002

@app.route('/test')
def test():
    return render_template("test.html")

@app.route('/download')
def download():
    psd_folder_path = '/Users/rkdud/Desktop/21_2_record/NIIL/WebToon_Photoshop_Script/Web/CartoonChapter'
    files_list = os.listdir(psd_folder_path)
    print(files_list)
    return render_template("download.html", path_data=files_list)

if __name__ == '__main__':
    app.run(debug=True)
    # debug=True : 해당 파일의 코드를 수정할 때마다 Flask가 변경된 것을 인식하고 다시 시작한다.