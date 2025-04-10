from flask import Flask, render_template, request, send_from_directory, jsonify
from flask_socketio import SocketIO, emit
import pandas as pd
from datetime import datetime, timedelta
import os
import cv2
import eventlet

eventlet.monkey_patch()


app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key'
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

socketio = SocketIO(app)

def load_excel_data(file_path):
    df = pd.read_excel(file_path)
    data = []
    for _, row in df.iterrows():
        time_str = str(row["時刻"])
        time_obj = datetime.strptime(time_str, "%H:%M:%S")
        action_duration_str = str(row["動作時間"])
        action_duration_obj = datetime.strptime(action_duration_str, "%H:%M:%S")
        content = str(row["表示内容"]).replace('\n', '<br>')  # 改行をHTMLの改行タグに置換
        rewind_time = str(row["見返し時刻"]) if "見返し時刻" in row else "00:00:00"  # 見返し時刻を取得
        data.append({
            "time": time_obj.strftime("%H:%M:%S"),
            "action": row["動作"],
            "duration": action_duration_obj.strftime("%H:%M:%S"),
            "area": row["表示エリア"],
            "content": content,
            "rewind_time": rewind_time  # 見返し時刻を追加
        })
    
    return data

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/host')
def host():
    return render_template('host.html')

@app.route('/client')
def client():
    return render_template('client.html')

@app.route('/client/<client_id>')
def client_with_id(client_id):
    return render_template('client.html', client_id=client_id)

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'excel' not in request.files or 'video' not in request.files:
        return 'No file part', 400
    
    excel_file = request.files['excel']
    video_file = request.files['video']
    
    if excel_file.filename == '' or video_file.filename == '':
        return 'No selected file', 400
    
    if excel_file and video_file:
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'data.xlsx')
        video_path = os.path.join(app.config['UPLOAD_FOLDER'], 'video.mp4')
        excel_file.save(excel_path)
        video_file.save(video_path)
        
        global excel_data
        excel_data = load_excel_data(excel_path)
        
        return 'Files uploaded successfully', 200

@app.route('/video')
def serve_video():
    return send_from_directory(app.config['UPLOAD_FOLDER'], 'video.mp4')

@app.route('/server_ip')
def server_ip():
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # GoogleのDNSサーバーに接続してローカルIPアドレスを取得
        s.connect(('8.8.8.8', 80))
        local_ip = s.getsockname()[0]
    except Exception:
        local_ip = '127.0.0.1'
    finally:
        s.close()
    return local_ip

@app.route('/capture_photo')
def capture_photo():
    cap = cv2.VideoCapture(0)
    cap.set(cv2.CAP_PROP_AUTO_EXPOSURE, 0.75)  # 自動露出を有効にする
    cap.set(cv2.CAP_PROP_EXPOSURE, -4)  # 露出を調整する
    cap.set(cv2.CAP_PROP_FPS, 5)  # フレームレートを低くしてシャッタースピードを遅くする
    import time
    time.sleep(2)  # カメラが起動するのを待つ
    ret, frame = cap.read()
    if ret:
        photo_path = os.path.join(app.config['UPLOAD_FOLDER'], 'photo.jpg')
        cv2.imwrite(photo_path, frame)
        cap.release()
        return jsonify({"status": "success", "photo_path": photo_path})
    else:
        cap.release()
        return jsonify({"status": "failure"}), 500

@socketio.on('request_data')
def send_data(data):
    current_time = datetime.strptime(data['timestamp'], "%H:%M:%S")
    relevant_data = [entry for entry in excel_data if datetime.strptime(entry['time'], "%H:%M:%S") <= current_time < (datetime.strptime(entry['time'], "%H:%M:%S") + timedelta(hours=datetime.strptime(entry['duration'], "%H:%M:%S").hour, minutes=datetime.strptime(entry['duration'], "%H:%M:%S").minute, seconds=datetime.strptime(entry['duration'], "%H:%M:%S").second))]
    emit('update_display', {'data': relevant_data})

@socketio.on('submit_answer')
def handle_submit_answer(data):
    client_id = data.get('client_id')
    answer = data.get('answer')
    if client_id and answer:
        print(f"クライアント番号: {client_id}, 解答: {answer}")
    else:
        print("Invalid data received")

if __name__ == '__main__':
    import logging
    log = logging.getLogger('werkzeug')
    log.setLevel(logging.ERROR)  # ログレベルをERRORに設定

    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # GoogleのDNSサーバーに接続してローカルIPアドレスを取得
        s.connect(('8.8.8.8', 80))
        local_ip = s.getsockname()[0]
    except Exception:
        local_ip = '127.0.0.1'
    finally:
        s.close()
    print(f"サーバーのIPアドレス: {local_ip}")  # 追加: サーバーのIPアドレスを出力

    host_url = f'http://{local_ip}:5001/host'
    client_url = f'http://{local_ip}:5001/client'
    print(f"ホストとして参加: {host_url}")
    print(f"クライアントとして参加: {client_url}")    
    socketio.run(app, host='0.0.0.0', port=5001, debug=True, use_reloader=False)
