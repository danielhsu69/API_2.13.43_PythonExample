import Order
from flask import Flask, request

# 建立 Flask 應用程式物件
app = Flask(__name__)

# 設定路由
@app.route('/webhook', methods=['POST'])
def handle_webhook():
    # 取得 Webhook 傳送過來的資料
    data = request.json
    print(data)
    
    # 回傳回應
    return 'Webhook received successfully!'

# 啟動應用程式
if __name__ == '__main__':
    app.run()
