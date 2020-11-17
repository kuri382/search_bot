from flask import Flask, request, abort
import os
import scrape as sc
import auto_prtsc as ap

from linebot import (
    LineBotApi, WebhookHandler
)
from linebot.exceptions import (
    InvalidSignatureError
)
#LINEでのイベントを取得
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage,
)

app = Flask(__name__)

#環境変数取得
YOUR_CHANNEL_ACCESS_TOKEN = "8acSvOaFnfiZMZYu5R4uIpzCgCMeOG4erVNrA1nI3x44pKaQ9rCNhyAEP0k\
    /s8pgytYEBw/IITBgPiRMTz9fAX5SqTJy/wIkYvuYlGkLDzUv28wyODRUVyifI1faxnSf7VjvbfUUFYR3thS\
        xQlEpLAdB04t89/1O/w1cDnyilFU="
YOUR_CHANNEL_SECRET = "305b40ea346e572f2b25975dec7bdc51"

line_bot_api = LineBotApi(YOUR_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(YOUR_CHANNEL_SECRET)

#アプリケーション本体をopenすると実行される
@app.route("/")
def hello_world():
    return "hello world!"

#/callback　のリンクにアクセスしたときの処理。webhook用。
@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']

    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'
#ここまで定型

#メッセージ受信時のイベント
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    '''
    #オウム返しの場合
    #line_bot_apiのreply_messageメソッドでevent.message.text(ユーザのメッセージ)を返信
    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text=event.message.text))
    '''
    line_bot_api.reply_message(
        event.reply_token,
        TextSendMessage(text=sc.getWeather())
    )

if __name__ == "__main__":
#    app.run()
    port = int(os.getenv("PORT"))
    app.run(host="0.0.0.0", port=port)