import gspread
import json
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime
import os
import re
from flask import Flask, request, abort
from linebot.v3 import WebhookHandler
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    ReplyMessageRequest,
    TextMessage
)
from linebot.v3.webhooks import MessageEvent, TextMessageContent

def append_to_google_sheet(data):
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    client = gspread.authorize(creds)

    sheet = client.open("QC_Defect_Log").sheet1

    sheet.append_row([
        data["Type"],
        data["Line"],
        data["Defect"],
        data["Position"],
        data["Model"],
        data["Total"],
        data["SN"],
        data["Datetime"]
    ])

app = Flask(__name__)

CHANNEL_SECRET = os.environ["CHANNEL_SECRET"]
CHANNEL_ACCESS_TOKEN = os.environ["CHANNEL_ACCESS_TOKEN"]

configuration = Configuration(access_token=CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except Exception as e:
        print("Error:", e)
        abort(400)

    return 'OK'

@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    user_text = event.message.text.strip()

    # ทำงานเฉพาะข้อความที่ขึ้นต้นด้วย #report
    if not user_text.lower().startswith("#report"):
        return

    cleaned_text = user_text.split("#report", 1)[1].strip()

    # ✅ ถ้าพิมพ์แค่ #report อย่างเดียว → ส่งฟอร์มตัวอย่าง
    if cleaned_text == "":
        reply_text = (
            "📋 Please fill in the report using this format:\n\n"
            "#report\n"
            "Cosmetic Fail\n"
            "Line: \n"
            "Defect: \n"
            "Position: \n"
            "Model: \n"
            "Total: \n"
            "SN: "
            )
        
    else:
        try:
            # ไม่ต้องแปลง newline แล้ว
            lines = cleaned_text.split("\n")
            data = {
                "Type": lines[0].strip().title(),
                "Line": "",
                "Defect": "",
                "Position": "",
                "Model": "",
                "Total": 0,
                "SN": "",
                "Datetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            for line in lines[1:]:
                if line.lower().startswith("line:"):
                    data["Line"] = line.split(":",1)[1].strip().upper()
                elif line.lower().startswith("defect:"):
                    data["Defect"] = line.split(":",1)[1].strip()
                elif line.lower().startswith("position:"):
                    data["Position"] = line.split(":",1)[1].strip()
                elif line.lower().startswith("model:"):
                    data["Model"] = line.split(":",1)[1].strip().upper()
                elif line.lower().startswith("total:"):
                    data["Total"] = int(line.split(":",1)[1].strip())
                elif line.lower().startswith("sn:"):
                    data["SN"] = line.split(":",1)[1].strip().upper()
            # ตรวจว่าข้อมูลครบไหม
            if not all([data["Type"], data["Line"], data["Defect"], data["Model"], data["SN"]]):
                raise ValueError("Missing required field")
            append_to_google_sheet(data)
        
            reply_text = "✅ Report saved to Google Sheet successfully."
    
        except Exception as e:
            print("Format Error:", e)
            reply_text = (
                "⚠️ Format incorrect.\n\n"
                "Please Retry:\n\n"
            )

    # ส่งข้อความตอบกลับ
    with ApiClient(configuration) as api_client:
        line_bot_api = MessagingApi(api_client)
        line_bot_api.reply_message(
            ReplyMessageRequest(
                reply_token=event.reply_token,
                messages=[TextMessage(text=reply_text)]
            )
        )

if __name__ == "__main__":

    app.run(host="0.0.0.0", port=10000)





