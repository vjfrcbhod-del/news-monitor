from flask import Flask, jsonify, request
from flask_cors import CORS
from dotenv import load_dotenv
import requests, os

load_dotenv()
app = Flask(__name__)
CORS(app)

CLIENT_ID     = os.getenv("NAVER_CLIENT_ID")
CLIENT_SECRET = os.getenv("NAVER_CLIENT_SECRET")

def fetch_news(query, display=5):
    url = "https://openapi.naver.com/v1/search/news.json"
    headers = {
        "X-Naver-Client-Id": CLIENT_ID,
        "X-Naver-Client-Secret": CLIENT_SECRET,
    }
    params = {"query": query, "display": display, "sort": "date"}
    res = requests.get(url, headers=headers, params=params)
    return res.json().get("items", [])

@app.route("/api/news")
def get_news():
    companies = request.args.getlist("companies")
    result = {}
    for co in companies:
        items = fetch_news(co)
        result[co] = [{
            "title":       item["title"].replace("<b>","").replace("</b>",""),
            "link":        item["link"],
            "source":      item["originallink"],
            "pubDate":     item["pubDate"],
            "description": item["description"].replace("<b>","").replace("</b>",""),
        } for item in items]
    return jsonify(result)

if __name__ == "__main__":
    app.run(port=5000, debug=True)