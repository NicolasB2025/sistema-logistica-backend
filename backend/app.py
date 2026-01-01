from flask import Flask, request, jsonify
from flask_cors import CORS
import os

app = Flask(__name__)
CORS(app)

@app.route("/")
def home():
    return "Backend online"

@app.route("/ping", methods=["POST"])
def ping():
    data = request.json
    return jsonify({
        "status": "ok",
        "received": data
    })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
