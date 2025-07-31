# backend/run.py
from flask import Flask, jsonify
from flask_cors import CORS
import os

from app.routes.auth import auth_bp
from app.routes.generate import bp as generate_bp
from app.routes.upload import bp as upload_bp

# NEW: import the excel generate blueprint
from app.routes.excel_generate import excel_bp

app = Flask(__name__)

# CORS for /api/* (open for local dev)
CORS(app, resources={r"/api/*": {"origins": "*"}})

@app.after_request
def expose_headers(resp):
    resp.headers["Access-Control-Expose-Headers"] = "Content-Disposition"
    return resp

# your existing uploads folder config
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads", "templates")
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# register your existing blueprints
app.register_blueprint(auth_bp)
app.register_blueprint(generate_bp)
app.register_blueprint(upload_bp)

# register the new excel generator blueprint
app.register_blueprint(excel_bp)

# quick ping
@app.route("/api/ping")
def ping():
    return jsonify(ok=True)

@app.route("/")
def home():
    return "Hello, Creo Certificate Backend!"

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
