from flask import Flask, request, jsonify
import psycopg2
import os

app = Flask(__name__)

DATABASE_URL = os.environ.get("DATABASE_URL")

def get_db_connection():
    conn = psycopg2.connect(DATABASE_URL)
    return conn

@app.route('/tests', methods=['GET'])
def get_tests():
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute('SELECT id, savol, variantlar, tugri_javob FROM tests')
    tests = cur.fetchall()
    cur.close()
    conn.close()
    result = []
    for row in tests:
        result.append({
            "id": row[0],
            "savol": row[1],
            "variantlar": row[2],
            "tugri_javob": row[3]
        })
    return jsonify(result)

@app.route('/tests', methods=['POST'])
def add_tests():
    data = request.get_json()
    if isinstance(data, dict):
        data = [data]
    conn = get_db_connection()
    cur = conn.cursor()
    for test in data:
        cur.execute(
            "INSERT INTO tests (savol, variantlar, tugri_javob) VALUES (%s, %s, %s)",
            (test['savol'], test['variantlar'], test['tugri_javob'])
        )
    conn.commit()
    cur.close()
    conn.close()
    return jsonify({"status": "success"}), 201

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000) 