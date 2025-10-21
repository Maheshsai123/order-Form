from flask import Flask, render_template, request, jsonify
import os
from datetime import datetime
import json
import pandas as pd

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join('static', 'uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Document counter (simple)
doc_counter = 1

def get_next_doc_no():
    global doc_counter
    doc_no = f'DOC{doc_counter:04d}'
    doc_counter += 1
    return doc_no

@app.route('/')
def index():
    today = datetime.today().strftime('%Y-%m-%d')
    doc_no = get_next_doc_no()
    return render_template('base.html', doc_no=doc_no, today=today)

@app.route('/save', methods=['POST'])
def save():
    doc_no = request.form['doc_no']
    vendor = request.form['vendor']
    date = request.form['date']
    order_json = request.form['order_json']

    try:
        order_data = json.loads(order_json)
        rows = []
        for row in order_data:
            rows.append({
                "S.No": row[0],
                "Item": row[1],
                "Design": row[2],
                "Sizes": row[3],
                "Qty": row[4],
                "Image": row[5],
                "Color": row[6],
                "Details": row[7]
            })

        df = pd.DataFrame(rows)

        filename = f'{doc_no}_{vendor.replace(" ", "_")}.xlsx'
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        df.to_excel(filepath, index=False)

        return jsonify({"message": "Order saved", "file": filename})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
