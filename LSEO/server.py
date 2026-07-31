from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import io
import datetime
import os

from pdf_generator import generate_pdf
from docx_generator import generate_docx

app = Flask(__name__)
# Atļaujam Shopify lapai sūtīt pieprasījumus uz šo serveri
CORS(app)

@app.route('/generate/<file_type>', methods=['POST'])
def generate_doc(file_type):
    data = request.json

    # Ģenerējam faila nosaukumu
    doc_id = data.get('doc_id', 'BR_0000').replace(" ", "_")
    doc_type_name = data.get('doc_type', 'Pavadzime').replace(" ", "_")

    if file_type == 'pdf':
        buffer = generate_pdf(data)
        filename = f"{doc_type_name}_{doc_id}.pdf"
        mime = "application/pdf"
    elif file_type == 'docx':
        buffer = generate_docx(data)
        filename = f"{doc_type_name}_{doc_id}.docx"
        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    else:
        return jsonify({"error": "Nezināms formāts"}), 400

    # Nosūtām atpakaļ lietotājam lejupielādei
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype=mime
    )

if __name__ == '__main__':
    # Serveris klausās uz portu
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
