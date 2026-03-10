from flask import Flask, request, jsonify, render_template
from groq import Groq
import pypdf
import zipfile
from xml.etree import ElementTree as ET
import io

app = Flask(__name__)
client = Groq(api_key="REDACTED-GROQ-KEY")

def extract_text(file):
    filename = file.filename.lower()
    if filename.endswith('.pdf'):
        reader = pypdf.PdfReader(io.BytesIO(file.read()))
        return "".join(page.extract_text() for page in reader.pages)
    elif filename.endswith('.docx'):
        docx_bytes = io.BytesIO(file.read())
        with zipfile.ZipFile(docx_bytes) as z:
            xml = z.read('word/document.xml')
        tree = ET.fromstring(xml)
        ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        return " ".join(node.text for node in tree.iter(f'{ns}t') if node.text)
    else:
        return file.read().decode('utf-8')

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/review', methods=['POST'])
def review_contract():
    try:
        if 'contract' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['contract']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        text = extract_text(file)
        if not text or len(text.strip()) < 50:
            return jsonify({'error': 'Could not extract text from file'}), 400

        prompt = request.form.get('prompt', '')

        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {
                    "role": "system",
                    "content": "You are an expert legal contract reviewer."
                },
                {
                    "role": "user",
                    "content": f"{prompt}\n\nCONTRACT:\n{text}"
                }
            ]
        )
        return jsonify({'analysis': response.choices[0].message.content})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
