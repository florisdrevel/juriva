from flask import Flask, request, jsonify, render_template, session
from groq import Groq
import pypdf
import zipfile
from xml.etree import ElementTree as ET
import io
import os

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or "REDACTED-ADMIN-SECRET"

client = Groq(api_key=os.environ.get("GROQ_API_KEY") or "your-groq-key-here")

# ─────────────────────────────────────────
# ACCESS CODES — add or remove codes here
# ─────────────────────────────────────────
VALID_CODES = {
    "JURIVA-PILOT-1",
    "JURIVA-PILOT-2",
    "JURIVA-PILOT-3",
    "JURIVA-PILOT-4",
    "JURIVA-PILOT-5",
}

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

@app.route('/api/verify-code', methods=['POST'])
def verify_code():
    data = request.get_json()
    code = data.get('code', '').strip().upper()
    if code in VALID_CODES:
        session['authenticated'] = True
        session['code'] = code
        return jsonify({'success': True})
    return jsonify({'success': False, 'error': 'Ongeldige toegangscode'})

@app.route('/api/check-auth', methods=['GET'])
def check_auth():
    return jsonify({'authenticated': session.get('authenticated', False)})

@app.route('/api/review', methods=['POST'])
def review_contract():
    if not session.get('authenticated'):
        return jsonify({'error': 'Geen toegang. Voer uw toegangscode in.'}), 401
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
                    "content": """You are a senior legal contract reviewer with 20 years of experience.
Your analysis must be:
1. EVIDENCE-BASED: Always quote exact contract text before analysing it
2. PRECISE: Never paraphrase clauses — quote them verbatim
3. SPECIFIC: Reference exact article numbers
4. HONEST: If something is missing from the contract, say so
5. PRACTICAL: Give actionable advice, not vague warnings
You do not hallucinate. If you are unsure, you say so."""
                },
                {
                    "role": "user",
                    "content": f"{prompt}\n\nCONTRACT TEXT:\n{text}"
                }
            ],
            temperature=0.1
        )
        return jsonify({'analysis': response.choices[0].message.content})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
