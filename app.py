from flask import Flask, request, jsonify, render_template, session
from groq import Groq
import pypdf
import zipfile
from xml.etree import ElementTree as ET
import io
import os
from datetime import datetime, timedelta

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or "REDACTED-ADMIN-SECRET"

client = Groq(api_key=os.environ.get("GROQ_API_KEY") or "your-groq-key-here")

# ─────────────────────────────────────────
# ACCESS CODES
# Each code can only be activated once.
# After activation it expires in 14 days.
# Add new codes here as needed.
# ─────────────────────────────────────────
VALID_CODES = {
    "JURIVA-PILOT-1",
    "JURIVA-PILOT-2",
    "JURIVA-PILOT-3",
    "JURIVA-PILOT-4",
    "JURIVA-PILOT-5",
}

# Stores activation times: { "CODE": "2024-03-11T10:00:00" }
# This resets when the server restarts on Railway.
# Good enough for now — we'll add a database later.
code_activations = {}

TRIAL_DAYS = 14

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

def is_code_valid(code):
    if code not in VALID_CODES:
        return False, "invalid"
    if code in code_activations:
        activated_at = datetime.fromisoformat(code_activations[code])
        expires_at = activated_at + timedelta(days=TRIAL_DAYS)
        if datetime.now() > expires_at:
            return False, "expired"
    return True, "ok"

def days_remaining(code):
    if code not in code_activations:
        return TRIAL_DAYS
    activated_at = datetime.fromisoformat(code_activations[code])
    expires_at = activated_at + timedelta(days=TRIAL_DAYS)
    remaining = (expires_at - datetime.now()).days
    return max(0, remaining)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/verify-code', methods=['POST'])
def verify_code():
    data = request.get_json()
    code = data.get('code', '').strip().upper()

    valid, reason = is_code_valid(code)

    if reason == "expired":
        return jsonify({
            'success': False,
            'error': 'expired'
        })

    if not valid:
        return jsonify({
            'success': False,
            'error': 'invalid'
        })

    # Activate code on first use
    if code not in code_activations:
        code_activations[code] = datetime.now().isoformat()

    remaining = days_remaining(code)
    session['authenticated'] = True
    session['code'] = code

    return jsonify({
        'success': True,
        'days_remaining': remaining
    })

@app.route('/api/check-auth', methods=['GET'])
def check_auth():
    if not session.get('authenticated'):
        return jsonify({'authenticated': False})

    code = session.get('code', '')
    valid, reason = is_code_valid(code)

    if not valid:
        session.clear()
        return jsonify({'authenticated': False, 'reason': reason})

    return jsonify({
        'authenticated': True,
        'days_remaining': days_remaining(code)
    })

@app.route('/api/review', methods=['POST'])
def review_contract():
    if not session.get('authenticated'):
        return jsonify({'error': 'Geen toegang. Voer uw toegangscode in.'}), 401

    code = session.get('code', '')
    valid, reason = is_code_valid(code)
    if not valid:
        session.clear()
        return jsonify({'error': 'expired' if reason == 'expired' else 'Geen toegang.'}), 401

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
