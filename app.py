from flask import Flask, request, jsonify, render_template, session, send_file
from groq import Groq
import pypdf
import zipfile
from xml.etree import ElementTree as ET
import io
import os
import re
import datetime
from datetime import timedelta

from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY") or "REDACTED-ADMIN-SECRET"
client = Groq(api_key=os.environ.get("GROQ_API_KEY") or "REDACTED-GROQ-KEY")

# ─────────────────────────────────────────
# ACCESS CODES — paste your valid_codes.txt block here
# ─────────────────────────────────────────
VALID_CODES = {
    "JURIVA-PILOT-1",
    "JURIVA-PILOT-2",
    "JURIVA-PILOT-3",
    "JURIVA-PILOT-4",
    "JURIVA-PILOT-5",
}

code_activations = {}
TRIAL_DAYS = 14

# ─────────────────────────────────────────
# TEXT EXTRACTION
# ─────────────────────────────────────────
def extract_text_from_file(file):
    filename = file.filename.lower()
    content = file.read()
    if filename.endswith('.pdf'):
        reader = pypdf.PdfReader(io.BytesIO(content))
        text = "".join(page.extract_text() or '' for page in reader.pages)
    elif filename.endswith('.docx'):
        with zipfile.ZipFile(io.BytesIO(content)) as z:
            xml = z.read('word/document.xml')
        tree = ET.fromstring(xml)
        ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        text = " ".join(node.text for node in tree.iter(f'{ns}t') if node.text)
    else:
        text = content.decode('utf-8')
    del content
    return text

# ─────────────────────────────────────────
# AUTH HELPERS
# ─────────────────────────────────────────
def is_code_valid(code):
    if code not in VALID_CODES:
        return False, "invalid"
    if code in code_activations:
        activated_at = datetime.datetime.fromisoformat(code_activations[code])
        if datetime.datetime.now() > activated_at + timedelta(days=TRIAL_DAYS):
            return False, "expired"
    return True, "ok"

def days_remaining(code):
    if code not in code_activations:
        return TRIAL_DAYS
    activated_at = datetime.datetime.fromisoformat(code_activations[code])
    return max(0, (activated_at + timedelta(days=TRIAL_DAYS) - datetime.datetime.now()).days)

# ─────────────────────────────────────────
# DOCX REPORT GENERATOR
# ─────────────────────────────────────────
def _border_bottom(para, color='C8B89A', size=4):
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), str(size))
    b.set(qn('w:space'), '4'); b.set(qn('w:color'), color)
    pBdr.append(b); pPr.append(pBdr)

def _border_left(para, color='C8B89A', size=8):
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    l = OxmlElement('w:left')
    l.set(qn('w:val'), 'single'); l.set(qn('w:sz'), str(size))
    l.set(qn('w:space'), '8'); l.set(qn('w:color'), color)
    pBdr.append(l); pPr.append(pBdr)

def _runs(para, text, font='Arial', size=10, bold=False, italic=False, color=None):
    """Add text with **bold** markdown support."""
    color = color or RGBColor(0x3A, 0x34, 0x2C)
    parts = re.split(r'\*\*(.*?)\*\*', text)
    for i, part in enumerate(parts):
        r = para.add_run(part)
        r.font.name = font; r.font.size = Pt(size)
        r.font.bold = bold or (i % 2 == 1)
        r.font.italic = italic; r.font.color.rgb = color

def build_docx_report(analysis_text, lang='nl', contract_filename=''):
    doc = DocxDocument()
    for section in doc.sections:
        section.top_margin = Inches(1); section.bottom_margin = Inches(1)
        section.left_margin = Inches(1.2); section.right_margin = Inches(1.2)
    doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.size = Pt(10)

    ACCENT = RGBColor(0xC8, 0xB8, 0x9A)
    DARK   = RGBColor(0x1A, 0x18, 0x15)
    BODY   = RGBColor(0x3A, 0x34, 0x2C)
    MUTED  = RGBColor(0x8A, 0x7E, 0x70)
    DIMMED = RGBColor(0xB0, 0xA4, 0x94)
    QUOTE  = RGBColor(0x7A, 0x6C, 0x5A)

    # Logo
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(2)
    r = p.add_run('Juriva')
    r.font.name = 'Georgia'; r.font.size = Pt(30); r.font.color.rgb = DARK

    # Tagline
    tagline = 'Een tweede paar ogen dat nooit moe wordt.' if lang == 'nl' else 'A second pair of eyes that never gets tired.'
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(16)
    r = p.add_run(tagline)
    r.font.name = 'Georgia'; r.font.size = Pt(10.5); r.font.italic = True; r.font.color.rgb = MUTED

    # Divider
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(14)
    _border_bottom(p, 'C8B89A', 6)

    # Meta row
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(3)
    label = 'CONTRACTANALYSE RAPPORT' if lang == 'nl' else 'CONTRACT ANALYSIS REPORT'
    r = p.add_run(label + '   ')
    r.font.name = 'Arial'; r.font.size = Pt(8); r.font.bold = True
    r.font.all_caps = True; r.font.color.rgb = ACCENT
    r = p.add_run(datetime.datetime.now().strftime('%d %B %Y'))
    r.font.name = 'Arial'; r.font.size = Pt(8); r.font.color.rgb = DIMMED

    # Filename
    if contract_filename:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0); p.paragraph_format.space_after = Pt(20)
        lbl = 'Bestand: ' if lang == 'nl' else 'File: '
        r = p.add_run(lbl); r.font.name = 'Arial'; r.font.size = Pt(8.5); r.font.color.rgb = DIMMED
        r = p.add_run(contract_filename); r.font.name = 'Arial'; r.font.size = Pt(8.5); r.font.color.rgb = MUTED
    else:
        p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(10)

    # ── SECTIONS ──
    raw_sections = re.split(r'\n##\s+', '\n' + analysis_text.strip())
    for section in [s for s in raw_sections if s.strip()]:
        lines = section.strip().split('\n')
        heading_text = lines[0].strip()
        content_lines = lines[1:]

        # Section heading
        h = doc.add_paragraph()
        h.paragraph_format.space_before = Pt(18); h.paragraph_format.space_after = Pt(8)
        _border_bottom(h, 'C8B89A', 4)
        r = h.add_run(heading_text.upper())
        r.font.name = 'Arial'; r.font.size = Pt(8); r.font.bold = True
        r.font.all_caps = True; r.font.color.rgb = ACCENT

        i = 0
        while i < len(content_lines):
            line = content_lines[i].strip()
            if not line:
                i += 1; continue

            # Quote
            if re.match(r'^- (CITAAT|QUOTE):', line):
                text = re.sub(r'^- (CITAAT|QUOTE):\s*"?', '', line).rstrip('"')
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(6)
                _border_left(p, 'C8B89A', 8)
                r = p.add_run('\u201c' + text + '\u201d')
                r.font.name = 'Georgia'; r.font.size = Pt(9.5)
                r.font.italic = True; r.font.color.rgb = QUOTE
                i += 1; continue

            # Sub-labels
            m = re.match(r'^- (ARTIKEL|RISICO|IMPACT|TOELICHTING|ARTICLE|RISK|EXPLANATION|MISSING):\s*(.*)', line)
            if m:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)
                r = p.add_run(m.group(1) + ': ')
                r.font.name = 'Arial'; r.font.size = Pt(9); r.font.bold = True; r.font.color.rgb = DARK
                r = p.add_run(m.group(2))
                r.font.name = 'Arial'; r.font.size = Pt(9.5); r.font.color.rgb = BODY
                i += 1; continue

            # Numbered list
            if re.match(r'^\d+\.', line):
                p = doc.add_paragraph(style='List Number')
                p.paragraph_format.space_before = Pt(3); p.paragraph_format.space_after = Pt(3)
                _runs(p, re.sub(r'^\d+\.\s*', '', line), color=BODY)
                i += 1; continue

            # Bullet
            if line.startswith('- '):
                p = doc.add_paragraph(style='List Bullet')
                p.paragraph_format.space_before = Pt(3); p.paragraph_format.space_after = Pt(3)
                _runs(p, line[2:], color=BODY)
                i += 1; continue

            # Plain
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(3); p.paragraph_format.space_after = Pt(5)
            _runs(p, line, color=BODY)
            i += 1

    # ── FOOTER ──
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20); p.paragraph_format.space_after = Pt(8)
    _border_bottom(p, 'C8B89A', 4)

    disc = ('Dit rapport is uitsluitend bedoeld als ondersteunend instrument en vormt geen juridisch advies. '
            'Juriva B.V. aanvaardt geen aansprakelijkheid voor beslissingen genomen op basis van deze analyse. '
            'Contracten dienen altijd door een bevoegde advocaat te worden beoordeeld vóór ondertekening.'
            if lang == 'nl' else
            'This report is intended solely as a supporting instrument and does not constitute legal advice. '
            'Juriva B.V. accepts no liability for decisions made on the basis of this analysis. '
            'Contracts should always be reviewed by a qualified lawyer before signing.')
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(4)
    r = p.add_run(disc)
    r.font.name = 'Arial'; r.font.size = Pt(7.5); r.font.italic = True; r.font.color.rgb = DIMMED

    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(4)
    r = p.add_run('juriva.nl  \u00b7  info@juriva.nl')
    r.font.name = 'Arial'; r.font.size = Pt(7.5); r.font.color.rgb = ACCENT

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/verify-code', methods=['POST'])
def verify_code():
    data = request.get_json()
    code = data.get('code', '').strip().upper()
    valid, reason = is_code_valid(code)
    if reason == "expired":
        return jsonify({'success': False, 'error': 'expired'})
    if not valid:
        return jsonify({'success': False, 'error': 'invalid'})
    if code not in code_activations:
        code_activations[code] = datetime.datetime.now().isoformat()
    session['authenticated'] = True
    session['code'] = code
    session['terms_accepted'] = session.get('terms_accepted', False)
    return jsonify({'success': True, 'days_remaining': days_remaining(code)})

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
        'terms_accepted': session.get('terms_accepted', False),
        'days_remaining': days_remaining(code)
    })

@app.route('/api/accept-terms', methods=['POST'])
def accept_terms():
    if not session.get('authenticated'):
        return jsonify({'error': 'Not authenticated'}), 401
    session['terms_accepted'] = True
    return jsonify({'success': True})

def purge_session_data():
    for key in ['last_document_text', 'last_playbook_text', 'document_chunks']:
        session.pop(key, None)

@app.route('/api/review', methods=['POST'])
def review_contract():
    if not session.get('authenticated'):
        return jsonify({'error': 'Geen toegang.'}), 401
    if not session.get('terms_accepted'):
        return jsonify({'error': 'terms_not_accepted'}), 403
    code = session.get('code', '')
    valid, reason = is_code_valid(code)
    if not valid:
        session.clear()
        return jsonify({'error': 'expired' if reason == 'expired' else 'Geen toegang.'}), 401
    try:
        if 'contract' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        file = request.files['contract']
        if not file.filename:
            return jsonify({'error': 'No file selected'}), 400

        contract_text = extract_text_from_file(file)
        if not contract_text or len(contract_text.strip()) < 50:
            return jsonify({'error': 'Could not extract text from file'}), 400

        prompt = request.form.get('prompt', '')

        playbook_text = None
        playbook_context = ""
        if 'playbook' in request.files and request.files['playbook'].filename:
            playbook_text = extract_text_from_file(request.files['playbook'])
            playbook_context = f"""
PLAYBOOK INSTRUCTIONS:
The user has uploaded their firm's Gold Standard playbook.
Treat deviations from this playbook as the highest priority risks — above all other general risks.
Flag every clause that contradicts or deviates from the playbook standards.
Quote both the playbook standard AND the contract clause.

PLAYBOOK CONTENT:
{playbook_text[:3000]}

"""

        second_doc_text = None
        cross_ref_context = ""
        if 'second_document' in request.files and request.files['second_document'].filename:
            second_doc_text = extract_text_from_file(request.files['second_document'])
            cross_ref_context = f"""
CROSS-REFERENCE INSTRUCTIONS:
The user has uploaded a second document (e.g. SOW, appendix, or side agreement).
Compare both documents and flag every conflict including:
- Conflicting payment terms or amounts
- Conflicting dates or deadlines
- Conflicting liability caps
- Conflicting defined terms
- Anything promised in one document but missing in the other

SECOND DOCUMENT:
{second_doc_text[:3000]}

Add a section called ## DOCUMENT CONFLICTS at the end of your analysis.
"""

        full_prompt = f"{playbook_context}{cross_ref_context}{prompt}"

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

SCORING CALIBRATION — assign scores honestly using these anchors:
1 — Fully standard. Simple NDA, no asymmetry. Nothing unusual.
2 — Market-conforming. One minor deviation. E.g. standard employment contract, short notice period.
3 — Routine commercial contract. A few clauses worth noting. No serious exposure. E.g. normal SaaS, standard service agreement.
4 — Several non-standard clauses. Some asymmetry. Worth reviewing but manageable. E.g. SaaS with low liability cap only.
5 — Multiple problematic clauses. Real exposure in one area. Negotiation recommended before signing.
6 — Clear imbalance. One or two clauses with material financial or operational risk. E.g. unilateral termination + short notice.
7 — Serious risks across multiple areas. Compounding problems. E.g. liability cap + IP ambiguity + unfair termination.
8 — High danger. Multiple compounding risks that together create significant exposure. Hard to negotiate out of.
9 — Extremely dangerous. Existential clauses: unlimited liability, full IP transfer to counterparty, earn-out traps, punitive penalties.
10 — Predatory. Multiple clauses clearly designed to trap the signing party. Do not sign under any circumstances.

CRITICAL CALIBRATION RULES:
- Most standard commercial contracts score 3-5. Anchor to this range first.
- Only assign 7+ if you can name at least TWO specific clauses each independently justifying high risk.
- Only assign 8+ if those clauses compound each other to create existential exposure.
- A contract with ONE bad clause (even a very bad one) is a 5-6, not an 8.
- Do not inflate scores to seem thorough. Accurate scores build trust.

PRIVACY: You process documents in memory only. Never reference storing or saving documents.
You do not hallucinate. If you are unsure, say so."""
                },
                {
                    "role": "user",
                    "content": f"{full_prompt}\n\nCONTRACT TEXT:\n{contract_text}"
                }
            ],
            temperature=0.1
        )

        analysis = response.choices[0].message.content
        del contract_text
        if playbook_text: del playbook_text
        if second_doc_text: del second_doc_text
        purge_session_data()

        return jsonify({'analysis': analysis})

    except Exception as e:
        purge_session_data()
        return jsonify({'error': str(e)}), 500


@app.route('/api/download-report', methods=['POST'])
def download_report():
    if not session.get('authenticated'):
        return jsonify({'error': 'Geen toegang.'}), 401
    try:
        data = request.get_json()
        analysis_text = data.get('analysis', '')
        lang = data.get('lang', 'nl')
        contract_filename = data.get('filename', '')

        if not analysis_text:
            return jsonify({'error': 'No analysis provided'}), 400

        buf = build_docx_report(analysis_text, lang=lang, contract_filename=contract_filename)
        date_str = datetime.datetime.now().strftime('%Y-%m-%d')
        download_name = f"Juriva_Rapport_{date_str}.docx"

        return send_file(
            buf,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)
