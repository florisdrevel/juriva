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
# ACCESS CODES — paste your valid_codes.txt block here
# ─────────────────────────────────────────
VALID_CODES = {
   "JURIVA-PILOT-1",
    "JURIVA-PILOT-2",
    "JURIVA-PILOT-3",
    "JURIVA-PILOT-4",
    "JURIVA-PILOT-5",
   "JURIVA-DENTONS-001",  # amsterdam@dentons.com
    "JURIVA-HOUTHOFF-002",  # info@houthoff.com
    "JURIVA-LOYENSLOEFF-003",  # tom.van.helmond@loyensloeff.com
    "JURIVA-DLAPIPER-004",  # lex.oosterling@dlapiper.com
    "JURIVA-ACTLEGALNETH-005",  # terry.steffens@actlegal-netherlands.com
    "JURIVA-AMICEADVOCAT-006",  # j.vanvliet@amice-advocaten.nl
    "JURIVA-SIXLEGAL-007",  # info@sixlegal.nl
    "JURIVA-ROSS-008",  # info@ross.nl
    "JURIVA-CLIFFORDCHAN-009",  # markjan.arends@cliffordchance.com
    "JURIVA-OSBORNECLARK-010",  # jeroen.bedaux@osborneclarke.com
    "JURIVA-DANIELSHUISM-011",  # bleker@danielshuisman.nl
    "JURIVA-PELSRIJCKEN-012",  # abdessamad.elallaoui@pelsrijcken.nl
    "JURIVA-RECOUP-013",  # info@recoup.nl
    "JURIVA-BRINKHOF-014",  # info@brinkhof.com
    "JURIVA-NAUTADUTILH-015",  # jaco.belder@nautadutilh.com
    "JURIVA-FLORENT-016",  # kees.vandemeent@florent.nl
    "JURIVA-LEEWAY-017",  # marga.verwoert@leeway.nl
    "JURIVA-DELOITTE-018",  # fgrapperhaus@deloitte.nl
    "JURIVA-LVHADVOCATEN-019",  # info@lvh-advocaten.nl
    "JURIVA-LAWANDMORE-020",  # info@lawandmore.nl
    "JURIVA-RUSSELL-021",  # reinier.russell@russell.nl
    "JURIVA-VANDERMEIJAD-022",  # michielhoppenbrouwers@vandermeijadvocaten.nl
    "JURIVA-PRINSENKOSTE-023",  # mr.prins@prinsenkosteradvocaten.nl
    "JURIVA-MULTIWEB-024",  # vdheiden@multiweb.nl
    "JURIVA-ADVOCATENKAN-025",  # neervoort@advocatenkantoorneervoort.nl
    "JURIVA-KNUWERALKMAA-026",  # info@knuweralkmaar.nl
    "JURIVA-KNUWERDENHEL-027",  # info@knuwerdenhelder.nl
    "JURIVA-SPUISTRAAT10-028",  # info@spuistraat10.nl
    "JURIVA-MEIJERSCANAT-029",  # info@meijerscanatan.nl
    "JURIVA-DEBREIJ-030",  # info@debreij.nl
    "JURIVA-LEXENCE-031",  # info@lexence.com
    "JURIVA-STEK-032",  # info@stek.com
    "JURIVA-VANDOORNE-033",  # info@vandoorne.com
    "JURIVA-FLORENT-034",  # info@florent.nl
    "JURIVA-JBLAW-035",  # info@jblaw.nl
    "JURIVA-VRIMANMALAWY-036",  # info@vriman.nl
    "JURIVA-VANCAMPENLIE-037",  # info@vancampenliem.com
    "JURIVA-PLOUM-038",  # info@ploum.nl
    "JURIVA-ACTLEGALNETH-039",  # info@act.nl
    "JURIVA-AKD-040",  # info@akd.nl
    "JURIVA-BANNING-041",  # info@banning.nl
    "JURIVA-BARENTSKRANS-042",  # info@barentskrans.nl
    "JURIVA-BIRDBIRD-043",  # info@bird.nl
    "JURIVA-BOELSZANDERS-044",  # info@boels.nl
    "JURIVA-BONDADVOCATE-045",  # info@bond.nl
    "JURIVA-BRINKHOF-046",  # info@brinkhof.nl
    "JURIVA-BRONSGEESTDE-047",  # info@bronsgeest.nl
    "JURIVA-BUREAUBRANDE-048",  # info@bureau.nl
    "JURIVA-BUREN-049",  # info@buren.nl
    "JURIVA-CMS-050",  # info@cms.nl
    "JURIVA-DAVIDSADVOCA-051",  # info@davids.nl
    "JURIVA-DECLERCQADVO-052",  # info@de.nl
    "JURIVA-DIRKZWAGERLE-053",  # info@dirkzwager.nl
    "JURIVA-DUETADVOCATE-054",  # info@duet.nl
    "JURIVA-DVANADVOCATE-055",  # info@dvan.nl
    "JURIVA-DVDWADVOCATE-056",  # info@dvdw.nl
    "JURIVA-EVERSSOERJAT-057",  # info@evers.nl
    "JURIVA-FINCHDISPUTE-058",  # info@finch.nl
    "JURIVA-FIZADVOCATEN-059",  # info@fiz.nl
    "JURIVA-GREENBERGTRA-060",  # info@greenberg.nl
    "JURIVA-HOLLA-061",  # info@holla.nl
    "JURIVA-HEKKELMAN-062",  # info@hekkelman.nl
    "JURIVA-HVGLAW-063",  # info@hvg.nl
    "JURIVA-JAHAERAYMAKE-064",  # info@jahaeraymakers.nl
    "JURIVA-JEBBINKSOETE-065",  # info@jebbink.nl
    "JURIVA-KENNEDYVANDE-066",  # info@kennedy.nl
    "JURIVA-KIENHUISHOVI-067",  # info@kienhuishoving.nl
    "JURIVA-KIENHUISLEGA-068",  # info@kienhuis.nl
    "JURIVA-KPMGLAWNETHE-069",  # info@kpmg.nl
    "JURIVA-LAGRO-070",  # info@la.nl
    "JURIVA-LXAATTORNEYS-071",  # info@lxa.nl
    "JURIVA-MAVERICKADVO-072",  # info@maverick.nl
    "JURIVA-NEWGROUNDLAW-073",  # info@newground.nl
    "JURIVA-PELSRIJCKEN-074",  # info@pels.nl
    "JURIVA-QUINZNETHERL-075",  # info@quinz.nl
    "JURIVA-SCHAAPADVOCA-076",  # info@schaap.nl
    "JURIVA-SEEGERSLEBKO-077",  # info@seegers.nl
    "JURIVA-SIMMONSSIMMO-078",  # info@simmons.nl
    "JURIVA-SQUIREPATTON-079",  # info@squire.nl
    "JURIVA-STIBBENETHER-080",  # info@stibbe.nl
    "JURIVA-TENHOLTERNOO-081",  # info@ten.nl
    "JURIVA-TRIPADVOCATE-082",  # info@trip.nl
    "JURIVA-VANBENTHEMKE-083",  # info@van.nl
    "JURIVA-VESTIUSADVOC-084",  # info@vestius.nl
    "JURIVA-VISSERSCHAAP-085",  # info@visser.nl
    "JURIVA-WIJNSTAEL-086",  # info@wijn.nl
    "JURIVA-WINDTLEGRAND-087",  # info@windt.nl
    "JURIVA-WINTERTALING-088",  # info@wintertaling.nl
    "JURIVA-WLADIMIROFFA-089",  # info@wladimiroff.nl
    "JURIVA-YOURCORPORAT-090",  # info@your.nl
    "JURIVA-ALLENOVERYNE-091",  # info@allen.nl
    "JURIVA-BAKERMCKENZI-092",  # info@baker.nl
    "JURIVA-CLIFFORDCHAN-093",  # info@clifford.nl
    "JURIVA-DLAPIPERNETH-094",  # info@dla.nl
    "JURIVA-DENTONSNETHE-095",  # info@dentons.nl
    "JURIVA-JONESDAYAMST-096",  # info@jones.nl
    "JURIVA-LINKLATERSAM-097",  # info@linklaters.nl
    "JURIVA-LOYENSLOEFF-098",  # info@loyens.nl
    "JURIVA-NAUTADUTILH-099",  # info@nautadutilh.nl
    "JURIVA-HABRAKENRUTT-100",  # info@habrakenrutten.nl
    "JURIVA-LAWTONLAWYER-101",  # info@lawton.nl
    "JURIVA-TRIPELSADVOC-102",  # info@tripels.nl
    "JURIVA-SEVERIJNHULS-103",  # info@severijn.nl
    "JURIVA-LAWMOREEINDH-104",  # info@law.nl
    "JURIVA-LEXENCELITIG-105",  # info@lexence.nl
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
    # FEATURE 5: immediately delete content from memory after extraction
    del content
    return text

# ─────────────────────────────────────────
# AUTH HELPERS
# ─────────────────────────────────────────
def is_code_valid(code):
    if code not in VALID_CODES:
        return False, "invalid"
    if code in code_activations:
        activated_at = datetime.fromisoformat(code_activations[code])
        if datetime.now() > activated_at + timedelta(days=TRIAL_DAYS):
            return False, "expired"
    return True, "ok"

def days_remaining(code):
    if code not in code_activations:
        return TRIAL_DAYS
    activated_at = datetime.fromisoformat(code_activations[code])
    expires_at = activated_at + timedelta(days=TRIAL_DAYS)
    return max(0, (expires_at - datetime.now()).days)

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
        code_activations[code] = datetime.now().isoformat()
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

# ─────────────────────────────────────────
# FEATURE 5: PURGE SESSION
# ─────────────────────────────────────────
def purge_session_data():
    """Immediately wipe all document data from session after analysis."""
    keys_to_clear = ['last_document_text', 'last_playbook_text', 'document_chunks']
    for key in keys_to_clear:
        session.pop(key, None)

# ─────────────────────────────────────────
# CORE REVIEW ENDPOINT
# ─────────────────────────────────────────
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

        # Extract and immediately process — never persist to disk
        contract_text = extract_text_from_file(file)
        if not contract_text or len(contract_text.strip()) < 50:
            return jsonify({'error': 'Could not extract text from file'}), 400

        prompt = request.form.get('prompt', '')

        # ─────────────────────────────────────────
        # FEATURE 1: PLAYBOOK INTEGRATION
        # ─────────────────────────────────────────
        playbook_text = None
        playbook_context = ""
        if 'playbook' in request.files and request.files['playbook'].filename:
            playbook_file = request.files['playbook']
            playbook_text = extract_text_from_file(playbook_file)
            playbook_context = f"""
PLAYBOOK INSTRUCTIONS:
The user has uploaded their firm's Gold Standard playbook. 
You MUST treat deviations from this playbook as the highest priority risks — 
above all other general risks. Flag every clause that contradicts or deviates 
from the playbook standards. Quote both the playbook standard AND the contract clause.

PLAYBOOK CONTENT:
{playbook_text[:3000]}

"""

        # ─────────────────────────────────────────
        # FEATURE 3: MULTI-DOCUMENT CROSS-REFERENCING
        # ─────────────────────────────────────────
        second_doc_text = None
        cross_ref_context = ""
        if 'second_document' in request.files and request.files['second_document'].filename:
            second_file = request.files['second_document']
            second_doc_text = extract_text_from_file(second_file)
            cross_ref_context = f"""
CROSS-REFERENCE INSTRUCTIONS:
The user has uploaded a second document (e.g. SOW, appendix, or side agreement).
You MUST compare both documents and flag every conflict including:
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

        # FEATURE 5: process in memory only
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
PRIVACY: You process documents in memory only. Never reference storing or saving documents.
You do not hallucinate. If you are unsure, you say so."""
                },
                {
                    "role": "user",
                    "content": f"{full_prompt}\n\nCONTRACT TEXT:\n{contract_text}"
                }
            ],
            temperature=0.1
        )

        analysis = response.choices[0].message.content

        # FEATURE 5: purge all document data immediately after analysis
        del contract_text
        if playbook_text:
            del playbook_text
        if second_doc_text:
            del second_doc_text
        purge_session_data()

        return jsonify({'analysis': analysis})

    except Exception as e:
        purge_session_data()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)

