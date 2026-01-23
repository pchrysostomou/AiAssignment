import os
import json
import time
import re
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, Response, stream_with_context, jsonify, session, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from dotenv import load_dotenv
import anthropic
import openai
from google import generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import stripe
import PyPDF2
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.units import inch

# Load environment variables
load_dotenv()

# Initialize Flask app
app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-secret-key-change-me')
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URI', 'sqlite:///ai_architect.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['REPORTS_FOLDER'] = 'reports'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['REPORTS_FOLDER'], exist_ok=True)

# Initialize extensions
db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

# Stripe configuration
stripe_key = os.getenv('STRIPE_SECRET_KEY', '')
if stripe_key:
    stripe.api_key = stripe_key
    masked_key = stripe_key[:7] + '...' + stripe_key[-4:] if len(stripe_key) > 11 else '***'
    print(f"✅ Stripe Key Loaded: {masked_key}")
else:
    print("⚠️ WARNING: STRIPE_SECRET_KEY not found in .env")

DOMAIN = os.getenv('DOMAIN', 'http://localhost:5000')

# Configuration
MAX_ROUNDS = 10
SUCCESS_THRESHOLD = 90
STAGNATION_THRESHOLD = 1.5
HISTORY_PRUNING_START = 4
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'txt'}

# Pricing tiers
PRICING = {
    'essay_std': {'price': 2000, 'display': '£20', 'words': 2000},
    'essay_ext': {'price': 4000, 'display': '£40', 'words': 4000},
    'plagiarism': {'price': 1000, 'display': '£10'},
    'ai_check': {'price': 1000, 'display': '£10'},
    'grader': {'price': 1500, 'display': '£15'}
}

# Initialize AI clients
anthropic_client = anthropic.Anthropic(api_key=os.getenv('ANTHROPIC_API_KEY'))
openai.api_key = os.getenv('OPENAI_API_KEY')
genai.configure(api_key=os.getenv('GOOGLE_API_KEY'))

# Database Models
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    essays = db.relationship('Essay', backref='author', lazy=True)
    reports = db.relationship('Report', backref='author', lazy=True)

    def set_password(self, password):
        self.password_hash = generate_password_hash(password)

    def check_password(self, password):
        return check_password_hash(self.password_hash, password)

class Essay(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    title = db.Column(db.String(200), nullable=False)
    instructions = db.Column(db.Text, nullable=False)
    final_content = db.Column(db.Text, nullable=False)
    final_score = db.Column(db.Integer, nullable=False)
    rounds_used = db.Column(db.Integer, nullable=False)
    stop_reason = db.Column(db.String(100), nullable=False)
    word_count_limit = db.Column(db.Integer, nullable=False, default=2000)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Report(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    tool_type = db.Column(db.String(50), nullable=False)  # plagiarism, ai_check, grader
    title = db.Column(db.String(200), nullable=False)
    result_data = db.Column(db.Text, nullable=False)
    file_path = db.Column(db.String(500))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# File handling
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text_from_pdf(file_path):
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
            return text.strip()
    except Exception as e:
        raise Exception(f"Error reading PDF: {str(e)}")

def extract_text_from_docx(file_path):
    try:
        doc = Document(file_path)
        text = ""
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
        return text.strip()
    except Exception as e:
        raise Exception(f"Error reading DOCX: {str(e)}")

def extract_text_from_txt(file_path):
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read().strip()
    except Exception as e:
        raise Exception(f"Error reading TXT: {str(e)}")

def parse_uploaded_file(file_path, filename):
    ext = filename.rsplit('.', 1)[1].lower()
    if ext == 'pdf':
        return extract_text_from_pdf(file_path)
    elif ext == 'docx':
        return extract_text_from_docx(file_path)
    elif ext == 'txt':
        return extract_text_from_txt(file_path)
    else:
        raise Exception("Unsupported file type")

# AI Helper Functions
def call_with_retry(func, max_retries=3, delay=2):
    for attempt in range(max_retries):
        try:
            return func()
        except Exception as e:
            if attempt == max_retries - 1:
                raise
            time.sleep(delay * (attempt + 1))
    return None

def call_claude_writer(prompt, max_tokens=4000):
    def api_call():
        response = anthropic_client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text
    return call_with_retry(api_call)

def call_gpt4_examiner(prompt):
    def api_call():
        response = openai.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        return response.choices[0].message.content
    return call_with_retry(api_call)

def call_gemini_reviewer(prompt):
    def api_call():
        model = genai.GenerativeModel('gemini-1.5-pro')
        response = model.generate_content(prompt)
        return response.text
    return call_with_retry(api_call)

def extract_score(examiner_response):
    try:
        matches = re.findall(r'\b(\d{1,3})\b', examiner_response)
        for match in matches:
            score = int(match)
            if 0 <= score <= 100:
                return score
        return 0
    except:
        return 0

# TOOL A: The Detective (Plagiarism Checker)
def check_plagiarism_gemini(text):
    """Use Gemini to check for plagiarism with web search"""
    prompt = f"""You are a plagiarism detection expert. Analyze the following text for potential plagiarism by searching your knowledge base and web sources.

TEXT TO ANALYZE:
{text}

For each suspicious passage:
1. Identify the exact text that appears plagiarized
2. Find the original source URL if available
3. Calculate similarity percentage

Return your findings in this JSON format:
{{
    "overall_plagiarism_score": 0-100,
    "matches": [
        {{
            "text": "suspicious passage",
            "source": "URL or source name",
            "similarity": 0-100
        }}
    ],
    "verdict": "ORIGINAL/SUSPICIOUS/PLAGIARIZED"
}}"""

    try:
        model = genai.GenerativeModel('gemini-1.5-pro')
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return json.dumps({"error": str(e), "overall_plagiarism_score": 0, "matches": [], "verdict": "ERROR"})

def generate_plagiarism_pdf(text, analysis_result, user_id):
    """Generate PDF report for plagiarism check"""
    filename = f"plagiarism_report_{user_id}_{int(time.time())}.pdf"
    filepath = os.path.join(app.config['REPORTS_FOLDER'], filename)
    
    doc = SimpleDocTemplate(filepath, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    
    # Title
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=24, textColor=colors.HexColor('#b537f2'))
    story.append(Paragraph("🔍 The Detective - Plagiarism Report", title_style))
    story.append(Spacer(1, 0.3*inch))
    
    # Metadata
    story.append(Paragraph(f"<b>Generated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles['Normal']))
    story.append(Spacer(1, 0.2*inch))
    
    # Analysis Result
    try:
        result = json.loads(analysis_result)
        score = result.get('overall_plagiarism_score', 0)
        verdict = result.get('verdict', 'UNKNOWN')
        
        verdict_style = ParagraphStyle('Verdict', parent=styles['Heading2'], 
                                      textColor=colors.red if score > 30 else colors.green)
        story.append(Paragraph(f"Verdict: {verdict} ({score}% match)", verdict_style))
        story.append(Spacer(1, 0.2*inch))
        
        # Matches
        matches = result.get('matches', [])
        if matches:
            story.append(Paragraph("<b>Suspicious Passages:</b>", styles['Heading3']))
            story.append(Spacer(1, 0.1*inch))
            
            for i, match in enumerate(matches, 1):
                match_text = match.get('text', 'N/A')
                source = match.get('source', 'Unknown')
                similarity = match.get('similarity', 0)
                
                story.append(Paragraph(f"<b>Match {i} ({similarity}% similar):</b>", styles['Normal']))
                story.append(Paragraph(f'"{match_text}"', styles['Italic']))
                story.append(Paragraph(f"<b>Source:</b> {source}", styles['Normal']))
                story.append(Spacer(1, 0.15*inch))
    except:
        story.append(Paragraph("Analysis result format error", styles['Normal']))
    
    # Original Text
    story.append(PageBreak())
    story.append(Paragraph("<b>Original Text Analyzed:</b>", styles['Heading3']))
    story.append(Spacer(1, 0.1*inch))
    story.append(Paragraph(text[:2000] + "..." if len(text) > 2000 else text, styles['Normal']))
    
    doc.build(story)
    return filename

# TOOL B: The Oracle (AI Detection)
def check_ai_content(text):
    """Use GPT-4 to detect AI-generated content"""
    prompt = f"""You are an AI content detection expert. Analyze the following text to determine if it was written by AI or a human.

TEXT TO ANALYZE:
{text}

Look for AI indicators:
- Repetitive phrasing patterns
- Overly formal or perfect grammar
- Lack of personal voice
- Generic transitions
- AI buzzwords (delve, tapestry, multifaceted, landscape, etc.)

Return JSON format:
{{
    "ai_probability": 0-100,
    "indicators": ["list of specific AI indicators found"],
    "verdict": "HUMAN/LIKELY_HUMAN/UNCERTAIN/LIKELY_AI/AI",
    "explanation": "brief explanation"
}}"""

    try:
        response = openai.chat.completions.create(
            model="gpt-4o",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        return response.choices[0].message.content
    except Exception as e:
        return json.dumps({"error": str(e), "ai_probability": 0, "verdict": "ERROR"})

# TOOL C: The Grader (Strict Marking)
def grade_assignment(brief_text, essay_text):
    """Grade assignment using all 3 professors"""
    
    # Prof. Quill - Content Analysis
    quill_prompt = f"""You are Prof. Quill, an expert academic evaluator. Analyze this student essay against the assignment brief.

ASSIGNMENT BRIEF:
{brief_text}

STUDENT ESSAY:
{essay_text}

Evaluate:
1. How well does it address the brief?
2. Content quality and depth
3. Structure and organization

Provide a score (0-100) and detailed feedback."""

    # Dr. Strict - Technical Grading
    strict_prompt = f"""You are Dr. Strict, a harsh academic grader. Grade this essay strictly.

ASSIGNMENT BRIEF:
{brief_text}

STUDENT ESSAY:
{essay_text}

Evaluate:
1. Grammar and writing quality
2. Citation and referencing
3. Academic rigor

Provide a score (0-100) and identify specific weaknesses."""

    # Dean Logic - Overall Assessment
    logic_prompt = f"""You are Dean Logic, the final academic authority. Provide an overall assessment.

ASSIGNMENT BRIEF:
{brief_text}

STUDENT ESSAY:
{essay_text}

Provide:
1. Overall score (0-100)
2. Key strengths
3. Critical improvements needed"""

    try:
        quill_response = call_claude_writer(quill_prompt)
        quill_score = extract_score(quill_response)
        
        strict_response = call_gpt4_examiner(strict_prompt)
        strict_score = extract_score(strict_response)
        
        logic_response = call_gemini_reviewer(logic_prompt)
        logic_score = extract_score(logic_response)
        
        average_score = round((quill_score + strict_score + logic_score) / 3)
        
        result = {
            "average_score": average_score,
            "quill_score": quill_score,
            "quill_feedback": quill_response,
            "strict_score": strict_score,
            "strict_feedback": strict_response,
            "logic_score": logic_score,
            "logic_feedback": logic_response
        }
        
        return json.dumps(result)
    except Exception as e:
        return json.dumps({"error": str(e), "average_score": 0})

# TOOL D: The Architect (Essay Writer with Citation Verification)
def generate_essay_with_citations_stream(instructions, word_count_limit):
    """Essay generation with citation verification"""
    yield f"data: {json.dumps({'type': 'log', 'message': f'🎓 The Academic Board convening (Target: {word_count_limit} words)...', 'professor': 'system'})}\n\n"
    
    # Initial Draft
    yield f"data: {json.dumps({'type': 'log', 'message': '✍️ Prof. Quill drafting (Human-like style)...', 'professor': 'quill'})}\n\n"
    
    writer_prompt = f"""You are Prof. Quill, a human ghostwriter. Write naturally like a real person, NOT an AI.

CRITICAL STYLE RULES:
- NEVER use AI buzzwords: delve, tapestry, multifaceted, landscape, realm, etc.
- Vary sentence length (mix short and long)
- Use contractions occasionally
- Include minor stylistic imperfections
- Write with a personal, authentic voice

CONTENT REQUIREMENTS:
- Target: {word_count_limit} words
- Include REAL academic citations (Harvard/APA format)
- Every major claim needs a citation
- Use actual published papers (author, year, title)

INSTRUCTIONS:
{instructions}

Write a comprehensive, well-researched essay with proper citations."""

    try:
        current_draft = call_claude_writer(writer_prompt)
        yield f"data: {json.dumps({'type': 'log', 'message': '✅ Initial draft complete', 'professor': 'quill'})}\n\n"
    except Exception as e:
        yield f"data: {json.dumps({'type': 'error', 'message': f'❌ Error: {str(e)}', 'professor': 'quill'})}\n\n"
        return

    scores = []
    best_draft = current_draft
    best_score = 0
    stop_reason = ""

    yield f"data: {json.dumps({'type': 'log', 'message': f'🔄 Starting refinement (Max {MAX_ROUNDS} rounds)...', 'professor': 'system'})}\n\n"
    
    for round_num in range(1, MAX_ROUNDS + 1):
        yield f"data: {json.dumps({'type': 'log', 'message': f'━━━ Round {round_num}/{MAX_ROUNDS} ━━━', 'professor': 'system'})}\n\n"
        
        # Dr. Strict grades
        yield f"data: {json.dumps({'type': 'log', 'message': '🎯 Dr. Strict grading...', 'professor': 'strict'})}\n\n"
        
        examiner_prompt = f"""Grade this essay (0-100). Check quality, structure, and citations.

INSTRUCTIONS: {instructions}
ESSAY: {current_draft}

Score format: "Score: [number]" """

        try:
            examiner_response = call_gpt4_examiner(examiner_prompt)
            score = extract_score(examiner_response)
            scores.append(score)
            
            yield f"data: {json.dumps({'type': 'score', 'round': round_num, 'score': score, 'professor': 'strict'})}\n\n"
            
            verdict_msg = f"📊 Dr. Strict's verdict: {score}/100"
            yield f"data: {json.dumps({'type': 'log', 'message': verdict_msg, 'professor': 'strict'})}\n\n"
            
            if score > best_score:
                best_score = score
                best_draft = current_draft
                
        except Exception as e:
            yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Error: {str(e)}', 'professor': 'strict'})}\n\n"
            score = 0
            scores.append(0)

        if score >= SUCCESS_THRESHOLD:
            stop_reason = f"SUCCESS: Score {score} >= {SUCCESS_THRESHOLD}"
            yield f"data: {json.dumps({'type': 'log', 'message': f'🎉 {stop_reason}!', 'professor': 'system'})}\n\n"
            break

        # CITATION VERIFICATION by Dean Logic (Gemini)
        yield f"data: {json.dumps({'type': 'log', 'message': '🔍 Dean Logic verifying citations...', 'professor': 'logic'})}\n\n"
        
        citation_check_prompt = f"""You are Dean Logic, citation verification expert. Analyze ALL citations in this essay.

ESSAY:
{current_draft}

For EACH citation found:
1. Extract the full citation
2. Search your knowledge base to verify if the paper/source EXISTS
3. Check if it's relevant to the claim
4. Mark as: VERIFIED / SUSPICIOUS / FAKE / DEAD_LINK

Return JSON:
{{
    "citations_found": [
        {{
            "citation": "full citation text",
            "status": "VERIFIED/SUSPICIOUS/FAKE",
            "reason": "why"
        }}
    ],
    "fake_count": 0,
    "feedback": "overall citation quality feedback"
}}

CRITICAL: If ANY citation is FAKE or SUSPICIOUS, it MUST be replaced."""

        try:
            citation_response = call_gemini_reviewer(citation_check_prompt)
            yield f"data: {json.dumps({'type': 'log', 'message': '📋 Citation audit complete', 'professor': 'logic'})}\n\n"
            
            # Check for fake citations
            if "FAKE" in citation_response or "SUSPICIOUS" in citation_response:
                yield f"data: {json.dumps({'type': 'log', 'message': '⚠️ Fake citations detected! Requesting replacement...', 'professor': 'logic'})}\n\n"
        except Exception as e:
            citation_response = f"Citation check error: {str(e)}"
            yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Citation check error', 'professor': 'logic'})}\n\n"

        if round_num >= 3:
            recent_scores = scores[-2:]
            if len(recent_scores) == 2:
                improvement = recent_scores[-1] - recent_scores[-2]
                if improvement < STAGNATION_THRESHOLD:
                    stop_reason = f"STAGNATION: Improvement {improvement:.1f} < {STAGNATION_THRESHOLD}"
                    yield f"data: {json.dumps({'type': 'log', 'message': f'⏸️ {stop_reason}', 'professor': 'system'})}\n\n"
                    break

        if round_num == MAX_ROUNDS:
            stop_reason = f"LIMIT: Max {MAX_ROUNDS} rounds"
            yield f"data: {json.dumps({'type': 'log', 'message': f'🛑 {stop_reason}', 'professor': 'system'})}\n\n"
            break

        # Prof. Quill revises
        yield f"data: {json.dumps({'type': 'log', 'message': '✍️ Prof. Quill revising...', 'professor': 'quill'})}\n\n"
        
        refine_prompt = f"""Revise the essay based on feedback. CRITICAL: Replace any fake/suspicious citations with REAL ones.

ORIGINAL INSTRUCTIONS: {instructions}
CURRENT DRAFT: {current_draft}
EXAMINER FEEDBACK: {examiner_response}
CITATION AUDIT: {citation_response}

RULES:
- Maintain human-like writing (NO AI buzzwords)
- Replace ALL fake citations with verified sources
- Target {word_count_limit} words
- Aim for score > {SUCCESS_THRESHOLD}"""

        try:
            current_draft = call_claude_writer(refine_prompt)
            yield f"data: {json.dumps({'type': 'log', 'message': '✅ Revision complete', 'professor': 'quill'})}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'type': 'log', 'message': f'❌ Error: {str(e)}', 'professor': 'quill'})}\n\n"
            break

    if not stop_reason:
        stop_reason = "COMPLETED"
    
    yield f"data: {json.dumps({'type': 'log', 'message': f'🏁 Complete! {stop_reason}', 'professor': 'system'})}\n\n"
    yield f"data: {json.dumps({'type': 'log', 'message': f'🏆 Best Score: {best_score}/100 | Rounds: {len(scores)}', 'professor': 'system'})}\n\n"
    
    # Save to database
    try:
        essay = Essay(
            user_id=current_user.id,
            title=instructions[:100],
            instructions=instructions,
            final_content=best_draft,
            final_score=best_score,
            rounds_used=len(scores),
            stop_reason=stop_reason,
            word_count_limit=word_count_limit
        )
        db.session.add(essay)
        db.session.commit()
        
        yield f"data: {json.dumps({'type': 'complete', 'essay_id': essay.id, 'score': best_score, 'rounds': len(scores)})}\n\n"
    except Exception as e:
        yield f"data: {json.dumps({'type': 'error', 'message': f'Database error: {str(e)}', 'professor': 'system'})}\n\n"

# Routes
@app.route('/')
def index():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = User.query.filter_by(username=username).first()
        
        if user and user.check_password(password):
            login_user(user)
            flash('Login successful!', 'success')
            return redirect(url_for('dashboard'))
        else:
            flash('Invalid credentials', 'error')
    
    return render_template('login.html')

@app.route('/register', methods=['GET', 'POST'])
def register():
    if current_user.is_authenticated:
        return redirect(url_for('dashboard'))
    
    if request.method == 'POST':
        username = request.form.get('username')
        email = request.form.get('email')
        password = request.form.get('password')
        
        if User.query.filter_by(username=username).first():
            flash('Username exists', 'error')
            return redirect(url_for('register'))
        
        if User.query.filter_by(email=email).first():
            flash('Email registered', 'error')
            return redirect(url_for('register'))
        
        user = User(username=username, email=email)
        user.set_password(password)
        db.session.add(user)
        db.session.commit()
        
        flash('Registration successful!', 'success')
        return redirect(url_for('login'))
    
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    flash('Logged out', 'success')
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    essays = Essay.query.filter_by(user_id=current_user.id).order_by(Essay.created_at.desc()).limit(10).all()
    reports = Report.query.filter_by(user_id=current_user.id).order_by(Report.created_at.desc()).limit(10).all()
    return render_template('dashboard.html', essays=essays, reports=reports, user=current_user)

@app.route('/upload', methods=['POST'])
@login_required
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type'}), 400
    
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        extracted_text = parse_uploaded_file(filepath, filename)
        os.remove(filepath)
        
        return jsonify({'success': True, 'text': extracted_text})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/create-checkout-session', methods=['POST'])
@login_required
def create_checkout_session():
    try:
        data = request.get_json()
        tool_type = data.get('tool_type')
        
        if tool_type not in PRICING:
            return jsonify({'error': 'Invalid tool'}), 400
        
        pricing = PRICING[tool_type]
        
        # Store pending task
        session['pending_task'] = {
            'tool_type': tool_type,
            'data': data
        }
        
        product_names = {
            'essay_std': 'Essay Writing (Standard - 2000 words)',
            'essay_ext': 'Essay Writing (Extended - 4000 words)',
            'plagiarism': 'Plagiarism Check',
            'ai_check': 'AI Content Detection',
            'grader': 'Assignment Grading'
        }
        
        checkout_session = stripe.checkout.Session.create(
            payment_method_types=['card'],
            line_items=[{
                'price_data': {
                    'currency': 'gbp',
                    'unit_amount': pricing['price'],
                    'product_data': {
                        'name': product_names.get(tool_type, 'Academic Tool'),
                        'description': f'The Academic Board - {product_names.get(tool_type)}',
                    },
                },
                'quantity': 1,
            }],
            mode='payment',
            allow_promotion_codes=True,
            success_url=DOMAIN + f'/payment-success?session_id={{CHECKOUT_SESSION_ID}}&tool={tool_type}',
            cancel_url=DOMAIN + '/dashboard',
            client_reference_id=str(current_user.id),
            customer_email=current_user.email,
        )
        
        return jsonify({'id': checkout_session.id})
    except Exception as e:
        print(f"❌ Stripe Error: {str(e)}")
        return jsonify({'error': str(e)}), 403

@app.route('/payment-success')
@login_required
def payment_success():
    session_id = request.args.get('session_id')
    tool_type = request.args.get('tool')
    
    if session_id:
        try:
            checkout_session = stripe.checkout.Session.retrieve(session_id)
            
            if checkout_session.payment_status == 'paid':
                flash('Payment successful!', 'success')
                return redirect(url_for('dashboard') + f'?start_tool={tool_type}')
            else:
                flash('Payment failed', 'error')
        except Exception as e:
            flash(f'Error: {str(e)}', 'error')
    
    return redirect(url_for('dashboard'))

# Tool-specific generation routes
@app.route('/generate-essay', methods=['POST'])
@login_required
def generate_essay():
    instructions = request.form.get('instructions')
    word_count = int(request.form.get('word_count', 2000))
    
    if not instructions:
        return jsonify({'error': 'Instructions required'}), 400
    
    return Response(
        stream_with_context(generate_essay_with_citations_stream(instructions, word_count)),
        mimetype='text/event-stream',
        headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'}
    )

@app.route('/check-plagiarism', methods=['POST'])
@login_required
def check_plagiarism():
    text = request.form.get('text')
    
    if not text:
        return jsonify({'error': 'Text required'}), 400
    
    try:
        analysis = check_plagiarism_gemini(text)
        pdf_filename = generate_plagiarism_pdf(text, analysis, current_user.id)
        
        report = Report(
            user_id=current_user.id,
            tool_type='plagiarism',
            title='Plagiarism Check',
            result_data=analysis,
            file_path=pdf_filename
        )
        db.session.add(report)
        db.session.commit()
        
        return jsonify({
            'success': True,
            'report_id': report.id,
            'analysis': analysis,
            'pdf_url': f'/download-report/{report.id}'
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/check-ai', methods=['POST'])
@login_required
def check_ai():
    text = request.form.get('text')
    
    if not text:
        return jsonify({'error': 'Text required'}), 400
    
    try:
        analysis = check_ai_content(text)
        
        report = Report(
            user_id=current_user.id,
            tool_type='ai_check',
            title='AI Detection',
            result_data=analysis
        )
        db.session.add(report)
        db.session.commit()
        
        return jsonify({
            'success': True,
            'report_id': report.id,
            'analysis': analysis
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/grade-assignment', methods=['POST'])
@login_required
def grade_assignment_route():
    brief = request.form.get('brief')
    essay = request.form.get('essay')
    
    if not brief or not essay:
        return jsonify({'error': 'Both brief and essay required'}), 400
    
    try:
        grading = grade_assignment(brief, essay)
        
        report = Report(
            user_id=current_user.id,
            tool_type='grader',
            title='Assignment Grading',
            result_data=grading
        )
        db.session.add(report)
        db.session.commit()
        
        return jsonify({
            'success': True,
            'report_id': report.id,
            'grading': grading
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download-report/<int:report_id>')
@login_required
def download_report(report_id):
    report = Report.query.get_or_404(report_id)
    
    if report.user_id != current_user.id:
        flash('Unauthorized', 'error')
        return redirect(url_for('dashboard'))
    
    if report.file_path:
        filepath = os.path.join(app.config['REPORTS_FOLDER'], report.file_path)
        return send_file(filepath, as_attachment=True)
    else:
        return jsonify({'error': 'No file available'}), 404

@app.route('/download-essay/<int:essay_id>')
@login_required
def download_essay(essay_id):
    essay = Essay.query.get_or_404(essay_id)
    
    if essay.user_id != current_user.id:
        flash('Unauthorized', 'error')
        return redirect(url_for('dashboard'))
    
    doc = Document()
    title = doc.add_heading(essay.title, 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph(f"Generated: {essay.created_at.strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph(f"Score: {essay.final_score}/100")
    doc.add_paragraph(f"Rounds: {essay.rounds_used}")
    doc.add_paragraph("")
    
    doc.add_heading('Instructions', 1)
    doc.add_paragraph(essay.instructions)
    doc.add_paragraph("")
    
    doc.add_heading('Essay', 1)
    doc.add_paragraph(essay.final_content)
    
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    
    return Response(
        file_stream.getvalue(),
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        headers={'Content-Disposition': f'attachment; filename=essay_{essay_id}.docx'}
    )

# Initialize database
with app.app_context():
    print("🔧 Creating database...")
    db.create_all()
    print("✅ Database ready!")
    
    if not User.query.filter_by(username='demo').first():
        demo = User(username='demo', email='demo@example.com')
        demo.set_password('demo123')
        db.session.add(demo)
        db.session.commit()
        print("✅ Demo user created")

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000, threaded=True)