import os
import json
import time
import re
import requests
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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from reportlab.lib.units import inch

load_dotenv()

app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-secret-key-change-me')
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URI', 'sqlite:///academic_suite.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['REPORTS_FOLDER'] = 'reports'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['REPORTS_FOLDER'], exist_ok=True)

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

stripe_key = os.getenv('STRIPE_SECRET_KEY', '')
if stripe_key:
    stripe.api_key = stripe_key
    print(f"✅ Stripe Key Loaded")
else:
    print("⚠️ WARNING: STRIPE_SECRET_KEY not found")

DOMAIN = os.getenv('DOMAIN', 'http://localhost:5000')

# Updated pricing structure
PRICING = {
    'writer': {'price_per_1k': 1000, 'display': '£10 per 1,000 words'},  # Dynamic pricing
    'plagiarism': {'price': 1000, 'display': '£10'},
    'ai_check': {'price': 1000, 'display': '£10'},
    'grader': {'price': 1000, 'display': '£10'}  # Updated to £10
}

anthropic_client = anthropic.Anthropic(api_key=os.getenv('ANTHROPIC_API_KEY'))
openai.api_key = os.getenv('OPENAI_API_KEY')
genai.configure(api_key=os.getenv('GOOGLE_API_KEY'))

# MOCK EMAIL FUNCTION (For localhost development)
def send_email_mock(user_email, subject, body):
    """Mock email sender - prints to console instead of sending real emails"""
    print("\n" + "="*80)
    print("📧 MOCK EMAIL SENT (Console Log)")
    print("="*80)
    print(f"To: {user_email}")
    print(f"Subject: {subject}")
    print(f"Body:\n{body}")
    print("="*80 + "\n")

# Database Models
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    email = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    essays = db.relationship('Essay', backref='author', lazy=True)
    reports = db.relationship('Report', backref='author', lazy=True)
    payments = db.relationship('Payment', backref='user', lazy=True)

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
    word_count_limit = db.Column(db.Integer, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Report(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    tool_type = db.Column(db.String(50), nullable=False)
    title = db.Column(db.String(200), nullable=False)
    result_data = db.Column(db.Text, nullable=False)
    file_path = db.Column(db.String(500))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class Payment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey('user.id'), nullable=False)
    amount = db.Column(db.Integer, nullable=False)
    tool_type = db.Column(db.String(50), nullable=False)
    stripe_session_id = db.Column(db.String(200))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# File handling
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'pdf', 'docx', 'txt'}

def extract_text_from_pdf(file_path):
    """Extract text from PDF file"""
    with open(file_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        return "\n".join([page.extract_text() for page in pdf_reader.pages]).strip()

def extract_text_from_docx(file_path):
    """Extract text from DOCX file"""
    doc = Document(file_path)
    return "\n".join([p.text for p in doc.paragraphs]).strip()

def extract_text_from_file(file_path, filename):
    """Helper function to extract text from PDF or DOCX files"""
    ext = filename.rsplit('.', 1)[1].lower()
    if ext == 'pdf':
        return extract_text_from_pdf(file_path)
    elif ext == 'docx':
        return extract_text_from_docx(file_path)
    elif ext == 'txt':
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read().strip()
    raise Exception("Unsupported file type")

def parse_uploaded_file(file_path, filename):
    """Parse uploaded file and extract text"""
    return extract_text_from_file(file_path, filename)

# AI Helper Functions
def call_with_retry(func, max_retries=3):
    for attempt in range(max_retries):
        try:
            return func()
        except Exception as e:
            if attempt == max_retries - 1:
                raise
            time.sleep(2 * (attempt + 1))

def call_claude(prompt, max_tokens=4000):
    def api_call():
        response = anthropic_client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}]
        )
        return response.content[0].text
    return call_with_retry(api_call)

def call_gpt4(prompt, model="gpt-4o"):
    def api_call():
        response = openai.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        return response.choices[0].message.content
    return call_with_retry(api_call)

def call_gemini(prompt):
    def api_call():
        model = genai.GenerativeModel('gemini-1.5-pro')
        response = model.generate_content(prompt)
        return response.text
    return call_with_retry(api_call)

def extract_score(text):
    matches = re.findall(r'\b(\d{1,3})\b', text)
    for match in matches:
        score = int(match)
        if 0 <= score <= 100:
            return score
    return 0

def check_url_status(url):
    """Check if URL is live (HTTP 200) or broken"""
    try:
        response = requests.head(url, timeout=5, allow_redirects=True)
        return response.status_code == 200
    except:
        return False

# TOOL A: The Detective - 3-Agent Plagiarism Team
def plagiarism_hunter_gemini(text):
    """Step 1: Gemini hunts for sources with web search"""
    prompt = f"""You are The Hunter, a plagiarism detection expert with web search capabilities.

TEXT TO ANALYZE:
{text}

TASKS:
1. Search the web for potential sources of this text
2. Find matching passages
3. Extract URLs of sources
4. For each match, provide:
   - The suspicious passage
   - The source URL
   - Similarity percentage (0-100)

Return JSON format:
{{
    "matches": [
        {{
            "passage": "suspicious text",
            "source_url": "https://...",
            "similarity": 85
        }}
    ]
}}"""
    
    try:
        result = call_gemini(prompt)
        return result
    except Exception as e:
        return json.dumps({"matches": [], "error": str(e)})

def plagiarism_analyst_gpt(student_text, hunter_findings):
    """Step 2: GPT-4o analyzes similarity and paraphrasing"""
    prompt = f"""You are The Analyst, an expert in detecting plagiarism and paraphrasing.

STUDENT TEXT:
{student_text}

SOURCES FOUND BY HUNTER:
{hunter_findings}

TASKS:
1. Compare student text with each source
2. Detect direct copying vs paraphrasing
3. Assign "Suspicion Score" (0-100) for each match
4. Identify specific sentences that are problematic

Return JSON:
{{
    "overall_suspicion": 0-100,
    "analysis": [
        {{
            "student_sentence": "...",
            "source_sentence": "...",
            "suspicion_score": 0-100,
            "type": "DIRECT_COPY/PARAPHRASED/SIMILAR"
        }}
    ]
}}"""
    
    try:
        result = call_gpt4(prompt)
        return result
    except Exception as e:
        return json.dumps({"overall_suspicion": 0, "analysis": [], "error": str(e)})

def plagiarism_reporter_claude(student_text, hunter_data, analyst_data, user_id):
    """Step 3: Claude generates PDF report with highlighted text"""
    filename = f"plagiarism_report_{user_id}_{int(time.time())}.pdf"
    filepath = os.path.join(app.config['REPORTS_FOLDER'], filename)
    
    doc = SimpleDocTemplate(filepath, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []
    
    # Title
    title_style = ParagraphStyle('Title', parent=styles['Heading1'], 
                                 fontSize=24, textColor=colors.HexColor('#b537f2'))
    story.append(Paragraph("🔍 The Detective - Plagiarism Report", title_style))
    story.append(Spacer(1, 0.3*inch))
    
    # Metadata
    story.append(Paragraph(f"<b>Generated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles['Normal']))
    story.append(Spacer(1, 0.2*inch))
    
    # Parse findings
    try:
        hunter_json = json.loads(hunter_data)
        analyst_json = json.loads(analyst_data)
        
        overall_score = analyst_json.get('overall_suspicion', 0)
        verdict_color = colors.red if overall_score > 30 else colors.green
        
        verdict_style = ParagraphStyle('Verdict', parent=styles['Heading2'], textColor=verdict_color)
        story.append(Paragraph(f"Overall Suspicion: {overall_score}%", verdict_style))
        story.append(Spacer(1, 0.2*inch))
        
        # Suspicious passages
        story.append(Paragraph("<b>Suspicious Passages:</b>", styles['Heading3']))
        story.append(Spacer(1, 0.1*inch))
        
        for i, analysis in enumerate(analyst_json.get('analysis', [])[:10], 1):
            student_sent = analysis.get('student_sentence', 'N/A')
            suspicion = analysis.get('suspicion_score', 0)
            match_type = analysis.get('type', 'UNKNOWN')
            
            # Highlight in yellow for high suspicion
            if suspicion > 50:
                highlight_style = ParagraphStyle('Highlight', parent=styles['Normal'],
                                                backColor=colors.yellow)
                story.append(Paragraph(f"<b>Match {i} ({suspicion}% - {match_type}):</b>", styles['Normal']))
                story.append(Paragraph(f'"{student_sent}"', highlight_style))
            else:
                story.append(Paragraph(f"<b>Match {i} ({suspicion}% - {match_type}):</b>", styles['Normal']))
                story.append(Paragraph(f'"{student_sent}"', styles['Italic']))
            
            story.append(Spacer(1, 0.15*inch))
        
        # Reference Status Table
        story.append(PageBreak())
        story.append(Paragraph("<b>Reference Status:</b>", styles['Heading3']))
        story.append(Spacer(1, 0.1*inch))
        
        table_data = [['Source URL', 'Status']]
        for match in hunter_json.get('matches', [])[:15]:
            url = match.get('source_url', 'N/A')
            is_live = check_url_status(url)
            status = '✓ Live' if is_live else '✗ Broken'
            table_data.append([url[:60] + '...' if len(url) > 60 else url, status])
        
        table = Table(table_data, colWidths=[4*inch, 1.5*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        story.append(table)
        
    except Exception as e:
        story.append(Paragraph(f"Error parsing results: {str(e)}", styles['Normal']))
    
    doc.build(story)
    return filename

# TOOL B: The Architect - Essay Writer with SPARTAN ACADEMIC Protocol
def generate_essay_stream(instructions, word_count):
    yield f"data: {json.dumps({'type': 'log', 'message': f'🎓 The Academic Board convening ({word_count} words)...'})}\n\n"
    
    # Calculate buffer (15% extra)
    max_word_count = int(word_count * 1.15)
    
    # Initial draft
    yield f"data: {json.dumps({'type': 'log', 'message': '✍️ Prof. Quill drafting (Spartan precision)...'})}\n\n"
    
    writer_prompt = f"""You are Prof. Quill, operating under the SPARTAN ACADEMIC protocol.

CORE MANDATE: Write with Spartan precision. Be direct, authoritative, and concise. Avoid flowery language. Every sentence must add value.

CRITICAL WORD COUNT RULES:
1. TARGET: {word_count} words for the BODY TEXT ONLY (Introduction, Analysis, Conclusion)
2. MAXIMUM ALLOWED: {max_word_count} words for body text (+15% buffer)
3. EXCLUSION: The References/Bibliography section is EXCLUDED from this word count limit
4. DO NOT shorten the essay to fit references within the limit
5. Write the full essay body first, THEN add a complete References section separately

THE BAN LIST - STRICTLY FORBIDDEN WORDS (Essay FAILS if used):
❌ delve
❌ tapestry
❌ landscape
❌ leverage
❌ spearhead
❌ multifaceted
❌ underscore
❌ testament
❌ symphony
❌ rich
❌ realm
❌ myriad
❌ plethora
❌ paradigm
❌ robust

SPARTAN STYLE RULES:
- Direct, authoritative tone. No hedging ("perhaps", "might", "could be").
- Vary sentence length strategically. Short sentences drive points home. Longer sentences develop complex ideas.
- Avoid excessive transition words. Do NOT overuse: "Furthermore", "Moreover", "In conclusion", "Additionally".
- Use active voice. "The study proves X" not "X is proven by the study".
- Cut unnecessary words. "The fact that" → "That". "In order to" → "To".
- Contractions are acceptable when they strengthen voice (it's, don't, can't).
- Minor stylistic imperfections for authenticity (humans aren't perfect).

CONTENT REQUIREMENTS:
- Include REAL academic citations (Harvard/APA format)
- Every major claim needs a citation
- Use credible academic sources only
- Full References/Bibliography section at the end (NOT counted in word limit)

INSTRUCTIONS:
{instructions}

Remember: Write {word_count} words of sharp, direct body content. Then add references separately. No banned words. No fluff."""

    try:
        current_draft = call_claude(writer_prompt, max_tokens=6000)
        yield f"data: {json.dumps({'type': 'log', 'message': '✅ Initial draft complete'})}\n\n"
    except Exception as e:
        yield f"data: {json.dumps({'type': 'error', 'message': f'Error: {str(e)}'})}\n\n"
        return

    scores = []
    best_draft = current_draft
    best_score = 0

    for round_num in range(1, 6):
        yield f"data: {json.dumps({'type': 'log', 'message': f'━━━ Round {round_num}/5 ━━━'})}\n\n"
        
        # Dr. Strict grades
        yield f"data: {json.dumps({'type': 'log', 'message': '🎯 Dr. Strict grading...'})}\n\n"
        
        examiner_prompt = f"""Grade this essay (0-100). Apply SPARTAN ACADEMIC standards.

INSTRUCTIONS: {instructions}

ESSAY: {current_draft}

Evaluation criteria:
- Content quality and depth
- Argument structure and coherence
- Citation usage and academic rigor
- Writing style: Direct, authoritative, concise (Spartan)
- DEDUCT 20 POINTS if ANY banned words detected (delve, tapestry, landscape, leverage, spearhead, multifaceted, underscore, testament, symphony, rich, realm, myriad, plethora, paradigm, robust)

Provide a score (0-100) in format: 'Score: [number]'"""
        
        try:
            examiner_response = call_gpt4(examiner_prompt)
            score = extract_score(examiner_response)
            scores.append(score)
            
            yield f"data: {json.dumps({'type': 'score', 'round': round_num, 'score': score})}\n\n"
            verdict_msg = f"📊 Dr. Strict's verdict: {score}/100"
            yield f"data: {json.dumps({'type': 'log', 'message': verdict_msg})}\n\n"
            
            if score > best_score:
                best_score = score
                best_draft = current_draft
        except Exception as e:
            yield f"data: {json.dumps({'type': 'log', 'message': f'Error: {str(e)}'})}\n\n"
            score = 0
            scores.append(0)

        if score >= 90:
            yield f"data: {json.dumps({'type': 'log', 'message': '🎉 SUCCESS! High score achieved.'})}\n\n"
            break

        # Citation verification
        yield f"data: {json.dumps({'type': 'log', 'message': '🔍 Dean Logic verifying citations...'})}\n\n"
        
        citation_prompt = f"""Verify ALL citations in this essay. Check if papers/sources exist and are credible.

ESSAY: {current_draft}

For each citation:
1. Extract full citation details
2. Verify if source exists and is credible
3. Check relevance to the claim
4. Mark: VERIFIED / SUSPICIOUS / FAKE

Return JSON format with verification results and fake_count."""

        try:
            citation_response = call_gemini(citation_prompt)
            if "FAKE" in citation_response or "SUSPICIOUS" in citation_response:
                yield f"data: {json.dumps({'type': 'log', 'message': '⚠️ Fake/suspicious citations detected!'})}\n\n"
        except:
            citation_response = "Citation check error"

        # Revise
        yield f"data: {json.dumps({'type': 'log', 'message': '✍️ Prof. Quill revising (Spartan mode)...'})}\n\n"
        
        refine_prompt = f"""Revise the essay based on feedback. Replace any fake or suspicious citations with real ones. Apply SPARTAN ACADEMIC protocol.

CRITICAL WORD COUNT RULES (MUST FOLLOW):
1. TARGET: {word_count} words for BODY TEXT ONLY
2. MAXIMUM: {max_word_count} words for body text (+15% buffer)
3. References/Bibliography is EXCLUDED from word count
4. DO NOT cut essay content to fit references in the limit

THE BAN LIST - STRICTLY FORBIDDEN (Essay FAILS if used):
❌ delve, tapestry, landscape, leverage, spearhead, multifaceted, underscore, testament, symphony, rich, realm, myriad, plethora, paradigm, robust

SPARTAN STYLE:
- Direct, authoritative. No hedging.
- Vary sentence length. Short sentences punch. Longer ones develop.
- Cut transition word overuse (Furthermore, Moreover, Additionally).
- Active voice. Cut fluff.

ORIGINAL INSTRUCTIONS: {instructions}

CURRENT DRAFT: {current_draft}

EXAMINER FEEDBACK: {examiner_response}

CITATION VERIFICATION: {citation_response}

Improve the essay while maintaining {word_count} words of sharp body content (excluding references). Zero banned words. Zero fluff."""

        try:
            current_draft = call_claude(refine_prompt, max_tokens=6000)
            yield f"data: {json.dumps({'type': 'log', 'message': '✅ Revision complete'})}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'type': 'error', 'message': f'Error: {str(e)}'})}\n\n"
            break

    yield f"data: {json.dumps({'type': 'log', 'message': f'🏁 Complete! Best: {best_score}/100'})}\n\n"
    
    # Save
    try:
        essay = Essay(
            user_id=current_user.id,
            title=instructions[:100],
            instructions=instructions,
            final_content=best_draft,
            final_score=best_score,
            rounds_used=len(scores),
            word_count_limit=word_count
        )
        db.session.add(essay)
        db.session.commit()
        
        yield f"data: {json.dumps({'type': 'complete', 'essay_id': essay.id, 'score': best_score})}\n\n"
    except Exception as e:
        yield f"data: {json.dumps({'type': 'error', 'message': f'DB error: {str(e)}'})}\n\n"

# TOOL C: The Oracle (AI Detection)
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
        response = call_gpt4(prompt)
        return response
    except Exception as e:
        return json.dumps({"error": str(e), "ai_probability": 0, "verdict": "ERROR"})

# TOOL D: The Grader - 3-Round Consensus Debate
def grade_assignment(brief_text, essay_text):
    """Grade assignment using 3-round consensus debate among 3 professors"""
    
    # ROUND 1: Independent Initial Grading
    print("🎓 ROUND 1: Independent Grading...")
    
    # Prof. Quill - Content Analysis (Claude)
    quill_prompt_r1 = f"""You are Prof. Quill, an expert academic evaluator. Analyze this student essay against the assignment brief.

ASSIGNMENT BRIEF:
{brief_text}

STUDENT ESSAY:
{essay_text}

Evaluate independently:
1. How well does it address the brief?
2. Content quality and depth
3. Structure and organization

Provide:
- A score (0-100)
- Detailed feedback (2-3 paragraphs)
- Key strengths and weaknesses"""

    # Dr. Strict - Technical Grading (GPT-4o)
    strict_prompt_r1 = f"""You are Dr. Strict, a harsh academic grader. Grade this essay strictly and independently.

ASSIGNMENT BRIEF:
{brief_text}

STUDENT ESSAY:
{essay_text}

Evaluate independently:
1. Grammar and writing quality
2. Citation and referencing
3. Academic rigor

Provide:
- A score (0-100)
- Detailed feedback identifying specific weaknesses
- Technical issues found"""

    # Dean Logic - Overall Assessment (Gemini)
    logic_prompt_r1 = f"""You are Dean Logic, the final academic authority. Provide an independent overall assessment.

ASSIGNMENT BRIEF:
{brief_text}

STUDENT ESSAY:
{essay_text}

Provide independently:
- Overall score (0-100)
- Key strengths
- Critical improvements needed
- Holistic evaluation"""

    try:
        # Round 1: Get initial independent grades
        quill_response_r1 = call_claude(quill_prompt_r1)
        quill_score_r1 = extract_score(quill_response_r1)
        
        strict_response_r1 = call_gpt4(strict_prompt_r1)
        strict_score_r1 = extract_score(strict_response_r1)
        
        logic_response_r1 = call_gemini(logic_prompt_r1)
        logic_score_r1 = extract_score(logic_response_r1)
        
        print(f"Round 1 Scores - Quill: {quill_score_r1}, Strict: {strict_score_r1}, Logic: {logic_score_r1}")
        
        # ROUND 2: Consensus Phase - Professors read each other's feedback
        print("🎓 ROUND 2: Consensus Phase...")
        
        quill_prompt_r2 = f"""You are Prof. Quill. You've read the feedback from Dr. Strict and Dean Logic.

YOUR INITIAL ASSESSMENT:
Score: {quill_score_r1}/100
{quill_response_r1}

DR. STRICT'S ASSESSMENT:
Score: {strict_score_r1}/100
{strict_response_r1}

DEAN LOGIC'S ASSESSMENT:
Score: {logic_score_r1}/100
{logic_response_r1}

After reading your colleagues' feedback, reconsider your evaluation. Adjust your score if their points are valid. Provide:
- Revised score (0-100)
- Explanation of any changes
- Final detailed feedback"""

        strict_prompt_r2 = f"""You are Dr. Strict. You've read the feedback from Prof. Quill and Dean Logic.

YOUR INITIAL ASSESSMENT:
Score: {strict_score_r1}/100
{strict_response_r1}

PROF. QUILL'S ASSESSMENT:
Score: {quill_score_r1}/100
{quill_response_r1}

DEAN LOGIC'S ASSESSMENT:
Score: {logic_score_r1}/100
{logic_response_r1}

After reading your colleagues' feedback, reconsider your evaluation. Adjust your score if their points are valid. Provide:
- Revised score (0-100)
- Explanation of any changes
- Final detailed feedback"""

        logic_prompt_r2 = f"""You are Dean Logic. You've read the feedback from Prof. Quill and Dr. Strict.

YOUR INITIAL ASSESSMENT:
Score: {logic_score_r1}/100
{logic_response_r1}

PROF. QUILL'S ASSESSMENT:
Score: {quill_score_r1}/100
{quill_response_r1}

DR. STRICT'S ASSESSMENT:
Score: {strict_score_r1}/100
{strict_response_r1}

After reading your colleagues' feedback, reconsider your evaluation. Adjust your score if their points are valid. Provide:
- Revised score (0-100)
- Explanation of any changes
- Final detailed feedback"""

        # Round 2: Get revised consensus grades
        quill_response_r2 = call_claude(quill_prompt_r2)
        quill_score_r2 = extract_score(quill_response_r2)
        
        strict_response_r2 = call_gpt4(strict_prompt_r2)
        strict_score_r2 = extract_score(strict_response_r2)
        
        logic_response_r2 = call_gemini(logic_prompt_r2)
        logic_score_r2 = extract_score(logic_response_r2)
        
        print(f"Round 2 Scores - Quill: {quill_score_r2}, Strict: {strict_score_r2}, Logic: {logic_score_r2}")
        
        # ROUND 3: Final Calculation
        print("🎓 ROUND 3: Final Calculation...")
        
        # Use Round 2 scores (post-consensus) for final average
        average_score = round((quill_score_r2 + strict_score_r2 + logic_score_r2) / 3)
        
        result = {
            "average_score": average_score,
            "quill_score": quill_score_r2,
            "quill_feedback": quill_response_r2,
            "quill_initial_score": quill_score_r1,
            "strict_score": strict_score_r2,
            "strict_feedback": strict_response_r2,
            "strict_initial_score": strict_score_r1,
            "logic_score": logic_score_r2,
            "logic_feedback": logic_response_r2,
            "logic_initial_score": logic_score_r1,
            "debate_summary": f"Round 1: Quill {quill_score_r1}, Strict {strict_score_r1}, Logic {logic_score_r1}. Round 2 (Consensus): Quill {quill_score_r2}, Strict {strict_score_r2}, Logic {logic_score_r2}. Final Average: {average_score}/100"
        }
        
        print(f"✅ Final Average Score: {average_score}/100")
        
        return json.dumps(result)
    except Exception as e:
        print(f"❌ Error in grading: {str(e)}")
        return json.dumps({"error": str(e), "average_score": 0})

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
            return redirect(url_for('dashboard'))
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
        confirm_password = request.form.get('confirm_password')
        
        # Check if passwords match
        if password != confirm_password:
            flash('Passwords do not match!', 'error')
            return redirect(url_for('register'))
        
        if User.query.filter_by(username=username).first():
            flash('Username exists', 'error')
            return redirect(url_for('register'))
        
        user = User(username=username, email=email)
        user.set_password(password)
        db.session.add(user)
        db.session.commit()
        
        # Send mock welcome email
        email_body = f"""
Welcome to The Academic Board, {username}!

Your account has been successfully created.

Get started with our premium AI-powered academic tools:
- The Architect: AI Essay Writer
- The Detective: Plagiarism Checker
- The Oracle: AI Content Detector
- The Grader: Assignment Marking

Login now at: {DOMAIN}

Best regards,
The Academic Board Team
        """
        send_email_mock(email, "Welcome to The Academic Board!", email_body)
        
        flash('Welcome aboard! (Check console for mock email)', 'success')
        return redirect(url_for('login'))
    
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    return render_template('dashboard.html', user=current_user)

@app.route('/tools/writer')
@login_required
def tool_writer():
    return render_template('tool_writer.html', user=current_user)

@app.route('/tools/plagiarism')
@login_required
def tool_plagiarism():
    return render_template('tool_plagiarism.html', user=current_user)

@app.route('/tools/detector')
@login_required
def tool_detector():
    return render_template('tool_detector.html', user=current_user)

@app.route('/tools/grader')
@login_required
def tool_grader():
    return render_template('tool_grader.html', user=current_user)

@app.route('/billing')
@login_required
def billing():
    payments = Payment.query.filter_by(user_id=current_user.id).order_by(Payment.created_at.desc()).all()
    return render_template('billing.html', payments=payments, user=current_user)

@app.route('/my-essays')
@login_required
def my_essays():
    essays = Essay.query.filter_by(user_id=current_user.id).order_by(Essay.created_at.desc()).all()
    return render_template('my_essays.html', essays=essays, user=current_user)

@app.route('/settings')
@login_required
def settings():
    return render_template('settings.html', user=current_user)

# API Routes
@app.route('/api/support', methods=['POST'])
@login_required
def support_chat():
    """The Concierge - AI Support Bot with Error Handling"""
    try:
        data = request.get_json()
        user_message = data.get('message', '')
        
        if not user_message:
            return jsonify({'reply': 'Please provide a message.'}), 400
        
        # Check if OpenAI API key exists
        if not os.getenv('OPENAI_API_KEY'):
            return jsonify({'reply': 'System Error: Contact Admin - OpenAI API key not configured.'}), 500
        
        system_prompt = """You are The Concierge, the helpful support agent for The Academic Board.

PRICING:
- Essay Writer: £10 per 1,000 words (1,000-12,000 words range)
- Plagiarism Check: £10
- AI Detector: £10
- Assignment Grader: £10

TOOLS:
- The Architect: AI essay writer with dynamic pricing and citation verification
- The Detective: Plagiarism checker with PDF reports
- The Oracle: AI content detector
- The Grader: Strict assignment marking (now £10)

Be polite, concise, and helpful. Troubleshoot errors and explain features."""
        
        response = openai.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_message}
            ],
            temperature=0.7,
            max_tokens=300
        )
        
        bot_reply = response.choices[0].message.content
        return jsonify({'reply': bot_reply})
        
    except openai.AuthenticationError:
        return jsonify({'reply': 'System Error: Contact Admin - Invalid API credentials.'}), 500
    except openai.RateLimitError:
        return jsonify({'reply': 'System Error: Contact Admin - Rate limit exceeded.'}), 500
    except openai.APIError as e:
        return jsonify({'reply': f'System Error: Contact Admin - API error occurred.'}), 500
    except Exception as e:
        return jsonify({'reply': 'System Error: Contact Admin - An unexpected error occurred.'}), 500

@app.route('/upload', methods=['POST'])
@login_required
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    
    file = request.files['file']
    if not file.filename or not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file'}), 400
    
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        text = extract_text_from_file(filepath, filename)
        os.remove(filepath)
        
        return jsonify({'success': True, 'text': text})
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
        
        # Calculate amount based on tool type
        if tool_type == 'writer':
            # Dynamic pricing for writer
            word_count = data.get('word_count', 2000)
            amount = int((word_count / 1000) * PRICING['writer']['price_per_1k'])
            product_name = f'Essay Writing ({word_count:,} words)'
        else:
            # Fixed pricing for other tools
            amount = PRICING[tool_type]['price']
            product_names = {
                'plagiarism': 'Plagiarism Check',
                'ai_check': 'AI Detection',
                'grader': 'Assignment Grading'
            }
            product_name = product_names.get(tool_type, 'Academic Tool')
        
        session['pending_task'] = {'tool_type': tool_type, 'data': data}
        
        checkout_session = stripe.checkout.Session.create(
            payment_method_types=['card'],
            line_items=[{
                'price_data': {
                    'currency': 'gbp',
                    'unit_amount': amount,
                    'product_data': {'name': product_name},
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
                # Record payment
                payment = Payment(
                    user_id=current_user.id,
                    amount=checkout_session.amount_total,
                    tool_type=tool_type,
                    stripe_session_id=session_id
                )
                db.session.add(payment)
                db.session.commit()
                
                flash('Payment successful!', 'success')
                
                # Redirect to appropriate tool
                tool_routes = {
                    'writer': 'tool_writer',
                    'plagiarism': 'tool_plagiarism',
                    'ai_check': 'tool_detector',
                    'grader': 'tool_grader'
                }
                return redirect(url_for(tool_routes.get(tool_type, 'dashboard')) + '?paid=true')
        except Exception as e:
            flash(f'Error: {str(e)}', 'error')
    
    return redirect(url_for('dashboard'))

@app.route('/generate-essay', methods=['POST'])
@login_required
def generate_essay():
    instructions = request.form.get('instructions')
    word_count = int(request.form.get('word_count', 2000))
    
    if not instructions:
        return jsonify({'error': 'Instructions required'}), 400
    
    return Response(
        stream_with_context(generate_essay_stream(instructions, word_count)),
        mimetype='text/event-stream',
        headers={'Cache-Control': 'no-cache'}
    )

@app.route('/check-plagiarism', methods=['POST'])
@login_required
def check_plagiarism():
    text = request.form.get('text')
    
    if not text:
        return jsonify({'error': 'Text required'}), 400
    
    try:
        # Step 1: Hunter (Gemini)
        hunter_findings = plagiarism_hunter_gemini(text)
        
        # Step 2: Analyst (GPT-4o)
        analyst_findings = plagiarism_analyst_gpt(text, hunter_findings)
        
        # Step 3: Reporter (Claude) - Generate PDF
        pdf_filename = plagiarism_reporter_claude(text, hunter_findings, analyst_findings, current_user.id)
        
        # Save report
        report = Report(
            user_id=current_user.id,
            tool_type='plagiarism',
            title='Plagiarism Check',
            result_data=analyst_findings,
            file_path=pdf_filename
        )
        db.session.add(report)
        db.session.commit()
        
        return jsonify({
            'success': True,
            'report_id': report.id,
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
    
    return jsonify({'error': 'No file'}), 404

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