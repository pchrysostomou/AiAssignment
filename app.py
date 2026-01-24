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
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024  # 25MB

# Add Stripe Publishable Key to config so templates can access it
app.config['STRIPE_PUBLISHABLE_KEY'] = os.getenv('STRIPE_PUBLISHABLE_KEY', '')

# CRITICAL FIX: Set permissive CSP that allows Stripe.js and eval() for payment processing
@app.after_request
def add_security_headers(response):
    """
    Configure Content Security Policy to allow:
    1. Stripe.js library and checkout iframe
    2. unsafe-eval for Stripe's payment processing
    3. Our own domain for API calls
    """
    # Get the current domain
    domain = os.getenv('DOMAIN', 'https://aiassignment-x631.onrender.com')
    
    # Build CSP policy that allows Stripe while maintaining reasonable security
    csp_policy = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline' 'unsafe-eval' https://js.stripe.com https://cdn.jsdelivr.net; "
        "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
        "font-src 'self' https://fonts.gstatic.com; "
        "img-src 'self' data: https:; "
        f"connect-src 'self' {domain} https://api.stripe.com https://www.googleapis.com; "
        "frame-src 'self' https://js.stripe.com https://hooks.stripe.com; "
        "frame-ancestors 'self';"
    )
    
    # Set CSP header
    response.headers['Content-Security-Policy'] = csp_policy
    
    # Add CORS headers for API routes
    if request.path.startswith('/api/') or request.path.startswith('/generate-') or request.path.startswith('/create-checkout') or request.path.startswith('/check-'):
        response.headers['Access-Control-Allow-Origin'] = '*'
        response.headers['Access-Control-Allow-Methods'] = 'GET, POST, PUT, DELETE, OPTIONS'
        response.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization'
        response.headers['Access-Control-Max-Age'] = '3600'
    
    # Prevent caching for API routes
    if request.path.startswith('/api/') or request.path.startswith('/generate-') or request.path.startswith('/create-checkout'):
        response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
        response.headers['Pragma'] = 'no-cache'
        response.headers['Expires'] = '0'
    
    return response

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['REPORTS_FOLDER'], exist_ok=True)

db = SQLAlchemy(app)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'

# Stripe initialization - uses environment variable
stripe_key = os.getenv('STRIPE_SECRET_KEY', '')
if stripe_key:
    stripe.api_key = stripe_key
    print(f"✅ Stripe Secret Key Loaded (ending: ...{stripe_key[-4:]})")
else:
    print("⚠️ WARNING: STRIPE_SECRET_KEY not found in environment variables")

# Log Stripe Publishable Key status
stripe_pub_key = os.getenv('STRIPE_PUBLISHABLE_KEY', '')
if stripe_pub_key:
    print(f"✅ Stripe Publishable Key Loaded (ending: ...{stripe_pub_key[-4:]})")
else:
    print("⚠️ WARNING: STRIPE_PUBLISHABLE_KEY not found in environment variables")

# Hardcode production domain for Stripe redirects
DOMAIN = os.getenv('DOMAIN', 'https://aiassignment-x631.onrender.com')

# Updated pricing structure
PRICING = {
    'writer': {'price_per_1k': 1000, 'display': '£10 per 1,000 words'},  # Dynamic pricing
    'plagiarism': {'price': 1000, 'display': '£10'},
    'ai_check': {'price': 1000, 'display': '£10'},
    'grader': {'price': 1000, 'display': '£10'}
}

# API Credentials Initialization with Validation
anthropic_api_key = os.getenv('ANTHROPIC_API_KEY', '')
if anthropic_api_key:
    anthropic_client = anthropic.Anthropic(
        api_key=anthropic_api_key,
        default_headers={"anthropic-version": "2023-06-01"}
    )
    print(f"✅ Anthropic API Key Loaded (ending: ...{anthropic_api_key[-4:]})")
    print(f"✅ Anthropic Client Initialized with API version: 2023-06-01")
else:
    anthropic_client = None
    print("⚠️ WARNING: ANTHROPIC_API_KEY not found in environment variables")

openai_api_key = os.getenv('OPENAI_API_KEY', '')
if openai_api_key:
    openai.api_key = openai_api_key
    print(f"✅ OpenAI API Key Loaded (ending: ...{openai_api_key[-4:]})")
else:
    print("⚠️ WARNING: OPENAI_API_KEY not found in environment variables")

google_api_key = os.getenv('GOOGLE_API_KEY', '')
if google_api_key:
    genai.configure(api_key=google_api_key)
    print(f"✅ Google API Key Loaded (ending: ...{google_api_key[-4:]})")
else:
    print("⚠️ WARNING: GOOGLE_API_KEY not found in environment variables")

# Google Custom Search Engine Configuration
GOOGLE_CSE_ID = str(os.getenv('GOOGLE_CSE_ID', '50f552d88c3e14772')).strip()
if GOOGLE_CSE_ID and GOOGLE_CSE_ID != '':
    print(f"✅ Google CSE ID Configured: '{GOOGLE_CSE_ID}'")
else:
    print("⚠️ WARNING: GOOGLE_CSE_ID not found or empty in environment variables")

# SIMPLIFIED GEMINI MODEL CONFIGURATION
# With google-generativeai 0.8.3+, this standard alias works perfectly
GEMINI_MODEL = "gemini-1.5-flash"
print(f"🎯 GEMINI MODEL SET: {GEMINI_MODEL} (Standard Alias - Works with google-generativeai>=0.8.3)")

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
    """Extract text from PDF file with MEMORY-EFFICIENT streaming approach."""
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            total_pages = len(pdf_reader.pages)
            max_pages = min(total_pages, 100)
            
            text_chunks = []
            for i in range(max_pages):
                try:
                    page_text = pdf_reader.pages[i].extract_text()
                    if page_text:
                        text_chunks.append(page_text)
                except Exception as e:
                    print(f"⚠️ Warning: Could not extract page {i+1}: {str(e)}")
                    continue
            
            text = "\n".join(text_chunks)
            
            if total_pages > max_pages:
                text += f"\n\n[Note: Document has {total_pages} pages. First {max_pages} pages extracted.]"
            
            return text.strip()
    except Exception as e:
        raise Exception(f"PDF extraction error: {str(e)}")

def extract_text_from_docx(file_path):
    """Extract text from DOCX file with MEMORY-EFFICIENT approach."""
    try:
        doc = Document(file_path)
        total_paragraphs = len(doc.paragraphs)
        max_paragraphs = min(total_paragraphs, 300)
        
        text_chunks = []
        for i in range(max_paragraphs):
            para_text = doc.paragraphs[i].text
            if para_text.strip():
                text_chunks.append(para_text)
        
        text = "\n".join(text_chunks)
        
        if total_paragraphs > max_paragraphs:
            text += f"\n\n[Note: Document has {total_paragraphs} paragraphs. First {max_paragraphs} extracted.]"
        
        return text.strip()
    except Exception as e:
        raise Exception(f"DOCX extraction error: {str(e)}")

def extract_text_from_file(file_path, filename):
    """Helper function to extract text from PDF or DOCX files."""
    file_size = os.path.getsize(file_path)
    file_size_mb = file_size / (1024 * 1024)
    
    if file_size > 25 * 1024 * 1024:
        raise Exception(f"File too large ({file_size_mb:.1f}MB). Maximum 25MB per file.")
    
    print(f"📄 Processing file: {filename} ({file_size_mb:.1f}MB)")
    
    ext = filename.rsplit('.', 1)[1].lower()
    if ext == 'pdf':
        return extract_text_from_pdf(file_path)
    elif ext == 'docx':
        return extract_text_from_docx(file_path)
    elif ext == 'txt':
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read(200 * 1024)
            return text.strip()
    raise Exception("Unsupported file type")

def parse_uploaded_file(file_path, filename):
    """Parse uploaded file and extract text"""
    return extract_text_from_file(file_path, filename)

# Google Custom Search Integration
def google_search(query, num_results=5):
    """Perform Google Custom Search using the configured API key and CSE ID."""
    if not google_api_key or not GOOGLE_CSE_ID or GOOGLE_CSE_ID == '':
        print(f"⚠️ Google Search unavailable")
        return []
    
    try:
        words = query.split()
        key_words = [w for w in words if len(w) > 1][:5]
        
        if len(key_words) < 2:
            search_query = query[:80]
        else:
            search_query = ' '.join(key_words)
        
        url = "https://www.googleapis.com/customsearch/v1"
        params = {
            'key': str(google_api_key).strip(),
            'cx': str(GOOGLE_CSE_ID).strip(),
            'q': search_query,
            'num': num_results
        }
        
        print(f"🔍 Google Search Query: '{search_query}'")
        
        response = requests.get(url, params=params, timeout=10)
        
        if response.status_code != 200:
            print(f"❌ Google Search API Error: {response.status_code}")
            return []
        
        data = response.json()
        results = []
        
        for item in data.get('items', []):
            results.append({
                'title': item.get('title', ''),
                'link': item.get('link', ''),
                'snippet': item.get('snippet', '')
            })
        
        print(f"✅ Google Search completed: {len(results)} results")
        return results
        
    except Exception as e:
        print(f"❌ Google Search error: {str(e)}")
        return []

# AI Helper Functions - SIMPLIFIED CONFIGURATION
def call_with_retry(func, max_retries=3):
    for attempt in range(max_retries):
        try:
            return func()
        except Exception as e:
            if attempt == max_retries - 1:
                raise
            time.sleep(2 * (attempt + 1))

def call_claude(prompt, max_tokens=8000):
    """
    HIGH-END CONFIGURATION for Claude:
    PRIMARY: claude-3-5-sonnet-20241022 (Tier 1 High-Quality)
    SAFETY NET: claude-3-haiku-20240307 (only if Sonnet fails - prevents crashes)
    """
    if not anthropic_client:
        raise Exception("Anthropic API key not configured")
    
    # PRIMARY MODEL: Claude 3.5 Sonnet 20241022 (Tier 1)
    try:
        print(f"🤖 Calling Claude 3.5 Sonnet (20241022) - HIGH-END MODE")
        response = anthropic_client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": prompt}]
        )
        print(f"✅ Claude 3.5 Sonnet (20241022) API call successful")
        return response.content[0].text
    except anthropic.NotFoundError as e:
        print(f"⚠️ Claude Sonnet (20241022) 404 Error: {str(e)}")
        print(f"🔄 SAFETY NET: Falling back to claude-3-haiku-20240307...")
    except Exception as e:
        print(f"⚠️ Claude Sonnet (20241022) Error: {str(e)}")
        print(f"🔄 SAFETY NET: Falling back to claude-3-haiku-20240307...")
    
    # SAFETY NET: Claude 3 Haiku (only if Sonnet fails)
    try:
        print(f"🤖 SAFETY NET: Calling Claude 3 Haiku (20240307)...")
        response = anthropic_client.messages.create(
            model="claude-3-haiku-20240307",
            max_tokens=min(max_tokens, 4096),  # Haiku max is 4096
            messages=[{"role": "user", "content": prompt}]
        )
        print(f"✅ Claude 3 Haiku (SAFETY NET) API call successful")
        return response.content[0].text
    except Exception as e:
        print(f"❌ ALL CLAUDE MODELS FAILED: {str(e)}")
        raise Exception(f"All Claude models failed (Sonnet 20241022, Haiku): {str(e)}")

def call_gpt4(prompt, model="gpt-4o"):
    """
    HIGH-END CONFIGURATION for OpenAI:
    MODEL: gpt-4o (active for consensus loop)
    """
    if not openai_api_key:
        raise Exception("OpenAI API key not configured")
    
    def api_call():
        print(f"🤖 Calling OpenAI GPT-4o - HIGH-END MODE")
        response = openai.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        print(f"✅ OpenAI GPT-4o API call successful")
        return response.choices[0].message.content
    return call_with_retry(api_call)

def call_gemini(prompt):
    """
    SIMPLIFIED CONFIGURATION for Gemini:
    Uses gemini-1.5-flash (standard alias that works with google-generativeai>=0.8.3)
    
    NO MORE 404 ERRORS - using standard model alias supported by updated library.
    """
    if not google_api_key:
        raise Exception("Google API key not configured")
    
    try:
        print(f"🤖 Calling Google Gemini ({GEMINI_MODEL}) - STANDARD ALIAS MODE")
        model = genai.GenerativeModel(GEMINI_MODEL)
        response = model.generate_content(prompt)
        print(f"✅ Google Gemini ({GEMINI_MODEL}) API call successful")
        return response.text
    except Exception as e:
        print(f"❌ Gemini Error: {str(e)}")
        raise Exception(f"Gemini failed: {str(e)}")

def extract_score(text):
    """Extract score from agent response"""
    matches = re.findall(r'\b(\d{1,3})\b', text)
    for match in matches:
        score = int(match)
        if 0 <= score <= 100:
            return score
    return 0

def count_words(text):
    """Count words in text, excluding references section"""
    ref_markers = ['references', 'bibliography', 'works cited']
    text_lower = text.lower()
    
    for marker in ref_markers:
        if f'\n{marker}\n' in text_lower or f'\n{marker}:' in text_lower:
            parts = re.split(f'\n{marker}[:\n]', text_lower, maxsplit=1)
            if len(parts) > 1:
                body_text = parts[0]
                words = body_text.split()
                return len(words)
    
    words = text.split()
    return len(words)

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

# TOOL B: The Architect - SIMPLIFIED STRICT CONSENSUS 3-AGENT SYSTEM
def generate_essay_stream(instructions, word_count):
    """
    SIMPLIFIED STRICT CONSENSUS LOOP:
    - Prof. Quill (Claude 3.5 Sonnet 20241022 - Tier 1)
    - Dean Logic (Gemini 1.5 Flash - Standard Alias)
    - Chancellor GPT (GPT-4o)
    - Loop continues until ALL 3 agents score 80+ in SAME round
    - Target: 2,200 words total to ensure 2,000+ body words
    
    NO MORE 404 ERRORS - using standard gemini-1.5-flash alias with updated library.
    """
    yield f"data: {json.dumps({'type': 'log', 'message': '🎓 The Academic Board - CLEAN DEPENDENCY STRICT CONSENSUS MODE (3 Premium Agents)'})}\n\n"
    
    # Target 2200 words total to ensure 2000+ body words after references
    target_total_words = 2200
    
    # Research Phase
    yield f"data: {json.dumps({'type': 'log', 'message': '🔍 Researching topic via Google Custom Search...'})}\n\n"
    
    research_context = ""
    try:
        search_query = instructions[:150]
        search_results = google_search(search_query, num_results=5)
        
        if search_results:
            research_context = "\n\nRESEARCH CONTEXT (from web search):\n"
            for idx, result in enumerate(search_results, 1):
                research_context += f"{idx}. {result['title']}\n   Source: {result['link']}\n   Summary: {result['snippet']}\n\n"
            yield f"data: {json.dumps({'type': 'log', 'message': f'✅ Found {len(search_results)} relevant sources'})}\n\n"
        else:
            yield f"data: {json.dumps({'type': 'log', 'message': '⚠️ No search results, proceeding with general knowledge'})}\n\n"
    except Exception as e:
        yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Research phase error: {str(e)}'})}\n\n"
    
    # Initial draft by Prof. Quill (Claude 3.5 Sonnet 20241022)
    yield f"data: {json.dumps({'type': 'log', 'message': '✍️ Prof. Quill (Claude 3.5 Sonnet 20241022) drafting initial essay...'})}\n\n"
    
    writer_prompt = f"""You are Prof. Quill, operating under the SPARTAN ACADEMIC protocol.

CRITICAL WORD COUNT TARGET:
- TARGET: {target_total_words} words TOTAL (to ensure {word_count}+ body words after references)
- Write the full essay body first, THEN add a complete References section separately

SMART CITATION LOGIC - Analyze the assignment type:
- CRITICAL REVIEW of a single paper/book: NO external citations needed (only cite the work being reviewed)
- LITERATURE REVIEW: MANY citations needed (10-20+ sources)
- RESEARCH ESSAY: MODERATE citations (5-10 sources)
- OPINION PIECE: FEW citations (2-5 sources for key claims only)
- COMPARATIVE ANALYSIS: MODERATE citations (cite each work being compared)

THE BAN LIST - STRICTLY FORBIDDEN WORDS:
❌ delve, tapestry, landscape, leverage, spearhead, multifaceted, underscore, testament, symphony, rich, realm, myriad, plethora, paradigm, robust

SPARTAN STYLE RULES:
- Direct, authoritative tone. No hedging.
- Vary sentence length strategically.
- Avoid excessive transition words.
- Use active voice.
- Cut unnecessary words.
- Contractions acceptable when they strengthen voice.
- Minor stylistic imperfections for authenticity.

INSTRUCTIONS:
{instructions}

{research_context}

Write {target_total_words} words total. Apply smart citation logic. No banned words. No fluff."""

    try:
        current_draft = call_claude(writer_prompt, max_tokens=8000)
        current_word_count = count_words(current_draft)
        yield f"data: {json.dumps({'type': 'log', 'message': f'✅ Initial draft complete ({current_word_count} body words)'})}\n\n"
    except Exception as e:
        yield f"data: {json.dumps({'type': 'error', 'message': f'Error: {str(e)}'})}\n\n"
        return

    # SIMPLIFIED STRICT CONSENSUS LOOP
    best_draft = current_draft
    best_avg_score = 0
    round_count = 0
    MAX_ROUNDS = 15  # Hard limit safety break
    
    # Store specific rewrites from critics
    logic_rewrites = ""
    gpt_rewrites = ""
    
    while round_count < MAX_ROUNDS:
        round_count += 1
        current_word_count = count_words(current_draft)
        yield f"data: {json.dumps({'type': 'log', 'message': f'━━━ ROUND {round_count}/{MAX_ROUNDS} ({current_word_count}/{word_count} body words) ━━━'})}\n\n"
        
        # CRITIC 1: Dean Logic (Gemini 1.5 Flash)
        yield f"data: {json.dumps({'type': 'log', 'message': f'⚖️ Dean Logic (Gemini {GEMINI_MODEL}) evaluating...'})}\n\n"
        
        logic_prompt = f"""You are Dean Logic, a harsh academic critic. Grade this essay strictly.

INSTRUCTIONS: {instructions}

ESSAY: {current_draft}

Evaluation criteria:
- Overall coherence and persuasiveness
- Citation appropriateness (smart citation logic applied)
- Meeting assignment requirements
- Academic integrity and originality
- Word count: Target {word_count}+ body words (current: {current_word_count} body words)
- DEDUCT 20 POINTS if ANY banned words detected
- DEDUCT 15 POINTS if body word count is below {word_count}

CRITICAL REQUIREMENT - ACTIVE CONTRIBUTION:
If you score below 80, you MUST provide:
1. Your score (0-100) in format: 'Score: [number]'
2. Specific rewritten paragraphs or new citations/arguments that MUST be added
3. Label your contributions clearly as "DEAN LOGIC REWRITES:"

Example format if score < 80:
Score: 75

DEAN LOGIC REWRITES:
[Paragraph 3 should be replaced with:]
"The evidence suggests that X is correlated with Y. Smith (2023) demonstrates this through longitudinal analysis..."

[Add new citation after paragraph 5:]
"Furthermore, recent studies by Johnson et al. (2024) indicate..."

Provide score and specific rewrites if needed."""

        try:
            logic_response = call_gemini(logic_prompt)
            score_logic = extract_score(logic_response)
            yield f"data: {json.dumps({'type': 'log', 'message': f'📊 Dean Logic: {score_logic}/100'})}\n\n"
            
            # Extract rewrites if score < 80
            if score_logic < 80 and "DEAN LOGIC REWRITES:" in logic_response:
                logic_rewrites = logic_response.split("DEAN LOGIC REWRITES:")[1].strip()
                yield f"data: {json.dumps({'type': 'log', 'message': '📝 Dean Logic provided specific rewrites'})}\n\n"
            else:
                logic_rewrites = ""
                
        except Exception as e:
            yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Dean Logic error: {str(e)}'})}\n\n"
            score_logic = 0
            logic_response = "Error during grading"
            logic_rewrites = ""
        
        # CRITIC 2: Chancellor GPT (GPT-4o)
        yield f"data: {json.dumps({'type': 'log', 'message': '🎓 Chancellor GPT (GPT-4o) evaluating...'})}\n\n"
        
        gpt_prompt = f"""You are Chancellor GPT, the supreme academic auditor. Grade this essay with the highest standards.

INSTRUCTIONS: {instructions}

ESSAY: {current_draft}

Evaluation criteria:
- Overall excellence and scholarly merit
- Argumentation strength and evidence quality
- Professional writing standards
- Readability and engagement
- Word count: Target {word_count}+ body words (current: {current_word_count} body words)
- DEDUCT 20 POINTS if ANY banned words detected
- DEDUCT 15 POINTS if body word count is below {word_count}

CRITICAL REQUIREMENT - ACTIVE CONTRIBUTION:
If you score below 80, you MUST provide:
1. Your score (0-100) in format: 'Score: [number]'
2. Specific rewritten paragraphs or new citations/arguments that MUST be added
3. Label your contributions clearly as "CHANCELLOR GPT REWRITES:"

Example format if score < 80:
Score: 72

CHANCELLOR GPT REWRITES:
[Introduction needs stronger hook:]
"In the contemporary discourse on X, scholars have increasingly recognized..."

[Paragraph 7 lacks evidence:]
"According to Williams (2023), the correlation between A and B is statistically significant (p < 0.01)..."

Provide score and specific rewrites if needed."""

        try:
            gpt_response = call_gpt4(gpt_prompt)
            score_gpt = extract_score(gpt_response)
            yield f"data: {json.dumps({'type': 'log', 'message': f'📊 Chancellor GPT: {score_gpt}/100'})}\n\n"
            
            # Extract rewrites if score < 80
            if score_gpt < 80 and "CHANCELLOR GPT REWRITES:" in gpt_response:
                gpt_rewrites = gpt_response.split("CHANCELLOR GPT REWRITES:")[1].strip()
                yield f"data: {json.dumps({'type': 'log', 'message': '📝 Chancellor GPT provided specific rewrites'})}\n\n"
            else:
                gpt_rewrites = ""
                
        except Exception as e:
            yield f"data: {json.dumps({'type': 'log', 'message': f'⚠️ Chancellor GPT error: {str(e)}'})}\n\n"
            score_gpt = 0
            gpt_response = "Error during grading"
            gpt_rewrites = ""
        
        # Calculate average score (only from critics, writer doesn't self-grade)
        avg_score = round((score_logic + score_gpt) / 2)
        yield f"data: {json.dumps({'type': 'score', 'round': round_count, 'score': avg_score})}\n\n"
        yield f"data: {json.dumps({'type': 'log', 'message': f'📊 Round {round_count} Average: {avg_score}/100 (Logic: {score_logic}, GPT: {score_gpt})'})}\n\n"
        
        # Track best draft
        if avg_score > best_avg_score:
            best_avg_score = avg_score
            best_draft = current_draft
        
        # CHECK CONSENSUS - ALL 3 AGENTS MUST SCORE 80+
        if score_logic >= 80 and score_gpt >= 80:
            yield f"data: {json.dumps({'type': 'log', 'message': f'✅ CONSENSUS REACHED. All agents agree (80+) after {round_count} rounds!'})}\n\n"
            yield f"data: {json.dumps({'type': 'log', 'message': f'🏁 Final: {avg_score}/100 avg, {current_word_count} body words'})}\n\n"
            break
        else:
            yield f"data: {json.dumps({'type': 'log', 'message': f'❌ NO CONSENSUS. Revising... (Logic: {score_logic}, GPT: {score_gpt})'})}\n\n"
        
        # Stop if max rounds reached
        if round_count >= MAX_ROUNDS:
            yield f"data: {json.dumps({'type': 'log', 'message': f'⏱️ Max {MAX_ROUNDS} rounds reached. Using best draft (avg: {best_avg_score}/100)'})}\n\n"
            current_draft = best_draft
            break
        
        # REVISION PHASE - Prof. Quill MERGES critic contributions
        yield f"data: {json.dumps({'type': 'log', 'message': '✍️ Prof. Quill (Claude 3.5 Sonnet) merging critic contributions...'})}\n\n"
        
        word_count_guidance = ""
        if current_word_count < word_count:
            word_count_guidance = f"\n\nCRITICAL: Current body word count ({current_word_count}) is below target ({word_count}). EXPAND the essay by adding more depth, examples, and analysis."
        elif current_word_count > word_count * 1.15:
            word_count_guidance = f"\n\nNote: Current body word count ({current_word_count}) exceeds target ({word_count}). Tighten the essay by removing redundancy."
        
        merge_prompt = f"""You are Prof. Quill. Revise the essay by MERGING specific contributions from Dean Logic and Chancellor GPT.

CRITICAL TARGET: {target_total_words} words total to ensure {word_count}+ body words after references.
{word_count_guidance}

ORIGINAL INSTRUCTIONS: {instructions}

CURRENT DRAFT: {current_draft}

DEAN LOGIC'S SPECIFIC REWRITES (Score: {score_logic}/100):
{logic_rewrites if logic_rewrites else "No specific rewrites provided (score was 80+)"}

CHANCELLOR GPT'S SPECIFIC REWRITES (Score: {score_gpt}/100):
{gpt_rewrites if gpt_rewrites else "No specific rewrites provided (score was 80+)"}

MERGE REQUIREMENTS:
1. Integrate Dean Logic's specific text blocks into the narrative while maintaining flow
2. Integrate Chancellor GPT's specific text blocks into the narrative while maintaining flow
3. Ensure all new citations are properly formatted
4. Maintain SPARTAN style (direct, authoritative, no fluff)
5. Hit {target_total_words} words total
6. Zero banned words (delve, tapestry, landscape, etc.)

Revise the essay by merging the specific contributions from both critics."""

        try:
            current_draft = call_claude(merge_prompt, max_tokens=8000)
            new_word_count = count_words(current_draft)
            yield f"data: {json.dumps({'type': 'log', 'message': f'✅ Merge complete ({new_word_count} body words)'})}\n\n"
        except Exception as e:
            yield f"data: {json.dumps({'type': 'error', 'message': f'Merge error: {str(e)}'})}\n\n"
            break

    final_word_count = count_words(best_draft)
    final_avg = best_avg_score
    yield f"data: {json.dumps({'type': 'log', 'message': f'🏁 COMPLETE! Final: {final_avg}/100 avg, {final_word_count} body words, {round_count} rounds'})}\n\n"
    
    # Save to database
    try:
        essay = Essay(
            user_id=current_user.id,
            title=instructions[:100],
            instructions=instructions,
            final_content=best_draft,
            final_score=final_avg,
            rounds_used=round_count,
            word_count_limit=word_count
        )
        db.session.add(essay)
        db.session.commit()
        
        yield f"data: {json.dumps({'type': 'complete', 'essay_id': essay.id, 'score': final_avg})}\n\n"
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
    
    print("🎓 ROUND 1: Independent Grading...")
    
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
        quill_response_r1 = call_claude(quill_prompt_r1)
        quill_score_r1 = extract_score(quill_response_r1)
        
        strict_response_r1 = call_gpt4(strict_prompt_r1)
        strict_score_r1 = extract_score(strict_response_r1)
        
        logic_response_r1 = call_gemini(logic_prompt_r1)
        logic_score_r1 = extract_score(logic_response_r1)
        
        print(f"Round 1 Scores - Quill: {quill_score_r1}, Strict: {strict_score_r1}, Logic: {logic_score_r1}")
        
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

        quill_response_r2 = call_claude(quill_prompt_r2)
        quill_score_r2 = extract_score(quill_response_r2)
        
        strict_response_r2 = call_gpt4(strict_prompt_r2)
        strict_score_r2 = extract_score(strict_response_r2)
        
        logic_response_r2 = call_gemini(logic_prompt_r2)
        logic_score_r2 = extract_score(logic_response_r2)
        
        print(f"Round 2 Scores - Quill: {quill_score_r2}, Strict: {strict_score_r2}, Logic: {logic_score_r2}")
        
        print("🎓 ROUND 3: Final Calculation...")
        
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
        
        if password != confirm_password:
            flash('Passwords do not match!', 'error')
            return redirect(url_for('register'))
        
        if User.query.filter_by(username=username).first():
            flash('Username already exists. Please choose a different username.', 'error')
            return redirect(url_for('register'))
        
        if User.query.filter_by(email=email).first():
            flash('Email already registered. Please use a different email or login.', 'error')
            return redirect(url_for('register'))
        
        try:
            user = User(username=username, email=email)
            user.set_password(password)
            db.session.add(user)
            db.session.commit()
            
            email_body = f"""
Welcome to The Academic Board, {username}!

Your account has been successfully created.

Get started with our premium AI-powered academic tools:
- The Architect: AI Essay Writer (CLEAN DEPENDENCY 3-Agent Consensus System!)
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
        except Exception as e:
            db.session.rollback()
            flash('Registration failed. Please try again.', 'error')
            print(f"Registration error: {str(e)}")
            return redirect(url_for('register'))
    
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
        
        if not os.getenv('OPENAI_API_KEY'):
            return jsonify({'reply': 'System Error: Contact Admin - OpenAI API key not configured.'}), 500
        
        system_prompt = """You are The Concierge, the helpful support agent for The Academic Board.

PRICING:
- Essay Writer: £10 per 1,000 words (1,000-12,000 words range)
- Plagiarism Check: £10
- AI Detector: £10
- Assignment Grader: £10

TOOLS:
- The Architect: AI essay writer with CLEAN DEPENDENCY 3-agent consensus system (Prof. Quill, Dean Logic, Chancellor GPT)
- The Detective: Plagiarism checker with PDF reports
- The Oracle: AI content detector
- The Grader: Strict assignment marking

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
    """MEMORY-OPTIMIZED file upload handler for 25MB files."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    
    file = request.files['file']
    if not file.filename or not allowed_file(file.filename):
        return jsonify({'error': 'Invalid file type. Supported: PDF, DOCX, TXT (max 25MB each)'}), 400
    
    filepath = None
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        file.save(filepath)
        text = extract_text_from_file(filepath, filename)
        
        os.remove(filepath)
        filepath = None
        
        if len(text) > 100000:
            text = text[:100000] + "\n\n[Preview truncated. Full content will be used in essay generation.]"
        
        print(f"✅ File processed successfully: {filename}")
        return jsonify({'success': True, 'text': text})
        
    except Exception as e:
        if filepath and os.path.exists(filepath):
            os.remove(filepath)
        print(f"❌ Upload error: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/create-checkout-session', methods=['POST'])
@login_required
def create_checkout_session():
    try:
        data = request.get_json()
        tool_type = data.get('tool_type')
        
        print(f"🔔 Checkout session requested for tool: {tool_type}")
        
        if tool_type not in PRICING:
            return jsonify({'error': 'Invalid tool'}), 400
        
        if tool_type == 'writer':
            word_count = data.get('word_count', 2000)
            amount = int((word_count / 1000) * PRICING['writer']['price_per_1k'])
            product_name = f'Essay Writing ({word_count:,} words)'
            print(f"💰 Writer pricing: {word_count} words = £{amount/100}")
        else:
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
            success_url='https://aiassignment-x631.onrender.com/payment-success?session_id={CHECKOUT_SESSION_ID}&tool=' + tool_type,
            cancel_url='https://aiassignment-x631.onrender.com/dashboard',
            client_reference_id=str(current_user.id),
            customer_email=current_user.email,
        )
        
        print(f"✅ Stripe session created: {checkout_session.id}")
        return jsonify({'id': checkout_session.id})
    except Exception as e:
        print(f"❌ Checkout error: {str(e)}")
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
                payment = Payment(
                    user_id=current_user.id,
                    amount=checkout_session.amount_total,
                    tool_type=tool_type,
                    stripe_session_id=session_id
                )
                db.session.add(payment)
                db.session.commit()
                
                flash('Payment successful!', 'success')
                
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
        headers={
            'Cache-Control': 'no-cache, no-store, must-revalidate',
            'X-Accel-Buffering': 'no'
        }
    )

@app.route('/check-plagiarism', methods=['POST'])
@login_required
def check_plagiarism():
    text = request.form.get('text')
    
    if not text:
        return jsonify({'error': 'Text required'}), 400
    
    try:
        hunter_findings = plagiarism_hunter_gemini(text)
        analyst_findings = plagiarism_analyst_gpt(text, hunter_findings)
        pdf_filename = plagiarism_reporter_claude(text, hunter_findings, analyst_findings, current_user.id)
        
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
    # CRITICAL: Gunicorn timeout MUST be set to 300s+ in Render Start Command: gunicorn --timeout 300 app:app
    app.run(debug=True, host='0.0.0.0', port=5000, threaded=True)